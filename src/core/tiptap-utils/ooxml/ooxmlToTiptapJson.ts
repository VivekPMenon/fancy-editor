import type { JSONContent } from '@tiptap/react';
import { OOXML_NS, parseFlatOpcPackage, type OoxmlPackage } from './flatOpcPackage';
import { parseNumbering, type NumberingResolver } from './numbering';
import {
  parseStyles,
  parseRunProps,
  parseParaProps,
  type StylesResolver,
  type RunProps,
  type ParaProps,
} from './styles';

function tag(namespaceURI: string, localName: string, root: Element): Element[] {
  return Array.from(root.getElementsByTagNameNS(namespaceURI, localName));
}
function firstTag(namespaceURI: string, localName: string, root: Element): Element | undefined {
  return root.getElementsByTagNameNS(namespaceURI, localName)[0];
}
function wAttr(el: Element, name: string): string | null {
  return el.getAttributeNS(OOXML_NS.w, name) ?? el.getAttribute(`w:${name}`);
}
// Deliberately returns plain boolean, not a type predicate — every call
// site here already works with `Element` (from `.children`, not
// `.childNodes`), so narrowing to `Element` in the true branch collapses
// the false branch to `never` (TS reasons "already Element AND not
// Element" is impossible), which breaks normal else-if chains over the
// same variable.
function isEl(node: Element, ns: string, localName: string): boolean {
  return node.namespaceURI === ns && node.localName === localName;
}

const ORDERED_FORMATS = new Set(['decimal', 'upperRoman', 'lowerRoman', 'upperLetter', 'lowerLetter']);

// Word's numFmt values -> the equivalent CSS list-style-type, so an imported
// list renders the same markers (I/II, a/b, …) the document used instead of
// falling back to a generic cycle. Formats with no clean CSS analogue are
// left off and fall back to the default.
const NUMFMT_TO_CSS: Record<string, string> = {
  decimal: 'decimal',
  decimalZero: 'decimal-leading-zero',
  upperRoman: 'upper-roman',
  lowerRoman: 'lower-roman',
  upperLetter: 'upper-alpha',
  lowerLetter: 'lower-alpha',
};

// Internal-only sentinel node type, never actually emitted in the final
// JSON — convertParagraph splits on it (see there) to turn a
// <w:br w:type="column"/> into a real paragraph/heading boundary instead
// of an inline line break, giving CSS column-count natural block-level
// content to flow across.
const COLUMN_BREAK_TYPE = '__columnBreak__';

// Word highlight names -> CSS. Most Word names are also valid CSS color
// keywords (yellow/green/cyan/red/blue/magenta/…); only the "dark*"/gray
// variants need explicit hex.
const HIGHLIGHT_CSS: Record<string, string> = {
  darkYellow: '#808000',
  darkBlue: '#000080',
  darkCyan: '#008080',
  darkGreen: '#008000',
  darkMagenta: '#800080',
  darkRed: '#800000',
  darkGray: '#a9a9a9',
  lightGray: '#d3d3d3',
};
function highlightCss(name: string): string {
  return HIGHLIGHT_CSS[name] ?? name;
}

// Word <w:u w:val> -> CSS text-decoration-style (undefined = a plain solid
// single underline, the default the <u> tag already draws).
function underlineDecorationStyle(val: string): string | undefined {
  if (val === 'double') return 'double';
  if (/dotted/i.test(val)) return 'dotted';
  if (/dash/i.test(val)) return 'dashed';
  if (/wave|wavy/i.test(val)) return 'wavy';
  return undefined;
}

// Builds Tiptap marks from resolved run properties. color+fontSize+smallCaps
// share one textStyle mark (same-type marks can't coexist in ProseMirror);
// dstrike and strike both map to the strike mark, so only one is emitted.
function marksFromRunProps(p: RunProps, linkHref?: string): NonNullable<JSONContent['marks']> {
  const marks: NonNullable<JSONContent['marks']> = [];
  if (p.bold) marks.push({ type: 'bold' });
  if (p.italic) marks.push({ type: 'italic' });
  if (p.underline && p.underline !== 'none') {
    const style = underlineDecorationStyle(p.underline);
    marks.push(style ? { type: 'underline', attrs: { underlineStyle: style } } : { type: 'underline' });
  }
  if (p.dstrike) {
    marks.push({ type: 'strike', attrs: { strikeStyle: 'double' } });
  } else if (p.strike) {
    marks.push({ type: 'strike' });
  }
  const ts: Record<string, string> = {};
  if (p.color) ts.color = `#${p.color}`;
  if (p.fontSizePt) ts.fontSize = p.fontSizePt;
  if (p.smallCaps) ts.fontVariant = 'small-caps';
  if (Object.keys(ts).length) marks.push({ type: 'textStyle', attrs: ts });
  if (p.highlight) marks.push({ type: 'highlight', attrs: { color: highlightCss(p.highlight) } });
  if (p.vertAlign === 'subscript') marks.push({ type: 'subscript' });
  if (p.vertAlign === 'superscript') marks.push({ type: 'superscript' });
  if (linkHref) marks.push({ type: 'link', attrs: { href: linkHref } });
  return marks;
}

// Converts one <w:r> (a run) into zero or more inline JSONContent nodes —
// usually one text node, but a run can contain a line break (<w:br/>) or a
// drawing (image), and <w:t> could in principle be absent (formatting-only
// runs produce nothing visible).
function convertRun(runEl: Element, baseProps: RunProps, linkHref: string | undefined, ctx: ConvertContext): JSONContent[] {
  const rPr = firstTag(OOXML_NS.w, 'rPr', runEl);
  // Run's own properties override whatever the paragraph/style seeded.
  const effective: RunProps = { ...baseProps, ...parseRunProps(rPr) };
  const nodes: JSONContent[] = [];

  for (const child of Array.from(runEl.children)) {
    if (isEl(child, OOXML_NS.w, 't')) {
      const text = child.textContent ?? '';
      if (text) {
        const m = marksFromRunProps(effective, linkHref);
        nodes.push(m.length ? { type: 'text', text, marks: m } : { type: 'text', text });
      }
    } else if (isEl(child, OOXML_NS.w, 'br')) {
      const brType = wAttr(child, 'type');
      nodes.push({ type: brType === 'column' ? COLUMN_BREAK_TYPE : 'hardBreak' });
    }
    // <w:drawing> (images) is intentionally not handled in this loop — see
    // the dedicated pass right below, since images need to become
    // top-level sibling nodes, not inline content.
  }

  // <w:drawing> isn't inline content in Tiptap's schema (Image is block-
  // level) — collect any image(s) in this run separately so the caller can
  // splice them in as their own top-level nodes rather than inline text.
  for (const drawingEl of tag(OOXML_NS.w, 'drawing', runEl)) {
    const image = convertDrawing(drawingEl, ctx);
    if (image) {
      ctx.pendingImages.push(image);
    }
  }

  return nodes;
}

// <w:hyperlink> wraps one or more <w:r> runs and applies a link mark to
// all of them — r:id resolves through document.xml's own .rels
// (hyperlinks external targets are stored there, same file as image
// relationships). w:anchor (internal bookmark links) isn't resolved to
// anything Tiptap's link mark can represent, so those runs come through
// as plain unlinked text — a real, documented gap, not silently wrong.
function convertParagraphInlineContent(pEl: Element, baseRunProps: RunProps, ctx: ConvertContext): JSONContent[] {
  const nodes: JSONContent[] = [];
  for (const child of Array.from(pEl.children)) {
    if (isEl(child, OOXML_NS.w, 'r')) {
      nodes.push(...convertRun(child, baseRunProps, undefined, ctx));
    } else if (isEl(child, OOXML_NS.w, 'hyperlink')) {
      const rId = child.getAttributeNS(OOXML_NS.r, 'id') ?? child.getAttribute('r:id');
      const href = rId ? ctx.pkg.resolveRelationship('/word/document.xml', rId) : undefined;
      for (const runEl of tag(OOXML_NS.w, 'r', child)) {
        nodes.push(...convertRun(runEl, baseRunProps, href, ctx));
      }
    }
  }
  return nodes;
}

interface ListMembership {
  numId: string;
  ilvl: number;
  ordered: boolean;
  // CSS list-style-type for this level's numFmt (ordered lists only).
  listStyleType?: string;
}

function paragraphStyleId(pPr: Element | undefined): string | null {
  const pStyleEl = pPr && firstTag(OOXML_NS.w, 'pStyle', pPr);
  return pStyleEl ? wAttr(pStyleEl, 'val') : null;
}

// Effective paragraph properties = the style's resolved pPr (basedOn chain +
// docDefaults) with the paragraph's own direct pPr layered on top. Crucially
// this is where numbering that lives on a *style* (ListBullet/ListNumber,
// whose items carry no direct <w:numPr>) becomes visible.
function effectiveParaProps(pPr: Element | undefined, styleId: string | null, styles: StylesResolver): ParaProps {
  const stylePara = styleId ? styles.styleProps(styleId).para : styles.defaults.para;
  return { ...stylePara, ...parseParaProps(pPr) };
}

// numId "0" is Word's "explicitly no numbering" sentinel — treat as not a list.
function membershipFromPara(para: ParaProps, numbering: NumberingResolver): ListMembership | undefined {
  if (!para.numId || para.numId === '0') {
    return undefined;
  }
  const ilvl = para.ilvl ?? 0;
  const format = numbering.getLevelFormat(para.numId, ilvl);
  const ordered = format ? ORDERED_FORMATS.has(format) : true;
  return {
    numId: para.numId,
    ilvl,
    ordered,
    listStyleType: ordered && format ? NUMFMT_TO_CSS[format] : undefined,
  };
}

// Twips (1/1440 inch) -> px string. 1440 twips/in, 96px/in => px = twips/15.
function twipsToPx(twips: number | undefined): string | undefined {
  if (!twips) {
    return undefined;
  }
  return `${Math.round(twips / 15)}px`;
}

// <w:spacing line/lineRule> -> CSS line-height. "auto" is 240ths of a single
// line (a unitless multiple); "exact"/"atLeast" are twips (an absolute px).
function lineHeightFrom(para: ParaProps): string | undefined {
  if (para.lineValue === undefined) {
    return undefined;
  }
  if (!para.lineRule || para.lineRule === 'auto') {
    const mult = para.lineValue / 240;
    return `${Number.isInteger(mult) ? mult : Number(mult.toFixed(2))}`;
  }
  return `${Math.round(para.lineValue / 15)}px`;
}

// px string that keeps an explicit 0 (so e.g. No Spacing's 0-after becomes a
// real margin-bottom:0, overriding the CSS baseline).
function twipsToPxWithZero(twips: number): string {
  return `${Math.round(twips / 15)}px`;
}

// Turns resolved paragraph properties into Tiptap node attrs. Headings get
// only alignment (their size/spacing stay CSS-driven); body paragraphs also
// get indentation, spacing, and shading. Spacing (line-height, space
// before/after) is only emitted when it *differs* from docDefaults — the
// stylesheet baseline (App.css / PostFeedTab.css) already covers the default,
// so emitting it on every paragraph would be noise and would fight that CSS
// (e.g. re-loosening tight list items).
function paragraphAttrs(para: ParaProps, isHeading: boolean, defaults: ParaProps): Record<string, unknown> {
  const attrs: Record<string, unknown> = {};
  const align =
    para.jc === 'both' ? 'justify' : para.jc === 'center' || para.jc === 'right' ? para.jc : undefined;
  if (align) attrs.textAlign = align;
  if (isHeading) {
    return attrs;
  }
  const left = twipsToPx(para.indLeftTwips);
  if (left) attrs.marginLeft = left;
  const right = twipsToPx(para.indRightTwips);
  if (right) attrs.marginRight = right;
  if (para.indFirstLineTwips) {
    attrs.textIndent = twipsToPx(para.indFirstLineTwips);
  } else if (para.indHangingTwips) {
    attrs.textIndent = `-${twipsToPx(para.indHangingTwips)}`;
  }
  if (para.lineValue !== undefined && (para.lineValue !== defaults.lineValue || para.lineRule !== defaults.lineRule)) {
    const lineHeight = lineHeightFrom(para);
    if (lineHeight) attrs.lineHeight = lineHeight;
  }
  if (para.spaceBeforeTwips !== undefined && para.spaceBeforeTwips !== defaults.spaceBeforeTwips) {
    attrs.marginTop = twipsToPxWithZero(para.spaceBeforeTwips);
  }
  if (para.spaceAfterTwips !== undefined && para.spaceAfterTwips !== defaults.spaceAfterTwips) {
    attrs.marginBottom = twipsToPxWithZero(para.spaceAfterTwips);
  }
  if (para.shadingFill) attrs.shading = `#${para.shadingFill}`;
  return attrs;
}

interface ConvertContext {
  pkg: OoxmlPackage;
  numbering: NumberingResolver;
  styles: StylesResolver;
  // Images are collected here as they're found inside a paragraph's runs,
  // then spliced in as top-level sibling nodes after the paragraph is
  // built — Image is block-level in our schema, so it can't stay nested
  // inside the paragraph's own inline content array.
  pendingImages: JSONContent[];
  // The *owning paragraph's* <w:jc> (center/right), threaded down for
  // convertDrawing to use on inline images — Word expresses "this
  // standalone image is centered" as alignment on the wrapping paragraph,
  // not the image, same as the HTML/RTF import path's §11 finding. Since a
  // paragraph whose sole content is an image is dropped entirely (no
  // redundant empty paragraph next to it — see the body-walking loop),
  // that alignment has to be captured into the image itself *before* the
  // paragraph disappears, not recovered after the fact. Only meaningful
  // for inline images; anchored ones compute their own align from the
  // anchor's own position/wrap data (inferFloatFromAnchor), which takes
  // priority.
  currentParagraphJc?: 'center' | 'right';
}

// Returns an array, not a single node — usually length 1, but a
// <w:br w:type="column"/> inside this paragraph (see convertRun) splits it
// into multiple paragraph/heading nodes, one per column-break segment.
function convertParagraph(pEl: Element, ctx: ConvertContext): JSONContent[] {
  const pPr = firstTag(OOXML_NS.w, 'pPr', pEl);
  const styleId = paragraphStyleId(pPr);
  const headingLevel = styleId ? ctx.styles.resolveHeadingLevel(styleId) : undefined;
  const styleProps = styleId ? ctx.styles.styleProps(styleId) : ctx.styles.defaults;
  const effPara = { ...styleProps.para, ...parseParaProps(pPr) };

  // Run baseline for this paragraph: the style's resolved run props, refined
  // by the paragraph-mark rPr (pPr/rPr). Headings drop the size so their
  // heading-level + CSS sizing wins (unchanged behavior); other run props
  // (e.g. a heading style's color) still flow through.
  const pMarkRpr = pPr ? Array.from(pPr.children).find((c) => isEl(c, OOXML_NS.w, 'rPr')) : undefined;
  let baseRun: RunProps = { ...styleProps.run, ...parseRunProps(pMarkRpr) };
  if (headingLevel) {
    baseRun = { ...baseRun, fontSizePt: undefined };
  }

  ctx.currentParagraphJc = effPara.jc === 'center' || effPara.jc === 'right' ? effPara.jc : undefined;
  const content = convertParagraphInlineContent(pEl, baseRun, ctx);
  ctx.currentParagraphJc = undefined;

  // Table-of-Contents lines (TOC1/TOC2/… styles) render as a dedicated node
  // that lays out entry text, a dotted leader, and the page number — see
  // tocEntryExtension. The trailing numeric run is the page number.
  const tocMatch = styleId ? /^TOC(\d+)$/.exec(styleId) : null;
  if (tocMatch && content.length) {
    const level = Number(tocMatch[1]) || 1;
    let page = '';
    const inline = [...content];
    const last = inline[inline.length - 1];
    if (last && last.type === 'text' && /^\d+$/.test((last.text ?? '').trim())) {
      page = (last.text ?? '').trim();
      inline.pop();
    }
    return [{ type: 'tocEntry', attrs: { level, page }, content: inline }];
  }

  const segments: JSONContent[][] = [[]];
  for (const node of content) {
    if (node.type === COLUMN_BREAK_TYPE) {
      segments.push([]);
    } else {
      segments[segments.length - 1].push(node);
    }
  }

  const attrs = paragraphAttrs(effPara, Boolean(headingLevel), ctx.styles.defaults.para);
  return segments.map((seg) => {
    if (headingLevel) {
      return {
        type: 'heading',
        attrs: { level: headingLevel, ...attrs },
        ...(seg.length ? { content: seg } : {}),
      };
    }
    return {
      type: 'paragraph',
      ...(Object.keys(attrs).length ? { attrs } : {}),
      ...(seg.length ? { content: seg } : {}),
    };
  });
}

// wp:positionH doesn't carry a plain "left"/"right" the way HTML's legacy
// align attribute does — it's either an absolute offset (<wp:posOffset>)
// or a named <wp:align>. Heuristic, not exact: prefer an explicit
// <wp:align>; otherwise infer from wrapText (wrapSquare's "left" means
// text wraps on the image's left, i.e. the image itself sits on the
// *right*, and vice versa); default to float-left if neither gives a
// clear signal. Verified against sample-1.xml's one anchored image
// (posOffset 0, wrapText="bothSides" -> falls to the default, and visually
// that image does sit at the left margin there) but this is a narrower
// signal than what real documents may use (percentage-based positioning,
// wrapTight/wrapThrough, etc. aren't handled).
function inferFloatFromAnchor(anchorEl: Element): 'float-left' | 'float-right' {
  const positionH = firstTag(OOXML_NS.wp, 'positionH', anchorEl);
  const alignEl = positionH && firstTag(OOXML_NS.wp, 'align', positionH);
  const align = alignEl?.textContent?.trim();
  if (align === 'right') return 'float-right';
  if (align === 'left') return 'float-left';

  const wrapSquare = firstTag(OOXML_NS.wp, 'wrapSquare', anchorEl);
  const wrapText = wrapSquare?.getAttribute('wrapText');
  if (wrapText === 'left') return 'float-right';
  if (wrapText === 'right') return 'float-left';

  return 'float-left';
}

// An Insert -> Online Video: <wp15:webVideoPr embeddedHtml="<iframe ... src=
// 'https://www.youtube.com/embed/ID?...' ...>"> carries the real embed. The
// embeddedHtml attribute value is XML-decoded by the parser, so it's a plain
// iframe string we can pull the src out of — far better than the paste path's
// pattern-match, since the URL is explicit. We emit the same `youtube` node
// the paste path does (via the Youtube extension), so both render identically.
// Non-YouTube players fall through to the poster image (best we can do).
function youtubeNodeFromWebVideo(
  drawingEl: Element,
  width: number | undefined,
  height: number | undefined,
): JSONContent | undefined {
  const webVideoEl = firstTag(OOXML_NS.wp15, 'webVideoPr', drawingEl);
  const embeddedHtml = webVideoEl?.getAttribute('embeddedHtml') ?? '';
  const src = /src="([^"]+)"/i.exec(embeddedHtml)?.[1];
  const youtubeId = src ? /(?:youtube(?:-nocookie)?\.com)\/embed\/([\w-]+)/i.exec(src)?.[1] : undefined;
  if (!youtubeId) {
    return undefined;
  }
  return {
    type: 'youtube',
    attrs: {
      src: `https://www.youtube.com/embed/${youtubeId}`,
      ...(width ? { width } : {}),
      ...(height ? { height } : {}),
    },
  };
}

function convertDrawing(drawingEl: Element, ctx: ConvertContext): JSONContent | undefined {
  const inlineEl = firstTag(OOXML_NS.wp, 'inline', drawingEl);
  const anchorEl = firstTag(OOXML_NS.wp, 'anchor', drawingEl);
  const container = inlineEl ?? anchorEl;
  if (!container) {
    return undefined;
  }

  const extentEl = firstTag(OOXML_NS.wp, 'extent', container);
  // EMUs (English Metric Units) -> px, 914400 EMU per inch, 96px per inch.
  const width = extentEl ? Math.round((Number(extentEl.getAttribute('cx')) / 914400) * 96) : undefined;
  const height = extentEl ? Math.round((Number(extentEl.getAttribute('cy')) / 914400) * 96) : undefined;

  // Online video takes priority over the poster image it also carries.
  const youtube = youtubeNodeFromWebVideo(drawingEl, width, height);
  if (youtube) {
    return youtube;
  }

  const blipEl = firstTag(OOXML_NS.a, 'blip', container);
  const rId = blipEl ? (blipEl.getAttributeNS(OOXML_NS.r, 'embed') ?? blipEl.getAttribute('r:embed')) : null;
  const mediaPartName = rId ? ctx.pkg.resolveRelationship('/word/document.xml', rId) : undefined;
  const mediaPart = mediaPartName ? ctx.pkg.getPart(mediaPartName) : undefined;
  if (!mediaPart?.base64) {
    return undefined;
  }
  const mimeType = mediaPart.contentType || guessMimeFromPartName(mediaPart.name);
  const src = `data:${mimeType};base64,${mediaPart.base64}`;

  const docPrEl = firstTag(OOXML_NS.wp, 'docPr', container);
  const alt = docPrEl?.getAttribute('descr') ?? docPrEl?.getAttribute('name') ?? undefined;

  const align = anchorEl ? inferFloatFromAnchor(anchorEl) : ctx.currentParagraphJc;

  return {
    type: 'image',
    attrs: {
      src,
      ...(alt ? { alt } : {}),
      ...(width ? { width } : {}),
      ...(height ? { height } : {}),
      ...(align ? { align } : {}),
    },
  };
}

function guessMimeFromPartName(name: string): string {
  if (name.endsWith('.png')) return 'image/png';
  if (name.endsWith('.jpg') || name.endsWith('.jpeg')) return 'image/jpeg';
  if (name.endsWith('.gif')) return 'image/gif';
  if (name.endsWith('.svg')) return 'image/svg+xml';
  return 'application/octet-stream';
}

function convertTable(tblEl: Element, ctx: ConvertContext): JSONContent {
  const rows: JSONContent[] = [];
  for (const trEl of tag(OOXML_NS.w, 'tr', tblEl)) {
    const cells: JSONContent[] = [];
    for (const tcEl of tag(OOXML_NS.w, 'tc', trEl)) {
      const tcPr = firstTag(OOXML_NS.w, 'tcPr', tcEl);
      const gridSpanEl = tcPr && firstTag(OOXML_NS.w, 'gridSpan', tcPr);
      const colspan = gridSpanEl ? Number(wAttr(gridSpanEl, 'val') ?? '1') : 1;
      // Cell shading (<w:shd w:fill>) — dropping it not only lost the banner
      // color but hid white-on-fill header text entirely (white text on the
      // default white cell). "auto" means no fill.
      const shdEl = tcPr && firstTag(OOXML_NS.w, 'shd', tcPr);
      const fill = shdEl ? wAttr(shdEl, 'fill') : null;
      const backgroundColor = fill && fill !== 'auto' ? `#${fill}` : undefined;
      const cellContent = tag(OOXML_NS.w, 'p', tcEl)
        .filter((p) => p.parentElement === tcEl) // direct children only — a nested table's <w:p>s shouldn't merge in
        .flatMap((p) => convertParagraph(p, ctx)) // usually 1 node per <w:p>, more if a column break split it
        .map((p) => (p.content ? p : { ...p, content: [] })); // paragraph node must have `content` even if empty for table cells to render sensibly
      const cellAttrs = {
        ...(colspan > 1 ? { colspan } : {}),
        ...(backgroundColor ? { backgroundColor } : {}),
      };
      cells.push({
        type: 'tableCell',
        ...(Object.keys(cellAttrs).length ? { attrs: cellAttrs } : {}),
        content: cellContent.length ? cellContent : [{ type: 'paragraph' }],
      });
    }
    rows.push({ type: 'tableRow', content: cells });
  }
  return { type: 'table', content: rows };
}

// Groups a flat run of consecutive list-membership paragraphs into proper
// nested bulletList/orderedList + listItem structure, using a stack keyed
// by ilvl — same shape of algorithm as wordListReconstruction.ts's
// buildListTree, but driven by real numId/ilvl/format instead of guessing
// from indentation.
// Finds the font-size of the first text run in a (list-item) paragraph — the
// size to hoist onto the <li> so its marker matches the item text.
function firstTextStyleFontSize(node: JSONContent): string | undefined {
  if (node.type === 'text') {
    const textStyle = node.marks?.find((mark) => mark.type === 'textStyle');
    const size = textStyle?.attrs?.fontSize;
    return typeof size === 'string' ? size : undefined;
  }
  for (const child of node.content ?? []) {
    const found = firstTextStyleFontSize(child);
    if (found) {
      return found;
    }
  }
  return undefined;
}

function buildListNode(paragraphs: { node: JSONContent; membership: ListMembership }[]): JSONContent {
  // Type/style are corrected from the first item actually placed in each list
  // (below) — Word can switch bullet<->number or change numFmt between
  // levels, and a list's format is that of the items it *directly* holds, not
  // the parent paragraph that spawned it.
  const root: JSONContent = { type: 'orderedList', content: [] };
  const stack: { ilvl: number; list: JSONContent }[] = [{ ilvl: -1, list: root }];

  for (const { node, membership } of paragraphs) {
    while (stack.length > 1 && stack[stack.length - 1].ilvl >= membership.ilvl) {
      stack.pop();
    }
    const parentList = stack[stack.length - 1].list;
    if (!parentList.content || parentList.content.length === 0) {
      parentList.type = membership.ordered ? 'orderedList' : 'bulletList';
      if (membership.listStyleType) {
        parentList.attrs = { ...(parentList.attrs ?? {}), listStyleType: membership.listStyleType };
      }
    }
    const listItem: JSONContent = { type: 'listItem', content: [node] };
    // Put the item's own font-size on the <li> too, not just its inner text —
    // the list marker (1., I., •) inherits its size from the <li> element, so
    // without this the number/bullet renders at the editor's base size while
    // the text sits at the document's, and the two visibly mismatch. Same
    // reasoning the paste path already applies in wordListReconstruction.
    const itemFontSize = firstTextStyleFontSize(node);
    if (itemFontSize) {
      listItem.attrs = { ...(listItem.attrs ?? {}), fontSize: itemFontSize };
    }
    (parentList.content ??= []).push(listItem);

    const childList: JSONContent = { type: 'orderedList', content: [] };
    // Lazily attach — only actually appended to the item if a deeper-level
    // paragraph shows up next; an empty trailing list looks wrong.
    stack.push({ ilvl: membership.ilvl, list: childList });
    (listItem as { _pendingChildList?: JSONContent })._pendingChildList = childList;
    (listItem as { _pendingChildListParent?: JSONContent })._pendingChildListParent = listItem;
  }

  // Second pass: attach any child list that actually received content.
  attachPendingChildLists(root);
  return root;
}

function attachPendingChildLists(node: JSONContent): void {
  if (!node.content) {
    return;
  }
  for (const child of node.content) {
    const pendingList = (child as { _pendingChildList?: JSONContent })._pendingChildList;
    if (pendingList && pendingList.content && pendingList.content.length > 0) {
      (child.content ??= []).push(pendingList);
    }
    delete (child as { _pendingChildList?: JSONContent })._pendingChildList;
    delete (child as { _pendingChildListParent?: JSONContent })._pendingChildListParent;
    attachPendingChildLists(child);
  }
}

// <w:cols w:num="N"> on a section's properties — no attribute at all (or
// num<=1) means the ordinary single-column case.
function sectionColumnCount(sectPrEl: Element): number {
  const colsEl = firstTag(OOXML_NS.w, 'cols', sectPrEl);
  const numStr = colsEl ? wAttr(colsEl, 'num') : null;
  const num = numStr ? Number(numStr) : 1;
  return Number.isFinite(num) && num > 1 ? num : 1;
}

export function convertOoxmlToTiptapJson(pkg: OoxmlPackage): JSONContent {
  const docRoot = pkg.requireXml('/word/document.xml');
  const bodyEl = firstTag(OOXML_NS.w, 'body', docRoot);
  if (!bodyEl) {
    throw new Error('OOXML document.xml has no <w:body>');
  }
  const numberingRoot = pkg.getPart('/word/numbering.xml')?.root;
  const numbering = numberingRoot ? parseNumbering(numberingRoot) : { getLevelFormat: () => undefined };
  const stylesRoot = pkg.getPart('/word/styles.xml')?.root;
  const emptyStyles: StylesResolver = {
    resolveHeadingLevel: () => undefined,
    styleProps: () => ({ para: {}, run: {} }),
    defaults: { para: {}, run: {} },
  };
  const styles = stylesRoot ? parseStyles(stylesRoot) : emptyStyles;
  const ctx: ConvertContext = { pkg, numbering, styles, pendingImages: [] };

  const blocks: JSONContent[] = [];
  // A <w:sectPr> — either embedded in a paragraph's <w:pPr> (marking where
  // that section ends) or as body's own final child (the document's last
  // section) — describes the section of content that PRECEDES it, back to
  // the previous section break or the start of the document. Blocks
  // accumulate here per section, not directly into `blocks`, so a
  // multi-column section can be wrapped in one columnLayout node once its
  // boundary is found; an ordinary single-column section just gets
  // flattened straight into `blocks` as before.
  let sectionBlocks: JSONContent[] = [];
  let currentListRun: { node: JSONContent; membership: ListMembership }[] = [];

  function flushListRun() {
    if (currentListRun.length > 0) {
      sectionBlocks.push(buildListNode(currentListRun));
      currentListRun = [];
    }
  }

  function flushSection(columns: number) {
    flushListRun();
    if (sectionBlocks.length === 0) {
      return;
    }
    if (columns > 1) {
      blocks.push({ type: 'columnLayout', attrs: { columns }, content: sectionBlocks });
    } else {
      blocks.push(...sectionBlocks);
    }
    sectionBlocks = [];
  }

  for (const child of Array.from(bodyEl.children)) {
    if (isEl(child, OOXML_NS.w, 'p')) {
      const pPr = firstTag(OOXML_NS.w, 'pPr', child);
      // List membership is read from the *effective* paragraph props, so items
      // whose numbering lives on their style (ListBullet/ListNumber) — with no
      // direct <w:numPr> — are still recognized as list items.
      const membership = membershipFromPara(effectiveParaProps(pPr, paragraphStyleId(pPr), styles), numbering);
      ctx.pendingImages = [];
      const nodes = convertParagraph(child, ctx);
      const hasImages = ctx.pendingImages.length > 0;
      // A paragraph carrying a floating image commonly carries the text
      // that should wrap around it too, in the SAME <w:p> — Word doesn't
      // care which run comes first internally, but CSS float only pulls
      // in content that comes *after* it in document order. Images must
      // go before the paragraph's own text, not after (confirmed against
      // a real capture where getting this backwards put the wrap-text
      // paragraph entirely above the image instead of beside it — floats
      // don't affect earlier siblings at all). A paragraph whose only
      // content WAS the image(s) — no other text — is dropped rather than
      // also emitted as a redundant empty paragraph right next to them.

      const sectPrEl = pPr && firstTag(OOXML_NS.w, 'sectPr', pPr);
      // A paragraph whose ONLY pPr content is a section-break marker — no
      // text, no images, not a column-break split — is pure structure, not
      // visible content, and shouldn't render as a stray blank paragraph.
      const isBareSectionMarker = Boolean(sectPrEl) && nodes.length === 1 && !nodes[0].content && !hasImages;

      if (membership) {
        sectionBlocks.push(...ctx.pendingImages);
        // A column break inside a list-item paragraph isn't a realistic
        // combination in practice — only the first resulting segment
        // becomes the list item's content if one somehow occurred.
        currentListRun.push({ node: nodes[0], membership });
      } else {
        flushListRun();
        sectionBlocks.push(...ctx.pendingImages);
        // A paragraph whose only content is a <w:sectPr> (section break
        // marker) or is otherwise genuinely empty is skipped rather than
        // emitted as a stray blank paragraph, matching how Word itself
        // treats it as a break, not visible content — UNLESS its only
        // content was the image(s) just pushed above (covered by
        // hasImages/nodes[0].content, same check as before), or a column
        // break split it into more than one segment, in which case all of
        // them are kept rather than arbitrarily dropping one.
        if (!isBareSectionMarker && (nodes.length > 1 || !hasImages || nodes[0].content)) {
          sectionBlocks.push(...nodes);
        }
      }

      if (sectPrEl) {
        flushSection(sectionColumnCount(sectPrEl));
      }
    } else if (isEl(child, OOXML_NS.w, 'tbl')) {
      flushListRun();
      sectionBlocks.push(convertTable(child, ctx));
    } else if (isEl(child, OOXML_NS.w, 'sectPr')) {
      // The document's final section properties, as a direct (non-
      // paragraph) body child — describes whatever accumulated since the
      // last section break, or the whole document if there was none.
      flushSection(sectionColumnCount(child));
    }
  }
  flushSection(1); // safety net in case the document has no trailing sectPr at all (shouldn't happen per spec, but don't silently drop content if it does)

  return { type: 'doc', content: blocks };
}

/**
 * Convenience entry point taking the raw flat-OPC XML string straight from
 * Body.getOoxml() — unpacks the package and converts in one call, for
 * callers that don't need the intermediate OoxmlPackage themselves.
 */
export function convertOoxmlStringToTiptapJson(flatOpcXml: string): JSONContent {
  return convertOoxmlToTiptapJson(parseFlatOpcPackage(flatOpcXml));
}
