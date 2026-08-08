import type { JSONContent } from '@tiptap/react';
import { OOXML_NS, parseFlatOpcPackage, type OoxmlPackage } from './flatOpcPackage';
import { parseNumbering, type NumberingResolver } from './numbering';
import { parseStyles, type StylesResolver } from './styles';

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

// Internal-only sentinel node type, never actually emitted in the final
// JSON — convertParagraph splits on it (see there) to turn a
// <w:br w:type="column"/> into a real paragraph/heading boundary instead
// of an inline line break, giving CSS column-count natural block-level
// content to flow across.
const COLUMN_BREAK_TYPE = '__columnBreak__';

interface RunMarks {
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
  strike?: boolean;
  color?: string;
  highlight?: string;
  vertAlign?: 'subscript' | 'superscript';
  linkHref?: string;
}

function marksFromRunMarks(m: RunMarks): NonNullable<JSONContent['marks']> {
  const marks: NonNullable<JSONContent['marks']> = [];
  if (m.bold) marks.push({ type: 'bold' });
  if (m.italic) marks.push({ type: 'italic' });
  if (m.underline) marks.push({ type: 'underline' });
  if (m.strike) marks.push({ type: 'strike' });
  if (m.color) marks.push({ type: 'textStyle', attrs: { color: m.color } });
  if (m.highlight) marks.push({ type: 'highlight', attrs: { color: m.highlight } });
  if (m.vertAlign === 'subscript') marks.push({ type: 'subscript' });
  if (m.vertAlign === 'superscript') marks.push({ type: 'superscript' });
  if (m.linkHref) marks.push({ type: 'link', attrs: { href: m.linkHref } });
  return marks;
}

// <w:color w:val="C00000"/> is already a hex triplet (or "auto") — just
// needs the leading '#' CSS expects.
function resolveColorAttr(rPr: Element): string | undefined {
  const colorEl = firstTag(OOXML_NS.w, 'color', rPr);
  const val = colorEl ? wAttr(colorEl, 'val') : null;
  if (!val || val === 'auto') {
    return undefined;
  }
  return `#${val}`;
}

function runMarksFromRPr(rPr: Element | undefined, inherited: RunMarks): RunMarks {
  if (!rPr) {
    return inherited;
  }
  const vertAlignEl = firstTag(OOXML_NS.w, 'vertAlign', rPr);
  const vertAlignVal = vertAlignEl ? wAttr(vertAlignEl, 'val') : null;
  return {
    ...inherited,
    bold: firstTag(OOXML_NS.w, 'b', rPr) ? wAttr(firstTag(OOXML_NS.w, 'b', rPr)!, 'val') !== '0' : inherited.bold,
    italic: firstTag(OOXML_NS.w, 'i', rPr) ? wAttr(firstTag(OOXML_NS.w, 'i', rPr)!, 'val') !== '0' : inherited.italic,
    underline: firstTag(OOXML_NS.w, 'u', rPr) ? wAttr(firstTag(OOXML_NS.w, 'u', rPr)!, 'val') !== 'none' : inherited.underline,
    strike: firstTag(OOXML_NS.w, 'strike', rPr) ? wAttr(firstTag(OOXML_NS.w, 'strike', rPr)!, 'val') !== '0' : inherited.strike,
    color: resolveColorAttr(rPr) ?? inherited.color,
    vertAlign:
      vertAlignVal === 'subscript' || vertAlignVal === 'superscript' ? vertAlignVal : inherited.vertAlign,
  };
}

// Converts one <w:r> (a run) into zero or more inline JSONContent nodes —
// usually one text node, but a run can contain a line break (<w:br/>) or a
// drawing (image), and <w:t> could in principle be absent (formatting-only
// runs produce nothing visible).
function convertRun(runEl: Element, inheritedMarks: RunMarks, ctx: ConvertContext): JSONContent[] {
  const rPr = firstTag(OOXML_NS.w, 'rPr', runEl);
  const marks = runMarksFromRPr(rPr, inheritedMarks);
  const nodes: JSONContent[] = [];

  for (const child of Array.from(runEl.children)) {
    if (isEl(child, OOXML_NS.w, 't')) {
      const text = child.textContent ?? '';
      if (text) {
        const m = marksFromRunMarks(marks);
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
function convertParagraphInlineContent(pEl: Element, ctx: ConvertContext): JSONContent[] {
  const nodes: JSONContent[] = [];
  for (const child of Array.from(pEl.children)) {
    if (isEl(child, OOXML_NS.w, 'r')) {
      nodes.push(...convertRun(child, {}, ctx));
    } else if (isEl(child, OOXML_NS.w, 'hyperlink')) {
      const rId = child.getAttributeNS(OOXML_NS.r, 'id') ?? child.getAttribute('r:id');
      const href = rId ? ctx.pkg.resolveRelationship('/word/document.xml', rId) : undefined;
      for (const runEl of tag(OOXML_NS.w, 'r', child)) {
        nodes.push(...convertRun(runEl, href ? { linkHref: href } : {}, ctx));
      }
    }
  }
  return nodes;
}

interface ListMembership {
  numId: string;
  ilvl: number;
  ordered: boolean;
}

function listMembershipFor(pPr: Element | undefined, numbering: NumberingResolver): ListMembership | undefined {
  const numPr = pPr && firstTag(OOXML_NS.w, 'numPr', pPr);
  if (!numPr) {
    return undefined;
  }
  const numIdEl = firstTag(OOXML_NS.w, 'numId', numPr);
  const ilvlEl = firstTag(OOXML_NS.w, 'ilvl', numPr);
  const numId = numIdEl ? wAttr(numIdEl, 'val') : null;
  const ilvl = ilvlEl ? Number(wAttr(ilvlEl, 'val') ?? '0') : 0;
  if (!numId) {
    return undefined;
  }
  const format = numbering.getLevelFormat(numId, ilvl);
  return { numId, ilvl, ordered: format ? ORDERED_FORMATS.has(format) : true };
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
  const pStyleEl = pPr && firstTag(OOXML_NS.w, 'pStyle', pPr);
  const styleId = pStyleEl ? wAttr(pStyleEl, 'val') : null;
  const headingLevel = styleId ? ctx.styles.resolveHeadingLevel(styleId) : undefined;

  const jcEl = pPr && firstTag(OOXML_NS.w, 'jc', pPr);
  const jc = jcEl ? wAttr(jcEl, 'val') : null;
  const textAlign = jc === 'center' || jc === 'right' || jc === 'both' ? (jc === 'both' ? 'justify' : jc) : undefined;

  ctx.currentParagraphJc = jc === 'center' || jc === 'right' ? jc : undefined;
  const content = convertParagraphInlineContent(pEl, ctx);
  ctx.currentParagraphJc = undefined;

  const segments: JSONContent[][] = [[]];
  for (const node of content) {
    if (node.type === COLUMN_BREAK_TYPE) {
      segments.push([]);
    } else {
      segments[segments.length - 1].push(node);
    }
  }

  return segments.map((seg) => {
    if (headingLevel) {
      return {
        type: 'heading',
        attrs: { level: headingLevel, ...(textAlign ? { textAlign } : {}) },
        ...(seg.length ? { content: seg } : {}),
      };
    }
    return {
      type: 'paragraph',
      ...(textAlign ? { attrs: { textAlign } } : {}),
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

function convertDrawing(drawingEl: Element, ctx: ConvertContext): JSONContent | undefined {
  const inlineEl = firstTag(OOXML_NS.wp, 'inline', drawingEl);
  const anchorEl = firstTag(OOXML_NS.wp, 'anchor', drawingEl);
  const container = inlineEl ?? anchorEl;
  if (!container) {
    return undefined;
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

  const extentEl = firstTag(OOXML_NS.wp, 'extent', container);
  // EMUs (English Metric Units) -> px, 914400 EMU per inch, 96px per inch.
  const width = extentEl ? Math.round(Number(extentEl.getAttribute('cx')) / 914400 * 96) : undefined;
  const height = extentEl ? Math.round(Number(extentEl.getAttribute('cy')) / 914400 * 96) : undefined;

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
      const cellContent = tag(OOXML_NS.w, 'p', tcEl)
        .filter((p) => p.parentElement === tcEl) // direct children only — a nested table's <w:p>s shouldn't merge in
        .flatMap((p) => convertParagraph(p, ctx)) // usually 1 node per <w:p>, more if a column break split it
        .map((p) => (p.content ? p : { ...p, content: [] })); // paragraph node must have `content` even if empty for table cells to render sensibly
      cells.push({
        type: 'tableCell',
        ...(colspan > 1 ? { attrs: { colspan } } : {}),
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
function buildListNode(paragraphs: { node: JSONContent; membership: ListMembership }[]): JSONContent {
  const root: JSONContent = { type: paragraphs[0].membership.ordered ? 'orderedList' : 'bulletList', content: [] };
  const stack: { ilvl: number; list: JSONContent }[] = [{ ilvl: -1, list: root }];

  for (const { node, membership } of paragraphs) {
    while (stack.length > 1 && stack[stack.length - 1].ilvl >= membership.ilvl) {
      stack.pop();
    }
    const parentList = stack[stack.length - 1].list;
    const listItem: JSONContent = { type: 'listItem', content: [node] };
    (parentList.content ??= []).push(listItem);

    const childList: JSONContent = { type: membership.ordered ? 'orderedList' : 'bulletList', content: [] };
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
  const styles = stylesRoot ? parseStyles(stylesRoot) : { resolveHeadingLevel: () => undefined };
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
      const membership = listMembershipFor(pPr, numbering);
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
