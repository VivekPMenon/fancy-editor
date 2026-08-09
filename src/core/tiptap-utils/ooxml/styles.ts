import { OOXML_NS } from './flatOpcPackage';

// Built-in Word heading style IDs -> Tiptap heading levels. A paragraph
// referencing one of these directly (<w:pStyle w:val="Heading2"/>) is
// unambiguous. Real documents often don't stop there, though — Word lets
// authors define *custom* styles derived from a built-in one
// (<w:style w:customStyle="1" w:styleId="Style1"><w:basedOn w:val="Heading2"/>),
// which is common (renaming/tweaking a heading style, or Word itself
// generating an opaque "Style1"-type ID after some edit history) — see
// sample-1.xml's "Lists" section header, which uses exactly this pattern.
// resolveHeadingLevel() walks the basedOn chain so those resolve too,
// instead of silently falling through to a plain paragraph.
const HEADING_STYLE_LEVELS: Record<string, number> = {
  Heading1: 1,
  Heading2: 2,
  Heading3: 3,
  Heading4: 4,
  Heading5: 5,
  Heading6: 6,
};

// Normalized paragraph-level formatting a style (or direct pPr) contributes.
// Twip-based measurements are kept raw here; the converter turns them into CSS.
export interface ParaProps {
  jc?: string;
  indLeftTwips?: number;
  indRightTwips?: number;
  indFirstLineTwips?: number;
  indHangingTwips?: number;
  // <w:spacing> line + lineRule ("auto" = multiple of single spacing, else twips).
  lineValue?: number;
  lineRule?: string;
  spaceBeforeTwips?: number;
  spaceAfterTwips?: number;
  numId?: string;
  ilvl?: number;
  shadingFill?: string; // paragraph <w:shd w:fill>
  hasTopBottomBorder?: boolean; // <w:pBdr> top/bottom (e.g. Intense Quote)
}

// Normalized run-level (character) formatting.
export interface RunProps {
  bold?: boolean;
  italic?: boolean;
  underline?: string; // <w:u w:val> — single/double/dotted/wave/none/...
  strike?: boolean;
  dstrike?: boolean;
  smallCaps?: boolean;
  caps?: boolean;
  color?: string; // hex, no '#'
  fontSizePt?: string; // e.g. "11pt"
  highlight?: string; // named color (yellow, green, ...)
  vertAlign?: 'subscript' | 'superscript';
}

export interface StyleProps {
  para: ParaProps;
  run: RunProps;
}

export interface StylesResolver {
  resolveHeadingLevel(styleId: string): number | undefined;
  // Effective paragraph + run formatting for a style, merged along its
  // basedOn chain (ancestors first) and over docDefaults — the properties a
  // paragraph inherits from `<w:pStyle>` alone, before its own direct
  // formatting. This is what lets style-only formatting (Quote's italics,
  // ListBullet's numbering, a caption's size) reach the output instead of
  // being invisible, the way it was when only headings + docDefaults size
  // were resolved.
  styleProps(styleId: string): StyleProps;
  // docDefaults, applied to every paragraph/run as the base of the cascade.
  defaults: StyleProps;
  // Convenience: docDefaults run size as a CSS pt string (back-compat).
  defaultFontSize?: string;
}

function tag(namespaceURI: string, localName: string, root: Element): Element[] {
  return Array.from(root.getElementsByTagNameNS(namespaceURI, localName));
}
function firstTag(namespaceURI: string, localName: string, root: Element): Element | undefined {
  return root.getElementsByTagNameNS(namespaceURI, localName)[0];
}
function wAttr(el: Element, name: string): string | null {
  return el.getAttributeNS(OOXML_NS.w, name) ?? el.getAttribute(`w:${name}`);
}

// Word stores font sizes in half-points (<w:sz w:val="22"/> = 11pt).
export function halfPointsToPt(val: string | null | undefined): string | undefined {
  if (!val) {
    return undefined;
  }
  const halfPoints = Number(val);
  if (!Number.isFinite(halfPoints)) {
    return undefined;
  }
  const pt = halfPoints / 2;
  return `${Number.isInteger(pt) ? pt : pt.toFixed(1)}pt`;
}

// A <w:XXX> toggle property (b, i, strike, smallCaps, …) is "on" unless it
// carries w:val="0"/"false"; absent element means "not set" (undefined) so it
// doesn't clobber an inherited true.
function toggle(rPr: Element, name: string): boolean | undefined {
  const el = firstTag(OOXML_NS.w, name, rPr);
  if (!el) {
    return undefined;
  }
  const val = wAttr(el, 'val');
  return val !== '0' && val !== 'false';
}

function num(val: string | null): number | undefined {
  if (val === null) {
    return undefined;
  }
  const n = Number(val);
  return Number.isFinite(n) ? n : undefined;
}

// Reads a <w:rPr> element into normalized RunProps. Only properties that are
// actually present are set, so merging up a style chain never overwrites an
// ancestor value with `undefined`.
export function parseRunProps(rPr: Element | undefined): RunProps {
  if (!rPr) {
    return {};
  }
  const props: RunProps = {};
  const b = toggle(rPr, 'b');
  if (b !== undefined) props.bold = b;
  const i = toggle(rPr, 'i');
  if (i !== undefined) props.italic = i;
  const uEl = firstTag(OOXML_NS.w, 'u', rPr);
  if (uEl) props.underline = wAttr(uEl, 'val') ?? 'single';
  const strike = toggle(rPr, 'strike');
  if (strike !== undefined) props.strike = strike;
  const dstrike = toggle(rPr, 'dstrike');
  if (dstrike !== undefined) props.dstrike = dstrike;
  const smallCaps = toggle(rPr, 'smallCaps');
  if (smallCaps !== undefined) props.smallCaps = smallCaps;
  const caps = toggle(rPr, 'caps');
  if (caps !== undefined) props.caps = caps;
  const colorEl = firstTag(OOXML_NS.w, 'color', rPr);
  const colorVal = colorEl ? wAttr(colorEl, 'val') : null;
  if (colorVal && colorVal !== 'auto') props.color = colorVal;
  const szEl = firstTag(OOXML_NS.w, 'sz', rPr);
  const size = szEl ? halfPointsToPt(wAttr(szEl, 'val')) : undefined;
  if (size) props.fontSizePt = size;
  const highlightEl = firstTag(OOXML_NS.w, 'highlight', rPr);
  const highlightVal = highlightEl ? wAttr(highlightEl, 'val') : null;
  if (highlightVal && highlightVal !== 'none') props.highlight = highlightVal;
  const vertAlignEl = firstTag(OOXML_NS.w, 'vertAlign', rPr);
  const vertAlignVal = vertAlignEl ? wAttr(vertAlignEl, 'val') : null;
  if (vertAlignVal === 'subscript' || vertAlignVal === 'superscript') props.vertAlign = vertAlignVal;
  return props;
}

// Reads a <w:pPr> element into normalized ParaProps.
export function parseParaProps(pPr: Element | undefined): ParaProps {
  if (!pPr) {
    return {};
  }
  const props: ParaProps = {};
  const jcEl = firstTag(OOXML_NS.w, 'jc', pPr);
  const jc = jcEl ? wAttr(jcEl, 'val') : null;
  if (jc) props.jc = jc;

  const indEl = firstTag(OOXML_NS.w, 'ind', pPr);
  if (indEl) {
    const left = num(wAttr(indEl, 'left') ?? wAttr(indEl, 'start'));
    const right = num(wAttr(indEl, 'right') ?? wAttr(indEl, 'end'));
    const firstLine = num(wAttr(indEl, 'firstLine'));
    const hanging = num(wAttr(indEl, 'hanging'));
    if (left !== undefined) props.indLeftTwips = left;
    if (right !== undefined) props.indRightTwips = right;
    if (firstLine !== undefined) props.indFirstLineTwips = firstLine;
    if (hanging !== undefined) props.indHangingTwips = hanging;
  }

  const spacingEl = firstTag(OOXML_NS.w, 'spacing', pPr);
  if (spacingEl) {
    const line = num(wAttr(spacingEl, 'line'));
    if (line !== undefined) {
      props.lineValue = line;
      props.lineRule = wAttr(spacingEl, 'lineRule') ?? undefined;
    }
    const before = num(wAttr(spacingEl, 'before'));
    const after = num(wAttr(spacingEl, 'after'));
    if (before !== undefined) props.spaceBeforeTwips = before;
    if (after !== undefined) props.spaceAfterTwips = after;
  }

  const numPr = firstTag(OOXML_NS.w, 'numPr', pPr);
  if (numPr) {
    const numIdEl = firstTag(OOXML_NS.w, 'numId', numPr);
    const ilvlEl = firstTag(OOXML_NS.w, 'ilvl', numPr);
    const numId = numIdEl ? wAttr(numIdEl, 'val') : null;
    if (numId) props.numId = numId;
    const ilvl = ilvlEl ? num(wAttr(ilvlEl, 'val')) : undefined;
    if (ilvl !== undefined) props.ilvl = ilvl;
  }

  const shdEl = firstTag(OOXML_NS.w, 'shd', pPr);
  const shdFill = shdEl ? wAttr(shdEl, 'fill') : null;
  if (shdFill && shdFill !== 'auto') props.shadingFill = shdFill;

  const pBdr = firstTag(OOXML_NS.w, 'pBdr', pPr);
  if (pBdr && (firstTag(OOXML_NS.w, 'top', pBdr) || firstTag(OOXML_NS.w, 'bottom', pBdr))) {
    props.hasTopBottomBorder = true;
  }

  return props;
}

// Merge b over a (b wins where defined), for both prop kinds.
function mergeRun(a: RunProps, b: RunProps): RunProps {
  return { ...a, ...b };
}
function mergePara(a: ParaProps, b: ParaProps): ParaProps {
  return { ...a, ...b };
}

export function parseStyles(stylesRoot: Element): StylesResolver {
  const basedOn = new Map<string, string>();
  const ownProps = new Map<string, StyleProps>();

  for (const styleEl of tag(OOXML_NS.w, 'style', stylesRoot)) {
    const styleId = wAttr(styleEl, 'styleId');
    if (!styleId) {
      continue;
    }
    const basedOnEl = firstTag(OOXML_NS.w, 'basedOn', styleEl);
    const basedOnId = basedOnEl ? wAttr(basedOnEl, 'val') : null;
    if (basedOnId) {
      basedOn.set(styleId, basedOnId);
    }
    // A style's own pPr/rPr are direct children; scope the reads to them so a
    // nested rPr inside pPr (paragraph mark props) doesn't bleed into either.
    const pPrEl = Array.from(styleEl.children).find((c) => isW(c, 'pPr'));
    const rPrEl = Array.from(styleEl.children).find((c) => isW(c, 'rPr'));
    ownProps.set(styleId, { para: parseParaProps(pPrEl), run: parseRunProps(rPrEl) });
  }

  const docDefaultsEl = firstTag(OOXML_NS.w, 'docDefaults', stylesRoot);
  const rPrDefaultEl = docDefaultsEl ? firstTag(OOXML_NS.w, 'rPrDefault', docDefaultsEl) : undefined;
  const pPrDefaultEl = docDefaultsEl ? firstTag(OOXML_NS.w, 'pPrDefault', docDefaultsEl) : undefined;
  const defaults: StyleProps = {
    run: parseRunProps(rPrDefaultEl ? firstTag(OOXML_NS.w, 'rPr', rPrDefaultEl) : undefined),
    para: parseParaProps(pPrDefaultEl ? firstTag(OOXML_NS.w, 'pPr', pPrDefaultEl) : undefined),
  };

  const effectiveCache = new Map<string, StyleProps>();

  function styleProps(styleId: string): StyleProps {
    const cached = effectiveCache.get(styleId);
    if (cached) {
      return cached;
    }
    // Build the basedOn chain from the given style up to its root ancestor,
    // then merge root-first so nearer styles win, over docDefaults.
    const chain: string[] = [];
    const seen = new Set<string>();
    let current: string | undefined = styleId;
    while (current && !seen.has(current)) {
      seen.add(current);
      chain.push(current);
      current = basedOn.get(current);
    }
    let para = defaults.para;
    let run = defaults.run;
    for (let i = chain.length - 1; i >= 0; i--) {
      const own = ownProps.get(chain[i]);
      if (own) {
        para = mergePara(para, own.para);
        run = mergeRun(run, own.run);
      }
    }
    const result = { para, run };
    effectiveCache.set(styleId, result);
    return result;
  }

  return {
    defaults,
    defaultFontSize: defaults.run.fontSizePt,
    styleProps,
    resolveHeadingLevel(styleId) {
      let current: string | undefined = styleId;
      const seen = new Set<string>();
      while (current && !seen.has(current)) {
        const level = HEADING_STYLE_LEVELS[current];
        if (level) {
          return level;
        }
        seen.add(current);
        current = basedOn.get(current);
      }
      return undefined;
    },
  };
}

function isW(el: Element, localName: string): boolean {
  return el.namespaceURI === OOXML_NS.w && el.localName === localName;
}
