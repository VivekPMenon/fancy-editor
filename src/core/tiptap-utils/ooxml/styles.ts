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

export interface StylesResolver {
  resolveHeadingLevel(styleId: string): number | undefined;
  // The document's default run font size (docDefaults -> rPrDefault), as a CSS
  // pt string like "11pt". This is where Word puts the body text size when a
  // run carries no explicit <w:sz> of its own — most body paragraphs. Without
  // it the OOXML path dropped size entirely and text fell back to the
  // editor's 16px base, reading noticeably larger than the source document
  // (and larger than the paste path, which sees the size inlined on spans).
  defaultFontSize?: string;
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

function tag(namespaceURI: string, localName: string, root: Element): Element[] {
  return Array.from(root.getElementsByTagNameNS(namespaceURI, localName));
}
function firstTag(namespaceURI: string, localName: string, root: Element): Element | undefined {
  return root.getElementsByTagNameNS(namespaceURI, localName)[0];
}
function wAttr(el: Element, name: string): string | null {
  return el.getAttributeNS(OOXML_NS.w, name) ?? el.getAttribute(`w:${name}`);
}

export function parseStyles(stylesRoot: Element): StylesResolver {
  const basedOn = new Map<string, string>();
  for (const styleEl of tag(OOXML_NS.w, 'style', stylesRoot)) {
    const styleId = wAttr(styleEl, 'styleId');
    const basedOnEl = firstTag(OOXML_NS.w, 'basedOn', styleEl);
    const basedOnId = basedOnEl ? wAttr(basedOnEl, 'val') : null;
    if (styleId && basedOnId) {
      basedOn.set(styleId, basedOnId);
    }
  }

  const docDefaultsEl = firstTag(OOXML_NS.w, 'docDefaults', stylesRoot);
  const rPrDefaultEl = docDefaultsEl ? firstTag(OOXML_NS.w, 'rPrDefault', docDefaultsEl) : undefined;
  const defaultSzEl = rPrDefaultEl ? firstTag(OOXML_NS.w, 'sz', rPrDefaultEl) : undefined;
  const defaultFontSize = defaultSzEl ? halfPointsToPt(wAttr(defaultSzEl, 'val')) : undefined;

  return {
    defaultFontSize,
    resolveHeadingLevel(styleId) {
      let current: string | undefined = styleId;
      // Cycle guard — a malformed styles.xml could in principle have a
      // basedOn loop; walking it forever would hang the whole conversion.
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
