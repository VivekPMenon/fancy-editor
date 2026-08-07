import { OOXML_NS } from './flatOpcPackage';

// OOXML list membership (<w:numPr> on a paragraph) only gives a numId +
// ilvl — the actual format (decimal, bullet, upperRoman, ...) per level
// lives in numbering.xml, cross-referenced through two levels of
// indirection: numId -> abstractNumId -> per-ilvl <w:lvl> definitions.
// This resolves that chain so callers can ask "is numId 6, level 1 a
// bullet or a number?" directly.
export type NumFormat = string;

export interface NumberingResolver {
  getLevelFormat(numId: string, ilvl: number): NumFormat | undefined;
}

function tag(namespaceURI: string, localName: string, root: Element): Element[] {
  return Array.from(root.getElementsByTagNameNS(namespaceURI, localName));
}

function wAttr(el: Element, name: string): string | null {
  return el.getAttributeNS(OOXML_NS.w, name) ?? el.getAttribute(`w:${name}`);
}

export function parseNumbering(numberingRoot: Element): NumberingResolver {
  const abstractLevels = new Map<string, Map<number, NumFormat>>();
  for (const abstractNumEl of tag(OOXML_NS.w, 'abstractNum', numberingRoot)) {
    const abstractNumId = wAttr(abstractNumEl, 'abstractNumId');
    if (!abstractNumId) {
      continue;
    }
    const levels = new Map<number, NumFormat>();
    for (const lvlEl of tag(OOXML_NS.w, 'lvl', abstractNumEl)) {
      const ilvlStr = wAttr(lvlEl, 'ilvl');
      const numFmtEl = tag(OOXML_NS.w, 'numFmt', lvlEl)[0];
      const format = numFmtEl ? wAttr(numFmtEl, 'val') : null;
      if (ilvlStr !== null && format) {
        levels.set(Number(ilvlStr), format);
      }
    }
    abstractLevels.set(abstractNumId, levels);
  }

  const numToAbstract = new Map<string, string>();
  for (const numEl of tag(OOXML_NS.w, 'num', numberingRoot)) {
    const numId = wAttr(numEl, 'numId');
    const abstractNumIdEl = tag(OOXML_NS.w, 'abstractNumId', numEl)[0];
    const abstractNumId = abstractNumIdEl ? wAttr(abstractNumIdEl, 'val') : null;
    if (numId && abstractNumId) {
      numToAbstract.set(numId, abstractNumId);
    }
    // Note: <w:lvlOverride> (per-numId level format overrides) isn't
    // handled — not present in sample-1.xml's numbering.xml, and a real
    // gap if a future document relies on one. Would show up as this
    // resolver returning the abstract's base format instead of the
    // override.
  }

  return {
    getLevelFormat(numId, ilvl) {
      const abstractNumId = numToAbstract.get(numId);
      if (!abstractNumId) {
        return undefined;
      }
      return abstractLevels.get(abstractNumId)?.get(ilvl);
    },
  };
}
