// Body.getOoxml() returns "flat OPC" — a single XML document representing
// an entire .docx package (normally a ZIP of many files) as a flat list of
// <pkg:part> elements, each holding either XML content (<pkg:xmlData>, e.g.
// document.xml, numbering.xml, styles.xml) or binary content
// (<pkg:binaryData>, base64 text — embedded media like images). This module
// only unpacks that structure into an addressable form; it knows nothing
// about what any particular part's content *means* (that's ooxmlToTiptapJson.ts
// and friends).
export const OOXML_NS = {
  pkg: 'http://schemas.microsoft.com/office/2006/xmlPackage',
  w: 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
  r: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
  rel: 'http://schemas.openxmlformats.org/package/2006/relationships',
  wp: 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
  a: 'http://schemas.openxmlformats.org/drawingml/2006/main',
  pic: 'http://schemas.openxmlformats.org/drawingml/2006/picture',
  // Word 2012 drawing extensions — carries <wp15:webVideoPr> for an
  // Insert -> Online Video, whose embeddedHtml attribute holds the real embed
  // <iframe> (exact video URL), not just a poster image.
  wp15: 'http://schemas.microsoft.com/office/word/2012/wordprocessingDrawing',
} as const;

export interface OoxmlPart {
  name: string;
  contentType: string;
  /** Present for XML parts (<pkg:xmlData>) — the part's root element (e.g. <w:document>, <w:numbering>). */
  root?: Element;
  /** Present for binary parts (<pkg:binaryData>) — base64 text, whitespace already stripped. */
  base64?: string;
}

export interface OoxmlPackage {
  parts: Map<string, OoxmlPart>;
  getPart(name: string): OoxmlPart | undefined;
  /** Convenience for XML parts — throws if the part is missing or has no XML root, since callers that need it can't proceed without it. */
  requireXml(name: string): Element;
  /**
   * Resolves a relationship ID (e.g. an <a:blip r:embed="rId7">) to the
   * target part's full path, using `fromPart`'s own .rels file — OOXML
   * relationship targets are relative to the *referencing* part's
   * directory, not the package root (confirmed against a real capture:
   * document.xml's rels resolve "media/image1.gif" to "/word/media/image1.gif",
   * not "/media/image1.gif").
   */
  resolveRelationship(fromPart: string, relationshipId: string): string | undefined;
}

function tag(namespaceURI: string, localName: string, root: Element | Document): Element[] {
  return Array.from(root.getElementsByTagNameNS(namespaceURI, localName));
}

export function parseFlatOpcPackage(flatOpcXml: string): OoxmlPackage {
  const doc = new DOMParser().parseFromString(flatOpcXml, 'application/xml');
  const parserError = doc.getElementsByTagName('parsererror')[0];
  if (parserError) {
    throw new Error(`Failed to parse flat-OPC XML: ${parserError.textContent}`);
  }

  const parts = new Map<string, OoxmlPart>();
  for (const partEl of tag(OOXML_NS.pkg, 'part', doc)) {
    const name = partEl.getAttributeNS(OOXML_NS.pkg, 'name') ?? partEl.getAttribute('pkg:name') ?? '';
    const contentType = partEl.getAttributeNS(OOXML_NS.pkg, 'contentType') ?? partEl.getAttribute('pkg:contentType') ?? '';
    const xmlDataEl = tag(OOXML_NS.pkg, 'xmlData', partEl)[0];
    const binaryDataEl = tag(OOXML_NS.pkg, 'binaryData', partEl)[0];

    parts.set(name, {
      name,
      contentType,
      root: xmlDataEl?.firstElementChild ?? undefined,
      base64: binaryDataEl ? (binaryDataEl.textContent ?? '').replace(/\s+/g, '') : undefined,
    });
  }

  function getPart(name: string) {
    return parts.get(name);
  }

  function requireXml(name: string): Element {
    const root = parts.get(name)?.root;
    if (!root) {
      throw new Error(`OOXML package is missing expected XML part: ${name}`);
    }
    return root;
  }

  // .rels files live in a "_rels" subfolder alongside the part they
  // describe, named "<partFileName>.rels" — e.g. "/word/document.xml"'s
  // relationships are at "/word/_rels/document.xml.rels". Confirmed
  // against a real capture.
  function relsPartNameFor(partName: string): string {
    const lastSlash = partName.lastIndexOf('/');
    const dir = partName.slice(0, lastSlash);
    const fileName = partName.slice(lastSlash + 1);
    return `${dir}/_rels/${fileName}.rels`;
  }

  function resolveRelationship(fromPart: string, relationshipId: string): string | undefined {
    const relsPart = parts.get(relsPartNameFor(fromPart));
    if (!relsPart?.root) {
      return undefined;
    }
    // relsPart.root IS the <Relationships> element itself (that's the
    // xmlData's direct child) — its <Relationship> children are its
    // descendants, so a plain namespaced search within it finds them.
    const relEls = tag(OOXML_NS.rel, 'Relationship', relsPart.root);
    const match = relEls.find((el) => el.getAttribute('Id') === relationshipId);
    if (!match) {
      return undefined;
    }
    const target = match.getAttribute('Target');
    if (!target) {
      return undefined;
    }
    if (target.startsWith('http://') || target.startsWith('https://')) {
      // External relationship (hyperlinks) — not a package-relative path at all.
      return target;
    }
    // Resolve relative to fromPart's own directory (not the rels part's
    // directory, and not the package root).
    const dir = fromPart.slice(0, fromPart.lastIndexOf('/'));
    return `${dir}/${target}`;
  }

  return { parts, getPart, requireXml, resolveRelationship };
}
