import { unzipSync, strFromU8 } from 'fflate';
import { makePackage, type OoxmlPackage, type OoxmlPart } from './flatOpcPackage';

// A .docx is an OPC package delivered as a ZIP — the very same parts
// getOoxml() returns as flat-OPC XML (word/document.xml, styles.xml,
// numbering.xml, media/*, the _rels), just zipped instead of flattened into
// one XML doc. So this only has to *unzip* into the shared OoxmlPackage shape;
// the entire conversion layer (ooxmlToTiptapJson + styles/numbering/drawing)
// is reused unchanged. That's the whole point of keeping the package
// abstraction separate from the conversion — the upload path and the add-in
// path differ only in how the parts are unpacked.
//
// This lets us ingest documents with no Word add-in at all: read a .docx the
// user uploads, convert with the exact same pipeline. See
// docs/word-integration-notes.md §18.

// Word part names in flat-OPC are absolute ("/word/document.xml"); ZIP entries
// are relative ("word/document.xml"). Key parts the same way so the shared
// package logic (and the converter's hard-coded "/word/..." lookups) match.
function toPartName(zipPath: string): string {
  return zipPath.startsWith('/') ? zipPath : `/${zipPath}`;
}

function isXmlPart(path: string): boolean {
  return /\.(xml|rels)$/i.test(path);
}

// btoa needs a binary string; chunk it so a large image can't blow the stack.
function bytesToBase64(bytes: Uint8Array): string {
  const chunkSize = 0x8000;
  let binary = '';
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}

export function parseDocxPackage(bytes: Uint8Array): OoxmlPackage {
  let files: Record<string, Uint8Array>;
  try {
    files = unzipSync(bytes);
  } catch (err) {
    throw new Error(`Not a readable .docx (ZIP) file: ${(err as Error).message}`);
  }

  const parts = new Map<string, OoxmlPart>();
  for (const [path, data] of Object.entries(files)) {
    if (path.endsWith('/')) {
      continue; // directory entry
    }
    const name = toPartName(path);
    if (isXmlPart(path)) {
      // Strip a leading UTF-8 BOM — some OOXML parts carry one and it trips
      // DOMParser (it would appear as a stray character before the root).
      const text = strFromU8(data).replace(/^﻿/, '');
      const doc = new DOMParser().parseFromString(text, 'application/xml');
      const parserError = doc.getElementsByTagName('parsererror')[0];
      // contentType left empty: convertDrawing falls back to the part name's
      // extension (guessMimeFromPartName), which is all we need for media.
      parts.set(name, {
        name,
        contentType: '',
        root: parserError ? undefined : (doc.documentElement ?? undefined),
      });
    } else {
      parts.set(name, { name, contentType: '', base64: bytesToBase64(data) });
    }
  }

  if (!parts.get('/word/document.xml')?.root) {
    throw new Error("This .docx has no readable /word/document.xml — it may be corrupt or not a Word document.");
  }

  return makePackage(parts);
}
