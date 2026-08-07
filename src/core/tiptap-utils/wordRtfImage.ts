// RTF represents a picture as `{\pict <picttype> <attributes...><hex bytes>}`
// where <picttype> is one of \pngblip/\jpegblip/\wmetafile/etc. Confirmed
// against a real Word clipboard RTF capture that <attributes> isn't just
// plain control words (\bliptagNNNN) — it can also include a nested
// destination group (`{\*\blipuid ...}`, a picture cache identifier) before
// the actual hex payload starts. A naive "skip words then read hex" scan
// undercounts unless it also skips that nested group. Only \pngblip/
// \jpegblip are extracted — \wmetafile/\emfblip are vector formats with no
// browser-native renderer, so there'd be nothing to do with their bytes.
export interface RtfPicture {
  mimeType: 'image/png' | 'image/jpeg';
  bytes: Uint8Array;
  // Character offset of the tag in the RTF string — the only thing that
  // lets a caller correlate multiple pictures back to multiple images in
  // true document order (Word's RTF and HTML clipboard formats are two
  // independent serializations of the same selection, but both preserve
  // left-to-right/top-to-bottom reading order).
  position: number;
}

const TAG_MIME: Record<string, RtfPicture['mimeType']> = {
  '\\pngblip': 'image/png',
  '\\jpegblip': 'image/jpeg',
};

export function extractRtfPictures(rtf: string): RtfPicture[] {
  const pictures: RtfPicture[] = [];
  for (const [tag, mimeType] of Object.entries(TAG_MIME)) {
    let searchFrom = 0;
    for (;;) {
      const idx = rtf.indexOf(tag, searchFrom);
      if (idx < 0) {
        break;
      }
      searchFrom = idx + tag.length;
      const hex = readHexPayload(rtf, idx + tag.length);
      // Real picture payloads are large; anything tiny here means the scan
      // landed somewhere unexpected (malformed/unfamiliar RTF variant) —
      // skip rather than risk producing a broken image from garbage bytes.
      if (hex.length >= 32) {
        pictures.push({ mimeType, bytes: hexToBytes(hex), position: idx });
      }
    }
  }
  // Scanned per-tag-type above (all \pngblip, then all \jpegblip), which
  // does NOT reflect true document order when formats are mixed — sort by
  // position before returning so callers can rely on array order matching
  // left-to-right document order.
  return pictures.sort((a, b) => a.position - b.position);
}

// Starting right after a picture-type tag (e.g. the position right after
// `\pngblip`), skips any plain RTF control words and one balanced `{...}`
// destination group, then reads the contiguous run of hex digits that
// follows — that run is the actual image payload.
function readHexPayload(rtf: string, start: number): string {
  let i = start;
  while (i < rtf.length) {
    const ch = rtf[i];
    if (ch === ' ' || ch === '\r' || ch === '\n' || ch === '\t') {
      i += 1;
      continue;
    }
    if (ch === '{') {
      // Skip a balanced group without parsing its contents — it's either
      // the \*\blipuid cache-id destination or something else we don't
      // need, but it isn't the image payload either way.
      let depth = 1;
      i += 1;
      while (i < rtf.length && depth > 0) {
        if (rtf[i] === '\\') {
          i += 2;
          continue;
        }
        if (rtf[i] === '{') {
          depth += 1;
        } else if (rtf[i] === '}') {
          depth -= 1;
        }
        i += 1;
      }
      continue;
    }
    if (ch === '\\') {
      // A control word: backslash, letters, optional numeric parameter,
      // optional single trailing space delimiter (RTF spec).
      i += 1;
      while (i < rtf.length && /[a-zA-Z]/.test(rtf[i])) {
        i += 1;
      }
      while (i < rtf.length && /[-0-9]/.test(rtf[i])) {
        i += 1;
      }
      if (rtf[i] === ' ') {
        i += 1;
      }
      continue;
    }
    // First non-whitespace, non-brace, non-control-word character — if
    // it's hex, this is where the payload starts; if not, bail rather than
    // silently returning garbage.
    break;
  }
  const hexStart = i;
  while (i < rtf.length && /[0-9a-fA-F\s]/.test(rtf[i])) {
    i += 1;
  }
  return rtf.slice(hexStart, i).replace(/\s+/g, '');
}

function hexToBytes(hex: string): Uint8Array {
  const clean = hex.length % 2 === 0 ? hex : hex.slice(0, -1);
  const bytes = new Uint8Array(clean.length / 2);
  for (let i = 0; i < bytes.length; i += 1) {
    bytes[i] = parseInt(clean.slice(i * 2, i * 2 + 2), 16);
  }
  return bytes;
}
