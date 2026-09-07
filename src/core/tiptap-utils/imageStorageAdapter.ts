// Swappable image-storage abstraction — per the API team discussion, pasted/
// imported images should be uploaded to their backend rather than embedded
// as base64 in the document. This is the ONLY file that should need to
// change when wiring that up for real; everything else (imageUploadExtension.ts,
// and any future caller) only ever calls uploadImage()/resolveImageSrc()
// below, never localStorage or a real endpoint directly — copy this whole
// file into the UBS workspace and replace the two function bodies.
//
// Contract:
//   - uploadImage(dataUri) sends the image bytes and resolves to the URL
//     that gets stored as the Tiptap image node's `src` attribute — and
//     therefore into the saved article JSON. A real backend serves that URL
//     directly over HTTP (e.g. "/api/images/{imageId}"), so nothing else
//     needs to know it was ever a data: URI.
//   - resolveImageSrc(src) exists ONLY because this mock has no real server
//     to answer "/api/images/{id}" — it turns that logical URL back into
//     something the browser can actually paint, at render time, without
//     touching the document/JSON. Once uploadImage() calls the real API,
//     "/api/images/{id}" is a genuine working URL and this becomes a
//     straight pass-through (see the TODO below).

const MOCK_URL_PREFIX = '/api/images/';
const STORAGE_KEY_PREFIX = 'publisher-plus:mock-image:';

function generateImageId(): string {
  if (typeof crypto !== 'undefined' && typeof crypto.randomUUID === 'function') {
    return crypto.randomUUID();
  }
  return `img-${Date.now()}-${Math.random().toString(36).slice(2, 10)}`;
}

/**
 * Uploads an image (as a data: URI — every path that produces a fresh image
 * in this app already hands us one: paste, the Word RTF quality upgrade,
 * DOCX/OOXML import) and resolves to the URL that should replace it as the
 * node's `src` from then on.
 *
 * MOCK implementation: stashes the data: URI in localStorage under a
 * generated id and returns "/api/images/{id}" — the exact shape the real API
 * is expected to return.
 *
 * TODO(UBS integration): replace this body with the real upload, e.g.
 *   const res = await fetch('/api/images', { method: 'POST', body: toBlob(dataUri) });
 *   const { imageId } = await res.json();
 *   return `/api/images/${imageId}`;
 */
export async function uploadImage(dataUri: string): Promise<string> {
  const id = generateImageId();
  try {
    localStorage.setItem(STORAGE_KEY_PREFIX + id, dataUri);
  } catch (err) {
    // localStorage full/unavailable (private browsing, quota) — surface it
    // the same way a real upload failure would, so the caller can leave the
    // inline image as-is rather than lose it silently.
    throw new Error(`Mock image upload failed to persist locally: ${(err as Error).message}`, { cause: err });
  }
  return `${MOCK_URL_PREFIX}${id}`;
}

/**
 * Resolves an image `src` to something the browser can actually render right
 * now. Mock URLs ("/api/images/{id}") are looked up in localStorage;
 * anything else (a real http(s) URL, an inline data: URI that hasn't been
 * uploaded yet) passes through unchanged.
 *
 * TODO(UBS integration): once uploadImage() returns real, directly-loadable
 * URLs, delete the localStorage lookup below and just `return src ?? '';` —
 * a real URL needs no client-side resolution.
 */
export function resolveImageSrc(src: string | null | undefined): string {
  if (!src || !src.startsWith(MOCK_URL_PREFIX)) {
    return src ?? '';
  }
  const id = src.slice(MOCK_URL_PREFIX.length);
  try {
    return localStorage.getItem(STORAGE_KEY_PREFIX + id) ?? src;
  } catch {
    return src;
  }
}
