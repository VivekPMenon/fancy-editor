import { Extension } from '@tiptap/core';
import { Fragment, Slice } from '@tiptap/pm/model';
import type { Node as ProseMirrorNode, NodeType } from '@tiptap/pm/model';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { preprocessWordHtml } from './wordHtmlPreprocessing';
import { extractRtfPictures, type RtfPicture } from './wordRtfImage';

// Upper bound on how many images in one paste we'll attempt to upgrade.
// Not a hard technical ceiling — it's a safety valve: the more images in
// one paste, the more that could go wrong in the size-based correlation
// below (see selectPictureForEachImage), and demos/articles here are a
// handful of paragraphs with a couple of images, not dozens.
const MAX_IMAGES_TO_UPGRADE = 5;

// Copy-pasting from Word gives the browser the same "filtered HTML" shape as
// body.getHtml()/Save As Web Page — a full mini-document with a <style>
// block driving heading color/font-size/font-family, Word's linked-
// thumbnail video markup, and fake numbered/bulleted paragraphs instead of
// real lists. Tiptap's default paste path parses that HTML straight into
// the schema with none of htmlToTiptapJson's preprocessing, so pasted
// content would silently lose all of it. Running the same
// preprocessWordHtml() steps here keeps paste and "load a saved article" on
// equal footing.
//
// Known gap (narrow): if the copied selection contains a genuine drawn/vector
// Shape (text box, freeform shape, WordArt — a real Shape, not a floating
// *picture*), Word can fail to populate `text/html` on the clipboard at all —
// prosemirror-view then takes the plain-text path and never calls this hook,
// so the *entire* paste silently loses all formatting, not just the shape. No
// code-level fix for that (there's nothing to preprocess). See §10b.
//
// What does NOT trigger it, both verified on the clipboard path:
//   - floating/anchored *pictures* (Square/Tight wrap, incl. the chart the RTF
//     handling below upgrades) — they populate `text/html` fine (§14);
//   - *comment anchors* — a paste of a selection with a comment (and a
//     floating image) retained full formatting (2026-08-14). The comment was
//     wrongly implicated earlier only because the original failing test also
//     contained a drawn Shape; isolating the variables cleared it. So the
//     trigger is genuine drawn Shapes specifically, not comments or pictures.
export const WordPaste = Extension.create({
  name: 'wordPaste',
  transformPastedHTML(html) {
    // Large/complex documents (footnotes, nested tables, embedded fonts)
    // could in principle trip up juice's CSS parsing on some malformed
    // stylesheet. If that happens, fall back to the original HTML rather
    // than losing the paste entirely — imperfect formatting beats nothing.
    try {
      return preprocessWordHtml(html);
    } catch (err) {
      console.error('WordPaste: preprocessing failed, pasting unprocessed HTML', err);
      return html;
    }
  },
  // Embedded pictures/charts (Insert -> Chart, or any OLE-embedded object)
  // only ever reach the HTML clipboard format as a legacy, low-resolution
  // raster fallback — Word's VML block carries a real vector version, but
  // it's wrapped in IE-only conditional comments no modern browser parses,
  // and even the raster fallback is typically upscaled well past its native
  // pixel size once it lands in our schema (confirmed by decoding an actual
  // captured GIF header: 430x216 native vs. a much taller declared display
  // size). Separately, and independently of any of that: Word's clipboard
  // *also* populates `text/rtf`, which represents embedded pictures inline
  // via `\pict` groups with real hex-encoded PNG/JPEG bytes (confirmed
  // against a real capture — valid PNG/JPEG magic bytes decoded directly
  // from it) — a completely different, often much higher-fidelity
  // representation than whatever the HTML fragment references. This is the
  // same class of fix CKEditor ships for the same reason (see
  // github.com/ckeditor/ckeditor5/pull/18363).
  //
  // ProseMirror parses clipboardData.getData('text/html') into `slice`
  // *before* calling handlePaste (confirmed by reading prosemirror-view's
  // doPaste source) — so by the time we get here, transformPastedHTML above
  // has already run and `slice` already reflects every other Word fix
  // (lists, alignment, spacing). We only ever touch `src` here.
  addProseMirrorPlugins() {
    return [
      // Pasting a whole imported document leaves the caret — and the scroll
      // position — at the very bottom, forcing a manual scroll back up to
      // review what just came in; reset to the top once the paste settles
      // (deferred a frame so it wins over ProseMirror's own scroll-caret-
      // into-view). But ONLY for a genuine document import: an in-editor
      // copy-paste of a small chunk shouldn't yank the view to the top.
      // ProseMirror tags its own clipboard HTML with data-pm-slice, which
      // reliably identifies an internal paste; and we additionally require the
      // paste to actually be large (a Word document, or a lot of text), since
      // a small paste won't have jumped the view anyway.
      new Plugin({
        key: new PluginKey('wordPasteScrollTop'),
        props: {
          handleDOMEvents: {
            paste: (view, event) => {
              const html = event.clipboardData?.getData('text/html') ?? '';
              const text = event.clipboardData?.getData('text/plain') ?? '';
              const isInternalPaste = html.includes('data-pm-slice');
              const isDocumentImport =
                /mso-|urn:schemas-microsoft|class=["']?Mso/i.test(html) || text.length > 1200;
              if (!isInternalPaste && isDocumentImport) {
                requestAnimationFrame(() => {
                  const scroller = view.dom.closest('.tiptap-editor');
                  if (scroller instanceof HTMLElement) {
                    scroller.scrollTop = 0;
                  }
                });
              }
              return false; // don't consume — let the paste proceed normally
            },
          },
        },
      }),
      new Plugin({
        key: new PluginKey('wordPasteImageQuality'),
        props: {
          handlePaste: (view, event, slice) => {
            try {
              const rtf = event.clipboardData?.getData('text/rtf');
              if (!rtf) {
                console.log('[WordPaste quality] bailed: no text/rtf on clipboard');
                return false;
              }
              const pictures = extractRtfPictures(rtf);
              if (pictures.length === 0) {
                console.log('[WordPaste quality] bailed: text/rtf present (', rtf.length, 'chars) but no pngblip/jpegblip pictures extracted');
                return false;
              }
              console.log(
                '[WordPaste quality] extracted pictures (document order):',
                pictures.map((p) => ({ mimeType: p.mimeType, bytes: p.bytes.length, position: p.position })),
              );

              const imageType = view.state.schema.nodes.image;
              if (!imageType) {
                console.log('[WordPaste quality] bailed: no image node type in schema (unexpected)');
                return false;
              }
              const imageNodes = collectNodesOfType(slice.content, imageType);
              if (imageNodes.length === 0 || imageNodes.length > MAX_IMAGES_TO_UPGRADE) {
                console.log(
                  '[WordPaste quality] bailed: pasted slice has',
                  imageNodes.length,
                  `image node(s), need 1-${MAX_IMAGES_TO_UPGRADE}`,
                );
                return false;
              }
              if (pictures.length < imageNodes.length) {
                // Not enough RTF candidates to confidently give every image
                // an upgrade — rather than guess which ones go without,
                // leave every image exactly as the HTML fragment already
                // resolved it.
                console.log(
                  '[WordPaste quality] bailed:',
                  imageNodes.length,
                  'image(s) but only',
                  pictures.length,
                  'usable RTF picture(s)',
                );
                return false;
              }

              const chosen = selectPictureForEachImage(pictures, imageNodes.length);
              const dataUrls = chosen.map((p) => `data:${p.mimeType};base64,${bytesToBase64(p.bytes)}`);

              const newContent = replaceImageSrcsInOrder(slice.content, imageType, dataUrls);
              const newSlice = new Slice(newContent, slice.openStart, slice.openEnd);
              view.dispatch(
                view.state.tr
                  .replaceSelection(newSlice)
                  .scrollIntoView()
                  .setMeta('paste', true)
                  .setMeta('uiEvent', 'paste'),
              );
              console.log(
                `WordPaste: upgraded ${chosen.length} pasted image(s) from RTF —`,
                chosen.map((p) => `${p.mimeType} ${p.bytes.length}b`).join(', '),
              );
              return true;
            } catch (err) {
              console.error('WordPaste: RTF image-quality upgrade failed, pasting unmodified', err);
              return false;
            }
          },
        },
      }),
    ];
  },
});

// Correlating N images to N-of-M RTF pictures with no shared identifier
// between the HTML and RTF clipboard formats — the best available signal is
// document order (both are serializations of the same selection, so
// left-to-right/top-to-bottom order should match) combined with size
// (real full-fidelity renditions run much bigger than the small cached
// icon/thumbnail pictures Word embeds alongside them — confirmed
// empirically: a 277-byte "icon" appeared consistently across every real
// capture regardless of what was actually copied).
//
// So: take the `imageCount` LARGEST pictures (filtering out the small
// junk/thumbnail candidates), then re-sort THOSE survivors back into
// document order before pairing position-wise with the image nodes (which
// are also walked in document order). This is a heuristic, not a guarantee
// — it assumes each real image contributes at most one dominant
// large-byte picture to the RTF stream, which held in every capture tested
// but isn't a hard property of the RTF format.
function selectPictureForEachImage(pictures: RtfPicture[], imageCount: number): RtfPicture[] {
  const topByFidelity = [...pictures].sort((a, b) => b.bytes.length - a.bytes.length).slice(0, imageCount);
  return topByFidelity.sort((a, b) => a.position - b.position);
}

function collectNodesOfType(fragment: Fragment, type: NodeType): ProseMirrorNode[] {
  const found: ProseMirrorNode[] = [];
  fragment.forEach((node) => {
    if (node.type === type) {
      found.push(node);
    }
    if (node.content.size > 0) {
      found.push(...collectNodesOfType(node.content, type));
    }
  });
  return found;
}

// Walks the fragment in document order, assigning `srcs[0]` to the first
// image node encountered, `srcs[1]` to the second, and so on — `srcs` must
// already be in the same document order as the image nodes themselves
// (collectNodesOfType and this function both walk depth-first,
// left-to-right, so that's guaranteed as long as callers sort `srcs` the
// same way, which selectPictureForEachImage does).
function replaceImageSrcsInOrder(fragment: Fragment, type: NodeType, srcs: string[]): Fragment {
  let nextIndex = 0;
  const walk = (frag: Fragment): Fragment => {
    const mapped: ProseMirrorNode[] = [];
    frag.forEach((node) => {
      if (node.type === type) {
        const src = srcs[nextIndex];
        nextIndex += 1;
        mapped.push(src ? node.type.create({ ...node.attrs, src }, node.content, node.marks) : node);
        return;
      }
      if (node.content.size > 0) {
        mapped.push(node.copy(walk(node.content)));
        return;
      }
      mapped.push(node);
    });
    return Fragment.fromArray(mapped);
  };
  return walk(fragment);
}

// btoa expects a "binary string" (one char per byte); String.fromCharCode
// on the whole array at once can blow the call stack for large images, so
// this builds it up in chunks.
function bytesToBase64(bytes: Uint8Array): string {
  const chunkSize = 0x8000;
  let binary = '';
  for (let i = 0; i < bytes.length; i += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(i, i + chunkSize));
  }
  return btoa(binary);
}
