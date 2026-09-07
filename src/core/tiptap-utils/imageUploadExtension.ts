import { Extension } from '@tiptap/core';
import Image from '@tiptap/extension-image';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import { uploadImage, resolveImageSrc } from './imageStorageAdapter';

// Extends the official Image node so its `src` resolves through the storage
// adapter at render time only (the live editor's own paint, and
// editor.getHTML()/generateHTML() for Post Feed) — the underlying node
// attribute, and therefore editor.getJSON()/the saved article JSON, keeps
// whatever URL imageUploadExtension's sweep below wrote there, untouched.
// Drop-in replacement for `Image` — configure it exactly the same way
// (allowBase64/resize options are untouched, only addAttributes is
// overridden, and only for `src`).
export const ResolvingImage = Image.extend({
  addAttributes() {
    // Tiptap's `this.parent()` typing in an `.extend()` override doesn't
    // narrow to the parent's actual attribute-config shape, hence the cast —
    // the parent Image node's addAttributes() genuinely returns this shape
    // (confirmed against @tiptap/extension-image's source: src/alt/title/
    // width/height, each `{ default: ... }`).
    const parentAttributes = (this.parent?.() ?? {}) as Record<string, { default: unknown }>;
    return {
      ...parentAttributes,
      src: {
        ...parentAttributes.src,
        renderHTML: (attributes: { src?: string }) => ({ src: resolveImageSrc(attributes.src) }),
      },
    };
  },
});

// Background sweep: any image node whose src is still a raw data: URI —
// freshly pasted, RTF-upgraded (wordPasteExtension.ts's image-quality
// plugin), or produced by the DOCX/OOXML import path — gets uploaded through
// the storage adapter and its src swapped in place for the returned URL.
// One choke point covers every path that can introduce a base64 image
// without touching any of that code.
//
// appendTransaction must return synchronously, so the upload itself can't be
// awaited inline — it's kicked off here and, once resolved, dispatches its
// own follow-up transaction. Matching on the data: URI string itself (rather
// than a captured position) means a slow upload racing against other edits
// is fine: the position is re-found fresh, at dispatch time, in whatever the
// document looks like by then.
const inFlightUploads = new Set<string>();

export const ImageUpload = Extension.create({
  name: 'imageUpload',
  addProseMirrorPlugins() {
    const editor = this.editor;
    return [
      new Plugin({
        key: new PluginKey('imageUpload'),
        appendTransaction(transactions, _oldState, newState) {
          if (!transactions.some((tr) => tr.docChanged)) {
            return null;
          }

          const pendingSrcs = new Set<string>();
          newState.doc.descendants((node) => {
            const src = node.attrs?.src;
            if (
              node.type.name === 'image' &&
              typeof src === 'string' &&
              src.startsWith('data:') &&
              !inFlightUploads.has(src)
            ) {
              pendingSrcs.add(src);
            }
          });

          pendingSrcs.forEach((dataUri) => {
            inFlightUploads.add(dataUri);
            uploadImage(dataUri)
              .then((url) => {
                const { state, dispatch } = editor.view;
                const tr = state.tr;
                let changed = false;
                state.doc.descendants((node, pos) => {
                  if (node.type.name === 'image' && node.attrs.src === dataUri) {
                    tr.setNodeAttribute(pos, 'src', url);
                    changed = true;
                  }
                });
                if (changed) {
                  dispatch(tr);
                }
              })
              .catch((err) => {
                console.error('Image upload failed — leaving the inline image as-is', err);
              })
              .finally(() => inFlightUploads.delete(dataUri));
          });

          return null;
        },
      }),
    ];
  },
});

// A raw clipboard image (a screenshot, an image copied from another app) has
// no HTML/RTF at all — just an image file on the clipboard — so today's
// paste path (wordPasteExtension.ts, entirely about Word's HTML+RTF shape)
// never sees it and nothing gets inserted. This adds that base capability:
// read any image file off the clipboard, insert it as a plain image node
// (as a data: URI), which ImageUpload above then picks up and uploads like
// any other freshly-inserted image.
export const PasteImageFile = Extension.create({
  name: 'pasteImageFile',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey('pasteImageFile'),
        props: {
          handlePaste: (view, event) => {
            // Defer entirely to the normal HTML paste path (WordPaste and
            // friends) whenever there's real HTML on the clipboard — Word
            // sometimes also puts a file representation alongside its HTML,
            // and this must never compete with that carefully-tuned
            // pipeline. This only ever fires for a "just an image, nothing
            // else" paste (a screenshot, a file copied from Explorer/Finder).
            if (event.clipboardData?.getData('text/html')) {
              return false;
            }
            const files = Array.from(event.clipboardData?.files ?? []).filter((file) =>
              file.type.startsWith('image/'),
            );
            if (files.length === 0) {
              return false;
            }
            event.preventDefault();
            files.forEach((file) => {
              const reader = new FileReader();
              reader.onload = () => {
                const dataUri = reader.result;
                if (typeof dataUri !== 'string') {
                  return;
                }
                const imageType = view.state.schema.nodes.image;
                if (!imageType) {
                  return;
                }
                view.dispatch(view.state.tr.replaceSelectionWith(imageType.create({ src: dataUri })));
              };
              reader.readAsDataURL(file);
            });
            return true;
          },
        },
      }),
    ];
  },
});
