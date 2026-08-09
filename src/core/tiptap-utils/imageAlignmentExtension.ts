import { Extension } from '@tiptap/core';
import { NodeSelection, TextSelection } from '@tiptap/pm/state';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';

// 'left'/'center'/'right' position the image as its own block (no text
// wrap — same as Word's simple, non-wrapping picture alignment, §11/§12).
// 'float-left'/'float-right' are real text-wrap (Word's Wrap Text =
// Square/Tight, §14/§15) — previously only ever set by Word-paste
// preprocessing (wordImageAlignment.ts), now also available as a live
// editing command, see EditorContextMenu.tsx's "Wrap Text" section.
export type ImageAlign = 'left' | 'center' | 'right' | 'float-left' | 'float-right';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    imageAlignment: {
      setImageAlign: (align: ImageAlign | null) => ReturnType;
    };
  }
}

// Gives the (block-level) Image node an alignment attribute of its own.
// Word never expresses this on the <img> itself — wordImageAlignment.ts's
// preprocessing pass transfers it there (as data-align) from the wrapping
// paragraph before parsing, since that paragraph doesn't survive parsing
// once the image is hoisted out of it. setImageAlign is the editing-time
// equivalent — the import path and the live-editing path both end up
// setting the same node attribute, just via different triggers.
//
// Deliberately renders ONLY the data-align attribute, never a `style`.
// @tiptap/extension-image's resize node view re-syncs an image's HTML
// attributes on every attribute-driven update by doing a wholesale
// `el.setAttribute('style', ...)` from the freshly rendered attribute set
// (see its onUpdate) — but width/height from a drag-resize are applied
// imperatively (`el.style.width = ...`), outside that attribute set. If we
// contributed a `style` here, committing a resize (which itself dispatches
// an attribute update, to persist width/height) would wipe that same
// `style` attribute clean, snapping the image back to its natural size the
// instant the drag handle is released. `data-align` avoids this because
// it's synced attribute-by-attribute, not as a whole-string overwrite.
// Actual centering/right/left alignment is done via CSS keyed off
// data-align — see App.css (`[data-resize-container]:has(...)`, needed
// because the live editor always wraps images in a flex resize container)
// and PostFeedTab.css (`img[data-align=...]`, for the plain exported-HTML
// case with no such wrapper).
export const ImageAlignment = Extension.create({
  name: 'imageAlignment',
  addCommands() {
    return {
      setImageAlign:
        (align: ImageAlign | null) =>
        ({ state, tr, dispatch }) => {
          const imageType = state.schema.nodes.image;
          const paragraphType = state.schema.nodes.paragraph;

          // Locate the image this command targets — a NodeSelection on it, or
          // the first image within the current selection range.
          let found: { pos: number; node: ProseMirrorNode } | null = null;
          const { selection } = state;
          if (selection instanceof NodeSelection && selection.node.type === imageType) {
            found = { pos: selection.from, node: selection.node };
          } else {
            state.doc.nodesBetween(selection.from, selection.to, (node, pos) => {
              if (!found && node.type === imageType) {
                found = { pos, node };
              }
            });
          }
          if (!found) {
            return false;
          }
          const imagePos = found.pos;
          const imageNode = found.node;

          tr.setNodeMarkup(imagePos, undefined, { ...imageNode.attrs, align });

          // For a text-wrap float, make sure a paragraph follows the image and
          // drop the cursor into it. A block image (often the last node) other-
          // wise leaves nowhere to type beside it — CSS float would wrap text
          // there, but there's no text position for the caret, so the wrap
          // "works" visually yet you can't add the wrapping content. See
          // docs/word-integration-notes.md §17.
          if (align === 'float-left' || align === 'float-right') {
            const afterPos = imagePos + imageNode.nodeSize;
            const nodeAfter = tr.doc.resolve(afterPos).nodeAfter;
            if (!nodeAfter || nodeAfter.type !== paragraphType) {
              tr.insert(afterPos, paragraphType.create());
            }
            // afterPos + 1 lands inside that (new or existing) paragraph.
            tr.setSelection(TextSelection.create(tr.doc, afterPos + 1));
            tr.scrollIntoView();
          }

          if (dispatch) {
            dispatch(tr);
          }
          return true;
        },
    };
  },
  addGlobalAttributes() {
    return [
      {
        types: ['image'],
        attributes: {
          align: {
            default: null,
            parseHTML: (element) => element.getAttribute('data-align'),
            renderHTML: (attributes) => {
              if (!attributes.align) {
                return {};
              }
              return { 'data-align': attributes.align };
            },
          },
        },
      },
    ];
  },
});
