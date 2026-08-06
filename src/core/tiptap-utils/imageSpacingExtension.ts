import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';

// Our Image node is a block-level atom (no text content of its own). If
// there's no paragraph immediately before/after it, there's nowhere for a
// typed character — including a plain space — to go: you can navigate a
// cursor next to the image (Tiptap's bundled Gapcursor handles that), but a
// gap cursor position isn't inside any text-containing node, so typing does
// nothing there. Word-imported images routinely land with no such gap.
//
// This guarantees one: on every document change, any top-level image
// missing a text-block sibling on either side gets an empty paragraph
// inserted there. Scoped to top-level document children only — the case
// that actually occurs (Word's standalone images are direct body children,
// not nested in a list item or table cell).
export const ImageSpacing = Extension.create({
  name: 'imageSpacing',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: new PluginKey('imageSpacing'),
        appendTransaction: (transactions, _oldState, newState) => {
          if (!transactions.some((tr) => tr.docChanged)) {
            return null;
          }

          const { doc, schema } = newState;
          const imageType = schema.nodes.image;
          const paragraphType = schema.nodes.paragraph;
          if (!imageType || !paragraphType) {
            return null;
          }

          const insertions = new Set<number>();
          doc.forEach((node, offset, index) => {
            if (node.type !== imageType) {
              return;
            }
            const before = index > 0 ? doc.child(index - 1) : null;
            const after = index < doc.childCount - 1 ? doc.child(index + 1) : null;

            if (!before?.isTextblock) {
              insertions.add(offset);
            }
            if (!after?.isTextblock) {
              insertions.add(offset + node.nodeSize);
            }
          });

          if (insertions.size === 0) {
            return null;
          }

          // Insert highest position first — inserting at a position only
          // shifts positions *after* it, so processing descending keeps
          // every other queued position valid without needing tr.mapping.
          const tr = newState.tr;
          Array.from(insertions)
            .sort((a, b) => b - a)
            .forEach((pos) => tr.insert(pos, paragraphType.create()));
          return tr;
        },
      }),
    ];
  },
});
