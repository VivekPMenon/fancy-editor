import { Node, mergeAttributes } from '@tiptap/core';

// Approximates Word's section-level multi-column layout (<w:cols>) using
// CSS `column-count` — the closest web equivalent, but not an exact match.
// Word's columns are a *page* layout concept: text reflows from column to
// column based on how much vertical space is left on the physical page
// (and can spill onto the next page's columns too). CSS columns reflow
// based on this container's own height, which has no relationship to
// wherever Word's page boundary happened to fall — same visual idea
// (read down one column, then the next), never a pixel-exact
// reproduction. Good enough for a document with a couple of paragraphs in
// a 2-column section; not a real page-layout engine.
//
// Renders as inline style (column-count/column-gap), not an external CSS
// rule keyed off an attribute — deliberately, so the live editor and
// Post Feed's static HTML both get the effect automatically from the same
// node, without needing the same rule duplicated in two stylesheets (a
// mistake made more than once with attribute-keyed CSS elsewhere in this
// codebase).
export const ColumnLayout = Node.create({
  name: 'columnLayout',
  group: 'block',
  content: 'block+',

  addAttributes() {
    return {
      columns: {
        default: 2,
        parseHTML: (element) => Number(element.getAttribute('data-columns')) || 2,
        renderHTML: (attributes) => ({
          'data-columns': attributes.columns,
          style: `column-count: ${attributes.columns}; column-gap: 32px;`,
        }),
      },
    };
  },

  parseHTML() {
    return [{ tag: 'div[data-column-layout]' }];
  },

  renderHTML({ HTMLAttributes }) {
    return ['div', mergeAttributes({ 'data-column-layout': '' }, HTMLAttributes), 0];
  },
});
