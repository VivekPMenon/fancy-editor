import { Node, mergeAttributes } from '@tiptap/core';

// A single Table-of-Contents line: entry text on the left, a dotted leader
// filling the middle, and the page number flush right — the visual shape of
// Word's TOC field. We deliberately do NOT reproduce the field itself (no live
// tab stops, no page recalculation); the OOXML import reads the entry's text
// and its already-computed page number and lays them out with flexbox + a
// dotted border for the leader, purely so it *looks* right. See
// docs/word-integration-notes.md.
export const TocEntry = Node.create({
  name: 'tocEntry',
  group: 'block',
  content: 'inline*',

  addAttributes() {
    return {
      level: {
        default: 1,
        parseHTML: (element) => Number(element.getAttribute('data-level')) || 1,
        renderHTML: (attributes) => ({ 'data-level': attributes.level }),
      },
      page: {
        default: '',
        parseHTML: (element) => element.querySelector('.toc-entry-page')?.textContent ?? '',
        // page is rendered as its own span in renderHTML, not an attribute.
        renderHTML: () => ({}),
      },
    };
  },

  parseHTML() {
    return [{ tag: 'div[data-toc-entry]', contentElement: '.toc-entry-text' }];
  },

  renderHTML({ node, HTMLAttributes }) {
    const level = Number(node.attrs.level) || 1;
    return [
      'div',
      mergeAttributes(HTMLAttributes, {
        'data-toc-entry': '',
        class: 'toc-entry',
        style: `padding-left: ${(level - 1) * 18}px`,
      }),
      ['span', { class: 'toc-entry-text' }, 0],
      ['span', { class: 'toc-entry-leader', contenteditable: 'false' }],
      ['span', { class: 'toc-entry-page', contenteditable: 'false' }, node.attrs.page ?? ''],
    ];
  },
});
