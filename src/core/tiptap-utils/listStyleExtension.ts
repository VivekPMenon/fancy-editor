import { Extension } from '@tiptap/core';

// Gives orderedList a `listStyleType` node attribute that round-trips as an
// inline `list-style-type` CSS value. Word documents choose a specific
// numbering format per nesting level (decimal, upper-roman, lower-alpha, …)
// and both our import pipelines resolve it — the OOXML path from
// numbering.xml's numFmt, the paste path from the literal marker token — but
// without somewhere to put it, that format was being dropped and every list
// fell back to a hardcoded CSS cycle that differed between the editor and the
// Post Feed (and matched neither the source document). This lets the real
// format travel on the node so every surface renders the same markers Word
// did. Lists authored in the editor leave it null and still fall back to the
// CSS cycle in App.css / PostFeedTab.css.
export const ListStyle = Extension.create({
  name: 'listStyle',
  addGlobalAttributes() {
    return [
      {
        types: ['orderedList'],
        attributes: {
          listStyleType: {
            default: null,
            parseHTML: (element) => element.style.listStyleType || null,
            renderHTML: (attributes) => {
              if (!attributes.listStyleType) {
                return {};
              }
              return { style: `list-style-type: ${attributes.listStyleType}` };
            },
          },
        },
      },
    ];
  },
});
