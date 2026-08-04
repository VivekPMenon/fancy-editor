import { Extension } from '@tiptap/core';

export interface IndentOptions {
  types: string[];
}

// Tiptap ships no official indent extension. Word expresses list-style
// paragraphs (a literal "1." or "·" typed inline, followed by a hanging
// indent) via `margin-left`/`text-indent` on the <p> itself — this just
// gives those two CSS properties a place to land as node attributes,
// mirroring how TextAlign/Color/etc. are wired up.
export const Indent = Extension.create<IndentOptions>({
  name: 'indent',
  addOptions() {
    return {
      types: ['paragraph', 'heading'],
    };
  },
  addGlobalAttributes() {
    return [
      {
        types: this.options.types,
        attributes: {
          marginLeft: {
            default: null,
            parseHTML: (element) => element.style.marginLeft || null,
            renderHTML: (attributes) => {
              if (!attributes.marginLeft) {
                return {};
              }
              return { style: `margin-left: ${attributes.marginLeft}` };
            },
          },
          textIndent: {
            default: null,
            parseHTML: (element) => element.style.textIndent || null,
            renderHTML: (attributes) => {
              if (!attributes.textIndent) {
                return {};
              }
              return { style: `text-indent: ${attributes.textIndent}` };
            },
          },
        },
      },
    ];
  },
});
