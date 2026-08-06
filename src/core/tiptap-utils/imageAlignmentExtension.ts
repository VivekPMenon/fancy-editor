import { Extension } from '@tiptap/core';

const ALIGN_STYLES: Record<string, string> = {
  center: 'display: block; margin-left: auto; margin-right: auto;',
  right: 'display: block; margin-left: auto; margin-right: 0;',
  left: 'display: block; margin-left: 0; margin-right: auto;',
};

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    imageAlignment: {
      setImageAlign: (align: 'left' | 'center' | 'right') => ReturnType;
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
export const ImageAlignment = Extension.create({
  name: 'imageAlignment',
  addCommands() {
    return {
      setImageAlign:
        (align: 'left' | 'center' | 'right') =>
        ({ commands }) =>
          commands.updateAttributes('image', { align }),
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
              const style = ALIGN_STYLES[attributes.align as string];
              if (!style) {
                return {};
              }
              return { style, 'data-align': attributes.align };
            },
          },
        },
      },
    ];
  },
});
