import { Mark, mergeAttributes } from '@tiptap/core';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    encloseCharacters: {
      toggleEncloseCharacters: () => ReturnType;
    };
  }
}

// Word's "Enclose Characters" — wraps the selection in a circle, matching
// the most common of Word's enclosure shapes (square/triangle/diamond
// aren't offered here — one shape covers the demo).
export const EncloseCharacters = Mark.create({
  name: 'encloseCharacters',
  parseHTML() {
    return [{ tag: 'span[data-enclose-characters]' }];
  },
  renderHTML({ HTMLAttributes }) {
    return [
      'span',
      mergeAttributes(HTMLAttributes, {
        'data-enclose-characters': '',
        style:
          'display: inline-flex; align-items: center; justify-content: center; min-width: 1.6em; height: 1.6em; padding: 0 0.15em; border: 1px solid currentColor; border-radius: 50%; line-height: 1;',
      }),
      0,
    ];
  },
  addCommands() {
    return {
      toggleEncloseCharacters:
        () =>
        ({ commands }) =>
          commands.toggleMark(this.name),
    };
  },
});
