import { Extension } from '@tiptap/core';

export interface ParagraphStyleOptions {
  types: string[];
}

const DEFAULT_BORDER = '1px solid #444444';

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    paragraphStyle: {
      setLineHeight: (value: string) => ReturnType;
      unsetLineHeight: () => ReturnType;
      setShading: (color: string) => ReturnType;
      unsetShading: () => ReturnType;
      toggleParagraphBorder: () => ReturnType;
    };
  }
}

// Word's "Line and paragraph spacing," "Shading," and "Borders" ribbon
// buttons — none of these have an official Tiptap extension, so this adds
// the three as plain node attributes, same pattern as the Indent extension.
// Borders are simplified to a single uniform 1px rule around the whole
// paragraph rather than Word's full per-side style/width/color control.
export const ParagraphStyle = Extension.create<ParagraphStyleOptions>({
  name: 'paragraphStyle',
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
          lineHeight: {
            default: null,
            parseHTML: (element) => element.style.lineHeight || null,
            renderHTML: (attributes) => {
              if (!attributes.lineHeight) {
                return {};
              }
              return { style: `line-height: ${attributes.lineHeight}` };
            },
          },
          // Space before/after a paragraph (Word's <w:spacing before/after>),
          // as top/bottom margins. Kept separate from the global p-margin CSS
          // so an imported paragraph's own spacing wins where present.
          marginTop: {
            default: null,
            parseHTML: (element) => element.style.marginTop || null,
            renderHTML: (attributes) => (attributes.marginTop ? { style: `margin-top: ${attributes.marginTop}` } : {}),
          },
          marginBottom: {
            default: null,
            parseHTML: (element) => element.style.marginBottom || null,
            renderHTML: (attributes) =>
              attributes.marginBottom ? { style: `margin-bottom: ${attributes.marginBottom}` } : {},
          },
          shading: {
            default: null,
            parseHTML: (element) => element.style.backgroundColor || null,
            renderHTML: (attributes) => {
              if (!attributes.shading) {
                return {};
              }
              return { style: `background-color: ${attributes.shading}` };
            },
          },
          border: {
            default: null,
            parseHTML: (element) => element.style.border || null,
            renderHTML: (attributes) => {
              if (!attributes.border) {
                return {};
              }
              return { style: `border: ${attributes.border}; padding: 4px 8px` };
            },
          },
        },
      },
    ];
  },
  addCommands() {
    return {
      setLineHeight:
        (value: string) =>
        ({ commands }) =>
          this.options.types.some((candidate) => commands.updateAttributes(candidate, { lineHeight: value })),
      unsetLineHeight:
        () =>
        ({ commands }) =>
          this.options.types.some((candidate) => commands.updateAttributes(candidate, { lineHeight: null })),
      setShading:
        (color: string) =>
        ({ commands }) =>
          this.options.types.some((candidate) => commands.updateAttributes(candidate, { shading: color })),
      unsetShading:
        () =>
        ({ commands }) =>
          this.options.types.some((candidate) => commands.updateAttributes(candidate, { shading: null })),
      toggleParagraphBorder:
        () =>
        ({ editor, commands }) => {
          const type = this.options.types.find((candidate) => editor.isActive(candidate));
          if (!type) {
            return false;
          }
          const hasBorder = !!editor.getAttributes(type).border;
          return commands.updateAttributes(type, { border: hasBorder ? null : DEFAULT_BORDER });
        },
    };
  },
});
