import { Extension, isAtStartOfNode } from '@tiptap/core';
import type { Editor } from '@tiptap/core';

export interface IndentOptions {
  types: string[];
}

const INDENT_STEP_PX = 24;

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    indent: {
      indent: () => ReturnType;
      outdent: () => ReturnType;
    };
  }
}

function stepMarginLeft(current: string | null, direction: 1 | -1): string | null {
  const match = /^(\d+(?:\.\d+)?)px$/.exec(current ?? '');
  const currentPx = match ? Number(match[1]) : 0;
  const nextPx = Math.max(0, currentPx + direction * INDENT_STEP_PX);
  return nextPx === 0 ? null : `${nextPx}px`;
}

// List items and table cells have their own Tab semantics (indent list
// depth / move between cells), handled by their own extensions — this
// finds the indentable block type only outside those contexts, so `indent`/
// `outdent` correctly defer (return false) instead of double-handling.
function getIndentableType(editor: Editor, types: string[]): string | null {
  if (editor.isActive('listItem') || editor.isActive('tableCell') || editor.isActive('tableHeader')) {
    return null;
  }
  return types.find((candidate) => editor.isActive(candidate)) ?? null;
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
  addCommands() {
    return {
      indent:
        () =>
        ({ editor, commands }) => {
          const type = getIndentableType(editor, this.options.types);
          if (!type) {
            return false;
          }
          const currentMarginLeft: string | null = editor.getAttributes(type).marginLeft ?? null;
          return commands.updateAttributes(type, { marginLeft: stepMarginLeft(currentMarginLeft, 1) });
        },
      outdent:
        () =>
        ({ editor, commands }) => {
          const type = getIndentableType(editor, this.options.types);
          if (!type) {
            return false;
          }
          const currentMarginLeft: string | null = editor.getAttributes(type).marginLeft ?? null;
          return commands.updateAttributes(type, { marginLeft: stepMarginLeft(currentMarginLeft, -1) });
        },
    };
  },
  addKeyboardShortcuts() {
    return {
      Tab: () => this.editor.commands.indent(),
      'Shift-Tab': () => this.editor.commands.outdent(),
      // Backspace at the very start of an indented block should remove the
      // indent first (matching Word/Google Docs), only merging with the
      // previous block once the indent is already back to zero.
      Backspace: () => {
        const { editor } = this;
        const type = getIndentableType(editor, this.options.types);
        if (!type || !editor.state.selection.empty) {
          return false;
        }
        const currentMarginLeft = editor.getAttributes(type).marginLeft ?? null;
        if (!currentMarginLeft || !isAtStartOfNode(editor.state)) {
          return false;
        }
        return editor.commands.outdent();
      },
    };
  },
});
