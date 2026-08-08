import { Extension, isAtStartOfNode } from '@tiptap/core';
import type { Editor } from '@tiptap/core';

export interface IndentOptions {
  types: string[];
}

// Whole-block indent step (the toolbar "Increase/Decrease Indent"), vs the
// first-line indent a Tab keypress applies — a first-line indent uses a
// full tab stop (0.5in ≈ 48px), matching Word.
const INDENT_STEP_PX = 24;
const FIRST_LINE_INDENT_STEP_PX = 48;

declare module '@tiptap/core' {
  interface Commands<ReturnType> {
    indent: {
      indent: () => ReturnType;
      outdent: () => ReturnType;
    };
  }
}

// Steps a `<value>px` string up or down by `step`, clamped at 0 (returns
// null at 0 so the attribute is dropped rather than rendered as 0px).
// Generic over margin-left and text-indent — both are px values here.
function stepPx(current: string | null, direction: 1 | -1, step: number): string | null {
  const match = /^(\d+(?:\.\d+)?)px$/.exec(current ?? '');
  const currentPx = match ? Number(match[1]) : 0;
  const nextPx = Math.max(0, currentPx + direction * step);
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
          return commands.updateAttributes(type, { marginLeft: stepPx(currentMarginLeft, 1, INDENT_STEP_PX) });
        },
      outdent:
        () =>
        ({ editor, commands }) => {
          const type = getIndentableType(editor, this.options.types);
          if (!type) {
            return false;
          }
          const currentMarginLeft: string | null = editor.getAttributes(type).marginLeft ?? null;
          return commands.updateAttributes(type, { marginLeft: stepPx(currentMarginLeft, -1, INDENT_STEP_PX) });
        },
    };
  },
  addKeyboardShortcuts() {
    // Collapsed cursor: a *first-line* indent (only the first line moves),
    // matching Word/Google Docs — not the whole-block margin shift.
    const stepFirstLine = (direction: 1 | -1): void => {
      const { editor } = this;
      const type = getIndentableType(editor, this.options.types);
      if (!type) {
        return;
      }
      const current: string | null = editor.getAttributes(type).textIndent ?? null;
      if (direction === -1 && !current) {
        // Nothing left on the first line to pull back — fall through to a
        // whole-block outdent so Shift-Tab still reduces any block indent.
        editor.commands.outdent();
        return;
      }
      editor.commands.updateAttributes(type, {
        textIndent: stepPx(current, direction, FIRST_LINE_INDENT_STEP_PX),
      });
    };

    // Range selection: shift the whole block's left margin, like Word does
    // when you Tab with a paragraph selected. Walk every indentable block the
    // selection touches and step each one's own margin-left individually —
    // updateAttributes would flatten them all to a single value, and
    // getIndentableType's isActive() check reports false the moment a
    // selection spans a block boundary (which is exactly the case that made
    // Tab fall through to the browser and move focus out of the editor).
    const stepBlockMarginForSelection = (direction: 1 | -1): void => {
      const { editor } = this;
      const { state } = editor;
      const { from, to } = state.selection;
      const { tr } = state;
      let changed = false;
      state.doc.nodesBetween(from, to, (node, pos) => {
        if (this.options.types.includes(node.type.name)) {
          const current: string | null = node.attrs.marginLeft ?? null;
          tr.setNodeAttribute(pos, 'marginLeft', stepPx(current, direction, INDENT_STEP_PX));
          changed = true;
        }
      });
      if (changed) {
        editor.view.dispatch(tr);
      }
    };

    // Shift the whole enclosing list right/left as a block via margin-left on
    // the list node itself — the fallback when Tab can't nest (see below).
    // Returns false when there's no list here, or when outdenting a list that
    // has no block shift left to remove (so the caller can un-nest instead).
    const stepListMargin = (direction: 1 | -1): boolean => {
      const { editor } = this;
      const listType = editor.isActive('orderedList')
        ? 'orderedList'
        : editor.isActive('bulletList')
          ? 'bulletList'
          : null;
      if (!listType) {
        return false;
      }
      const current: string | null = editor.getAttributes(listType).marginLeft ?? null;
      if (direction === -1 && !current) {
        return false;
      }
      return editor.commands.updateAttributes(listType, {
        marginLeft: stepPx(current, direction, INDENT_STEP_PX),
      });
    };

    // Tab/Shift-Tab return true (swallow the key) in every editable context
    // we own, so focus can never escape to the next form control — even when
    // the step itself is a no-op (e.g. Tab on the first item of a list, which
    // can't nest). Only table cells and code blocks defer (return false), so
    // their own Tab semantics — next-cell / insert-literal-tab — still win.
    const handleTab = (direction: 1 | -1): boolean => {
      const { editor } = this;
      if (editor.isActive('tableCell') || editor.isActive('tableHeader') || editor.isActive('codeBlock')) {
        return false;
      }
      // List item: nest / un-nest. sinkListItem/liftListItem return false
      // when the move isn't possible (the first item has no sibling to nest
      // under), so we can't rely on the list extension's own Tab binding —
      // its false return is exactly what let the bare key move focus out.
      // Handle it here and always swallow instead. When nesting is refused
      // (typically a selection covering the list's first item), fall back to
      // shifting the whole list right as a block, so a whole-list selection
      // still "moves as a whole" like a paragraph does.
      if (editor.isActive('listItem')) {
        if (direction === 1) {
          if (!editor.commands.sinkListItem('listItem')) {
            stepListMargin(1);
          }
        } else {
          // Shift-Tab peels back any block shift first, then un-nests.
          if (!stepListMargin(-1)) {
            editor.commands.liftListItem('listItem');
          }
        }
        return true;
      }
      if (editor.state.selection.empty) {
        stepFirstLine(direction);
      } else {
        stepBlockMarginForSelection(direction);
      }
      return true;
    };

    return {
      Tab: () => handleTab(1),
      'Shift-Tab': () => handleTab(-1),
      // Backspace at the very start of an indented block peels indentation
      // back before merging with the previous block (matching Word/Google
      // Docs) — first-line indent first, then block margin.
      Backspace: () => {
        const { editor } = this;
        const type = getIndentableType(editor, this.options.types);
        if (!type || !editor.state.selection.empty || !isAtStartOfNode(editor.state)) {
          return false;
        }
        const currentTextIndent: string | null = editor.getAttributes(type).textIndent ?? null;
        if (currentTextIndent) {
          return editor.commands.updateAttributes(type, {
            textIndent: stepPx(currentTextIndent, -1, FIRST_LINE_INDENT_STEP_PX),
          });
        }
        if (editor.getAttributes(type).marginLeft) {
          return editor.commands.outdent();
        }
        return false;
      },
    };
  },
});
