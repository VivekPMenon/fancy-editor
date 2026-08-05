import type { Editor } from '@tiptap/react';
import type { Node as PMNode } from '@tiptap/pm/model';
import { Fragment } from '@tiptap/pm/model';

export type CaseMode = 'upper' | 'lower' | 'title' | 'sentence';

function transformCase(text: string, mode: CaseMode): string {
  switch (mode) {
    case 'upper':
      return text.toUpperCase();
    case 'lower':
      return text.toLowerCase();
    case 'title':
      return text.replace(/\w\S*/g, (word) => word[0].toUpperCase() + word.slice(1).toLowerCase());
    case 'sentence':
      return text.toLowerCase().replace(/(^\s*\w|[.!?]\s*\w)/g, (match) => match.toUpperCase());
  }
}

// Word's "Change Case" applies to the selected text while preserving each
// run's own formatting (bold/color/etc.) — replacing the selection with a
// single insertContent(string) call would flatten all of that into plain
// text, so this walks each text node individually and rebuilds it with the
// same marks it already had.
export function changeSelectionCase(editor: Editor, mode: CaseMode) {
  const { state, view } = editor;
  const { from, to } = state.selection;
  if (from === to) {
    return;
  }

  const replacements: { from: number; to: number; text: string; marks: PMNode['marks'] }[] = [];
  state.doc.nodesBetween(from, to, (node, pos) => {
    if (!node.isText || !node.text) {
      return;
    }
    const nodeFrom = Math.max(from, pos);
    const nodeTo = Math.min(to, pos + node.nodeSize);
    const slice = node.text.slice(nodeFrom - pos, nodeTo - pos);
    replacements.push({ from: nodeFrom, to: nodeTo, text: transformCase(slice, mode), marks: node.marks });
  });

  const tr = state.tr;
  // Back-to-front so each replacement's positions stay valid as we go.
  for (let i = replacements.length - 1; i >= 0; i -= 1) {
    const r = replacements[i];
    tr.replaceWith(r.from, r.to, state.schema.text(r.text, r.marks));
  }
  view.dispatch(tr);
  editor.commands.focus();
}

function sortKey(node: PMNode): string {
  return node.textContent.trim().toLowerCase();
}

// Word's "Sort" reorders the selected paragraphs/list items alphabetically
// by their own text — resolved via ProseMirror's blockRange, which finds
// the nearest shared parent (a list, or the document itself) and the index
// range of its children the selection touches.
export function sortSelection(editor: Editor, direction: 'asc' | 'desc') {
  const { state, view } = editor;
  const { $from, $to } = state.selection;
  const range = $from.blockRange($to);
  if (!range) {
    return;
  }

  const { parent, startIndex, endIndex, start } = range;
  const nodes: PMNode[] = [];
  for (let i = startIndex; i < endIndex; i += 1) {
    nodes.push(parent.child(i));
  }
  if (nodes.length < 2) {
    return;
  }

  const sorted = [...nodes].sort((a, b) => {
    const cmp = sortKey(a).localeCompare(sortKey(b));
    return direction === 'asc' ? cmp : -cmp;
  });

  const end = start + nodes.reduce((sum, node) => sum + node.nodeSize, 0);
  const tr = state.tr.replaceWith(start, end, Fragment.fromArray(sorted));
  view.dispatch(tr);
  editor.commands.focus();
}
