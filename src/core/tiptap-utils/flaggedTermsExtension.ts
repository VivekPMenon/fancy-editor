import { Extension } from '@tiptap/core';
import { Plugin, PluginKey } from '@tiptap/pm/state';
import type { Node as ProseMirrorNode } from '@tiptap/pm/model';
import { Decoration, DecorationSet } from '@tiptap/pm/view';
import { FLAGGED_TERMS } from '../flaggedTerms';

const pluginKey = new PluginKey('flaggedTerms');

function buildDecorations(doc: ProseMirrorNode): DecorationSet {
  const decorations: Decoration[] = [];

  doc.descendants((node, pos) => {
    if (!node.isText || !node.text) {
      return;
    }
    const text = node.text;
    FLAGGED_TERMS.forEach(({ term, comment }) => {
      const regex = new RegExp(`\\b${term}\\b`, 'gi');
      let match: RegExpExecArray | null;
      while ((match = regex.exec(text))) {
        const from = pos + match.index;
        const to = from + match[0].length;
        decorations.push(Decoration.inline(from, to, { class: 'flagged-term', title: comment }));
        decorations.push(
          Decoration.widget(
            to,
            () => {
              const indicator = document.createElement('span');
              indicator.className = 'flagged-term-indicator';
              indicator.textContent = '⚠';
              indicator.title = comment;
              return indicator;
            },
            { side: 1 },
          ),
        );
      }
    });
  });

  return DecorationSet.create(doc, decorations);
}

/** Highlights hardcoded flagged terms (see core/flaggedTerms.ts) via decorations only — never touches document content. */
export const FlaggedTerms = Extension.create({
  name: 'flaggedTerms',
  addProseMirrorPlugins() {
    return [
      new Plugin({
        key: pluginKey,
        state: {
          init: (_, { doc }) => buildDecorations(doc),
          apply: (tr, old) => (tr.docChanged ? buildDecorations(tr.doc) : old),
        },
        props: {
          decorations(state) {
            return pluginKey.getState(state);
          },
        },
      }),
    ];
  },
});
