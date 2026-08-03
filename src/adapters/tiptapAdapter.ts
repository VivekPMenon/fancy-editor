import type { Editor } from '@tiptap/react';
import type { DocumentAdapter } from '../core/types';

export function createTiptapAdapter(editor: Editor): DocumentAdapter {
  return {
    async getContentHtml() {
      return editor.getHTML();
    },

    async getSelectionText() {
      const { from, to } = editor.state.selection;
      return editor.state.doc.textBetween(from, to, ' ');
    },

    async replaceSelection(text) {
      editor.chain().focus().deleteSelection().insertContent(text).run();
    },

    async insertHtml(html) {
      editor.chain().focus().deleteSelection().insertContent(html).run();
    },

    onContentChange(callback) {
      const handler = () => callback(editor.getText());
      handler();
      editor.on('update', handler);
      return () => editor.off('update', handler);
    },
  };
}
