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

    async setContentHtml(html) {
      editor.chain().focus().setContent(html).run();
    },

    async setContentJson(json) {
      // Tiptap consumes JSON natively — no HTML round-trip needed here,
      // unlike setContentHtml.
      editor.chain().focus().setContent(json).run();
    },

    onContentChange(callback) {
      const handler = () => callback(editor.getText());
      handler();
      editor.on('update', handler);
      return () => editor.off('update', handler);
    },

    async getContentWarnings() {
      // No floating-image concept on the web/Tiptap side — our Image
      // extension only supports simple inline images.
      return [];
    },

    async highlightFlaggedTerms() {
      // Already highlighted live by the FlaggedTerms decoration plugin — nothing to do.
      return 0;
    },
  };
}
