import type { Editor } from '@tiptap/react';
import type { DocumentAdapter } from '../core/types';

export function createTiptapAdapter(editor: Editor): DocumentAdapter {
  // Loading an article (or a blank doc) should show its top, not wherever the
  // caret happens to land — focus('start') puts the caret at the document
  // start and scrolls it into view, and this resets the scroll container
  // outright as a belt-and-suspenders against any late scroll-into-view.
  function scrollEditorToTop() {
    const scroller = editor.view.dom.closest('.tiptap-editor');
    if (scroller instanceof HTMLElement) {
      scroller.scrollTop = 0;
    }
  }

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
      editor.chain().setContent(html).focus('start').run();
      scrollEditorToTop();
    },

    async setContentJson(json) {
      // Tiptap consumes JSON natively — no HTML round-trip needed here,
      // unlike setContentHtml.
      editor.chain().setContent(json).focus('start').run();
      scrollEditorToTop();
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

    async getContentJson() {
      return editor.getJSON();
    },

    async highlightFlaggedTerms() {
      // Already highlighted live by the FlaggedTerms decoration plugin — nothing to do.
      return 0;
    },
  };
}
