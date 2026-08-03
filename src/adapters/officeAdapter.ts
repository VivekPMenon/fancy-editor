import type { DocumentAdapter } from '../core/types';

// Word's JS API has no per-keystroke content-changed event, so we poll the
// document body on a short interval as a pragmatic stand-in for "live" updates.
const POLL_INTERVAL_MS = 400;

export const officeAdapter: DocumentAdapter = {
  async getContentHtml() {
    return Word.run(async (context) => {
      const htmlResult = context.document.body.getHtml();
      await context.sync();
      return htmlResult.value;
    });
  },

  async getSelectionText() {
    return Word.run(async (context) => {
      const range = context.document.getSelection();
      range.load('text');
      await context.sync();
      return range.text;
    });
  },

  async replaceSelection(text) {
    return Word.run(async (context) => {
      const range = context.document.getSelection();
      range.insertText(text, Word.InsertLocation.replace);
      await context.sync();
    });
  },

  async insertHtml(html) {
    return Word.run(async (context) => {
      const range = context.document.getSelection();
      range.insertHtml(html, Word.InsertLocation.replace);
      await context.sync();
    });
  },

  onContentChange(callback) {
    let lastText: string | null = null;

    const poll = async () => {
      const text = await Word.run(async (context) => {
        const body = context.document.body;
        body.load('text');
        await context.sync();
        return body.text;
      });
      if (text !== lastText) {
        lastText = text;
        callback(text);
      }
    };

    poll();
    const intervalId = setInterval(poll, POLL_INTERVAL_MS);
    return () => clearInterval(intervalId);
  },
};
