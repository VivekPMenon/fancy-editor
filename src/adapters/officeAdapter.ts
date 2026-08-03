import type { DocumentAdapter } from '../core/types';

// Word's JS API has no per-keystroke content-changed event, so we poll the
// document body on a short interval as a pragmatic stand-in for "live" updates.
const POLL_INTERVAL_MS = 400;

// `body.getHtml()` references images via a path like
// "~WRS{...}_files/image001.jpg" — a leftover from Word's old "Web Page,
// Filtered" export convention that assumes a companion _files folder Word
// never actually creates for us. It does NOT embed images as base64. This
// patches each <img> tag's src with the real image bytes fetched separately
// via inlinePictures, matched positionally in document order.
function inlineImageSources(html: string, base64Images: string[]): string {
  if (base64Images.length === 0) {
    return html;
  }
  const container = document.createElement('div');
  container.innerHTML = html;
  container.querySelectorAll('img').forEach((img, index) => {
    const base64 = base64Images[index];
    if (base64) {
      img.src = `data:image/png;base64,${base64}`;
    }
  });
  return container.innerHTML;
}

export const officeAdapter: DocumentAdapter = {
  async getContentHtml() {
    return Word.run(async (context) => {
      const body = context.document.body;
      const htmlResult = body.getHtml();

      const pictures = body.inlinePictures;
      pictures.load('items');
      await context.sync();

      const base64Results = pictures.items.map((picture) => picture.getBase64ImageSrc());
      await context.sync();

      const base64Images = base64Results.map((result) => result.value);
      return inlineImageSources(htmlResult.value, base64Images);
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
