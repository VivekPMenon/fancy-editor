import type { DocumentAdapter } from '../core/types';
import { FLAGGED_TERMS } from '../core/flaggedTerms';
import { tiptapJsonToHtml } from '../core/tiptap-utils/htmlJsonConversion';

// Word's JS API has no per-keystroke content-changed event, so we poll the
// document body on a short interval as a pragmatic stand-in for "live" updates.
const POLL_INTERVAL_MS = 400;

// getBase64ImageSrc() returns raw base64 with no format indicator, so we
// always assumed PNG — wrong for "Insert Icons" pictures, which are SVG
// markup underneath (confirmed by decoding one: it starts with "<svg ...").
// Labeling SVG bytes as "image/png" produces a data URI the <img> decoder
// can't parse (declared MIME wins over sniffing), so the icon never
// renders. Decoding a short prefix to check for an SVG opening tag before
// picking the MIME type fixes that.
function isSvgBase64(base64: string): boolean {
  try {
    // atob() needs a multiple of 4 characters; 200 is plenty to see past
    // any XML prologue regardless of the full image's size.
    const prefix = atob(base64.slice(0, 200 - (200 % 4)));
    return /^\s*(<\?xml|<svg)/i.test(prefix);
  } catch {
    return false;
  }
}

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
      const mimeType = isSvgBase64(base64) ? 'image/svg+xml' : 'image/png';
      img.src = `data:${mimeType};base64,${base64}`;
    }
  });
  return container.innerHTML;
}

// Shared by setContentHtml and setContentJson — Body.insertHtml with
// `replace` swaps out the entire body content in one call.
function replaceBodyHtml(html: string): Promise<void> {
  return Word.run(async (context) => {
    context.document.body.insertHtml(html, Word.InsertLocation.replace);
    await context.sync();
  });
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

  async setContentHtml(html) {
    return replaceBodyHtml(html);
  },

  async setContentJson(json) {
    // Word has no JSON concept — this is the one adapter where "prefer
    // JSON" still means converting, but that conversion lives here rather
    // than in every caller. Proves the stored JSON (our source of truth)
    // round-trips through HTML back into Word.
    return replaceBodyHtml(tiptapJsonToHtml(json));
  },

  async getContentWarnings() {
    return Word.run(async (context) => {
      const shapes = context.document.body.shapes;
      shapes.load('items');
      await context.sync();

      shapes.items.forEach((shape) => shape.load(['type', 'name', 'altTextDescription', 'textWrap/type']));
      await context.sync();

      // Floating/anchored pictures are Word.Shape objects, not
      // InlinePicture — body.inlinePictures never sees them, so
      // getContentHtml() can't capture their bytes (Shape has no
      // getBase64ImageSrc()-equivalent at all). Word's JS API doesn't expose
      // a page number for a shape, so alt text / shape name / ordinal
      // position is the best available way to point the user at which one.
      // body.shapes holds every picture-type shape regardless of wrap —
      // inline pictures show up here too, so textWrap.type (not collection
      // membership) is what actually distinguishes floating from inline.
      const floatingImages = shapes.items.filter(
        (shape) => shape.type === Word.ShapeType.picture && shape.textWrap.type !== Word.ShapeTextWrapType.inline,
      );

      return floatingImages.map((shape, index) => {
        const label = shape.altTextDescription?.trim() || shape.name || `image ${index + 1}`;
        return `Floating image ${index + 1} of ${floatingImages.length} ("${label}") won't be captured — select it and use Picture Format → Wrap Text → In Line with Text, then save again.`;
      });
    });
  },

  async highlightFlaggedTerms() {
    return Word.run(async (context) => {
      let totalFound = 0;

      // No decoration-layer equivalent in Word, so this is a one-shot scan
      // rather than something we could run live on every keystroke — running
      // search+highlight+insertComment per keystroke would spam duplicate
      // comments and thrash document formatting.
      for (const { term, comment } of FLAGGED_TERMS) {
        const results = context.document.body.search(term, { matchCase: false, matchWholeWord: true });
        results.load('items');
        await context.sync();

        results.items.forEach((range) => {
          range.font.highlightColor = 'Yellow';
          range.insertComment(comment);
        });
        totalFound += results.items.length;
        await context.sync();
      }

      return totalFound;
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
