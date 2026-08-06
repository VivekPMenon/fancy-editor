import { Extension } from '@tiptap/core';
import { preprocessWordHtml } from './wordHtmlPreprocessing';

// Copy-pasting from Word gives the browser the same "filtered HTML" shape as
// body.getHtml()/Save As Web Page — a full mini-document with a <style>
// block driving heading color/font-size/font-family, Word's linked-
// thumbnail video markup, and fake numbered/bulleted paragraphs instead of
// real lists. Tiptap's default paste path parses that HTML straight into
// the schema with none of htmlToTiptapJson's preprocessing, so pasted
// content would silently lose all of it. Running the same
// preprocessWordHtml() steps here keeps paste and "load a saved article" on
// equal footing.
//
// Known gap: if the copied selection contains a floating shape (or a
// comment anchor), Word can fail to populate `text/html` on the clipboard
// at all — confirmed by tracing prosemirror-view's paste handling: with no
// html string, it takes the plain-text path and never calls this hook, so
// the *entire* paste silently loses all formatting, not just the shape.
// There's no code-level fix — same root limitation as officeAdapter's Shape
// gap (§2/§9b in docs/word-integration-notes.md), just hit via the clipboard
// instead of Office.js. See docs/word-integration-notes.md §10 for details.
export const WordPaste = Extension.create({
  name: 'wordPaste',
  transformPastedHTML(html) {
    // Large/complex documents (footnotes, nested tables, embedded fonts)
    // could in principle trip up juice's CSS parsing on some malformed
    // stylesheet. If that happens, fall back to the original HTML rather
    // than losing the paste entirely — imperfect formatting beats nothing.
    try {
      return preprocessWordHtml(html);
    } catch (err) {
      console.error('WordPaste: preprocessing failed, pasting unprocessed HTML', err);
      return html;
    }
  },
});
