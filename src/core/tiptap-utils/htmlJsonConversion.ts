import { generateHTML, generateJSON } from '@tiptap/html';
import type { JSONContent } from '@tiptap/react';
import { EDITOR_EXTENSIONS } from './editorExtensions';
import { preprocessWordHtml } from './wordHtmlPreprocessing';

// Converts captured HTML (e.g. Word's `body.getHtml()` output) into the same
// ProseMirror JSON schema the live Tiptap editor produces via `getJSON()`.
// This is also what normalizes it: anything outside EDITOR_EXTENSIONS'
// node/mark types (Mso classes, stray <style> blocks, font-face rules, …)
// simply can't be represented and gets dropped during parsing.
//
// preprocessWordHtml resolves the CSS cascade Word expresses via stylesheet
// classes (Tiptap's parseHTML only ever reads inline styles), rewrites
// linked-thumbnail video markup into a real embed, and reconstructs real
// <ol>/<ul>/<li> from Word's fake numbered/bulleted paragraphs — all things
// Tiptap's parser can't see or do on its own.
export function htmlToTiptapJson(html: string): JSONContent {
  const resolvedHtml = preprocessWordHtml(html);
  return generateJSON(resolvedHtml, EDITOR_EXTENSIONS);
}

// Renders normalized Tiptap JSON back to clean HTML, driven entirely by our
// own schema/CSS rather than whatever markup produced the original JSON.
export function tiptapJsonToHtml(json: JSONContent): string {
  return generateHTML(json, EDITOR_EXTENSIONS);
}

// Node types after which we insert a separator, so text from adjacent blocks
// (paragraphs, list items, table cells, …) doesn't run together.
const BLOCK_NODE_TYPES = new Set([
  'paragraph',
  'heading',
  'listItem',
  'tableRow',
  'tableCell',
  'tableHeader',
  'blockquote',
  'codeBlock',
]);

// Extracts plain text straight from Tiptap JSON — text nodes already hold
// real decoded characters (a `&nbsp;` in the source HTML becomes an actual
// U+00A0 space here, not the literal string "&nbsp;"), so this avoids the
// HTML-entity artifacts a regex-based tag-strip on raw HTML would leave in.
export function tiptapJsonToPlainText(node: JSONContent): string {
  let text = node.text ?? '';
  if (node.content) {
    text += node.content.map(tiptapJsonToPlainText).join('');
  }
  if (node.type && BLOCK_NODE_TYPES.has(node.type)) {
    text += ' ';
  }
  return text;
}
