import juice from 'juice';
import { generateHTML, generateJSON } from '@tiptap/html';
import type { JSONContent } from '@tiptap/react';
import { EDITOR_EXTENSIONS } from './editorExtensions';
import { resolveWordVideoEmbeds } from './wordVideoEmbed';

// Converts captured HTML (e.g. Word's `body.getHtml()` output) into the same
// ProseMirror JSON schema the live Tiptap editor produces via `getJSON()`.
// This is also what normalizes it: anything outside EDITOR_EXTENSIONS'
// node/mark types (Mso classes, stray <style> blocks, font-face rules, …)
// simply can't be represented and gets dropped during parsing.
//
// Word often expresses formatting (esp. heading color/font-size/font-family)
// as a <style> block rule (e.g. `h1 { color: ... }`) rather than an inline
// style on the element — and Tiptap's parseHTML only ever reads an element's
// own inline `style` attribute, never the stylesheet. juice() resolves the
// real CSS cascade and inlines the result onto each matching element first,
// so that formatting has something for parseHTML to actually find.
export function htmlToTiptapJson(html: string): JSONContent {
  const resolvedHtml = juice(resolveWordVideoEmbeds(html));
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
