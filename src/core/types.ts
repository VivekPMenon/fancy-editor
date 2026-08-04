import type { JSONContent } from '@tiptap/react';

export interface DocumentAdapter {
  getContentHtml(): Promise<string>;
  getSelectionText(): Promise<string>;
  replaceSelection(text: string): Promise<void>;
  /** Inserts formatted HTML (headings, lists, bold, etc.) at the current cursor/selection. */
  insertHtml(html: string): Promise<void>;
  /** Subscribes to the document's plain-text content as it changes. Returns an unsubscribe function. */
  onContentChange(callback: (text: string) => void): () => void;
  /** Pre-save validation warnings (host-specific — e.g. Word floating images that getContentHtml can't capture). */
  getContentWarnings(): Promise<string[]>;
}

export interface StoredArticle {
  html: string;
  json: JSONContent;
  savedAt: string;
}
