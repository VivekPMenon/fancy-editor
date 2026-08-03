import type { JSONContent } from '@tiptap/react';

export interface DocumentAdapter {
  getContentHtml(): Promise<string>;
  getSelectionText(): Promise<string>;
  replaceSelection(text: string): Promise<void>;
  /** Inserts formatted HTML (headings, lists, bold, etc.) at the current cursor/selection. */
  insertHtml(html: string): Promise<void>;
  /** Subscribes to the document's plain-text content as it changes. Returns an unsubscribe function. */
  onContentChange(callback: (text: string) => void): () => void;
}

export interface StoredArticle {
  html: string;
  json: JSONContent;
  savedAt: string;
}
