import type { JSONContent } from '@tiptap/react';

export interface DocumentAdapter {
  getContentHtml(): Promise<string>;
  getSelectionText(): Promise<string>;
  replaceSelection(text: string): Promise<void>;
  /** Inserts formatted HTML (headings, lists, bold, etc.) at the current cursor/selection. */
  insertHtml(html: string): Promise<void>;
  /** Replaces the entire document content — used to start a blank article. */
  setContentHtml(html: string): Promise<void>;
  /**
   * Replaces the entire document content from Tiptap JSON — our source of
   * truth for a stored article. Prefer this over setContentHtml(json→html)
   * when loading a stored article: on Tiptap, it skips an unnecessary
   * HTML round-trip; on Word, which has no JSON concept, it still converts
   * internally, but that conversion is the adapter's concern, not the
   * caller's.
   */
  setContentJson(json: JSONContent): Promise<void>;
  /** Subscribes to the document's plain-text content as it changes. Returns an unsubscribe function. */
  onContentChange(callback: (text: string) => void): () => void;
  /** Pre-save validation warnings (host-specific — e.g. Word floating images that getContentHtml can't capture). */
  getContentWarnings(): Promise<string[]>;
  /**
   * Scans for the hardcoded flagged-term list (core/flaggedTerms.ts) and highlights matches.
   * On Tiptap, highlighting is already live via a decoration plugin, so this is a no-op that
   * returns 0. On Word, there's no decoration-layer equivalent, so this does a one-shot scan:
   * applies a real highlight color and inserts a native Word comment per match.
   */
  highlightFlaggedTerms(): Promise<number>;
}

export interface StoredArticle {
  html: string;
  json: JSONContent;
  savedAt: string;
}
