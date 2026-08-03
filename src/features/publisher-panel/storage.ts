import type { JSONContent } from '@tiptap/react';
import type { StoredArticle } from '../../core/types';

const STORAGE_KEY = 'fancy-editor:article';

export function saveArticle(html: string, json: JSONContent): StoredArticle {
  const record: StoredArticle = { html, json, savedAt: new Date().toISOString() };
  localStorage.setItem(STORAGE_KEY, JSON.stringify(record));
  return record;
}

export function loadArticle(): StoredArticle | null {
  const raw = localStorage.getItem(STORAGE_KEY);
  return raw ? (JSON.parse(raw) as StoredArticle) : null;
}
