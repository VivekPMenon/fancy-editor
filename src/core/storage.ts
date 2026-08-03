import type { StoredArticle } from './types';

const STORAGE_KEY = 'publisher-plus:article';

export function saveArticle(html: string): StoredArticle {
  const record: StoredArticle = { html, savedAt: new Date().toISOString() };
  localStorage.setItem(STORAGE_KEY, JSON.stringify(record));
  return record;
}

export function loadArticle(): StoredArticle | null {
  const raw = localStorage.getItem(STORAGE_KEY);
  return raw ? (JSON.parse(raw) as StoredArticle) : null;
}
