import type { JSONContent } from '@tiptap/react';
import { MOCK_POSTS, type ArticleStatus, type FeedPost } from './mockPosts';
import { CURRENT_USER } from './currentUser';

// Demo-only, in-memory article store. Tracks which post the editor is
// currently working on (set by ArticlesPanel/MyArticlesLanding when an
// article is loaded or a new one is started), and lets the web "Publish"/
// "Save as Draft" actions push the editor's current content back into
// MOCK_POSTS so My Articles and the Article Feed reflect the edit.
// Everything here is in memory — MOCK_POSTS is a plain module-level array we
// mutate in place, so a browser refresh re-initializes the module and resets
// all of it, which is exactly the intended behavior (no persistence, clean
// slate each session).

let currentArticleId: string | null = null;

export function setCurrentArticleId(id: string | null): void {
  currentArticleId = id;
}

export function getCurrentArticleId(): string | null {
  return currentArticleId;
}

// First heading's text, else first paragraph's text (trimmed/truncated),
// else a fallback — just enough to give a newly published article a sensible
// name in the feed and My Articles list.
function deriveTitle(json: JSONContent): string {
  const blocks = json.content ?? [];
  const firstHeading = blocks.find((b) => b.type === 'heading');
  const firstParagraph = blocks.find((b) => b.type === 'paragraph' && (b.content?.length ?? 0) > 0);
  const source = firstHeading ?? firstParagraph;
  const text = (source?.content ?? [])
    .map((n) => (n.type === 'text' ? (n.text ?? '') : ''))
    .join('')
    .trim();
  if (!text) {
    return 'Untitled Article';
  }
  return text.length > 80 ? `${text.slice(0, 80)}…` : text;
}

// Saves the editor's current content into MOCK_POSTS at the given status.
// Updates the post currently being edited if one is loaded; otherwise
// creates a new one (the "Create Article" workflow) and makes it the
// current one, so a repeat save updates rather than duplicates. Shared by
// both the "Publish" and "Save as Draft" actions — they differ only in the
// status they pass in. Returns the affected post.
function saveArticle(json: JSONContent, html: string, status: ArticleStatus): FeedPost {
  const existing = currentArticleId ? MOCK_POSTS.find((p) => p.id === currentArticleId) : undefined;
  const now = new Date().toISOString();

  if (existing) {
    existing.json = json;
    existing.html = html;
    existing.title = deriveTitle(json);
    existing.status = status;
    existing.lastUpdatedBy = CURRENT_USER.name;
    existing.lastUpdatedAt = now;
    return existing;
  }

  const newPost: FeedPost = {
    id: `post-${Date.now()}`,
    title: deriveTitle(json),
    author: CURRENT_USER.name,
    publishedAt: now.slice(0, 10),
    category: 'General',
    html,
    json,
    status,
    createdBy: CURRENT_USER.name,
    lastUpdatedBy: CURRENT_USER.name,
    lastUpdatedAt: now,
  };
  // Top of the feed/list so it's immediately visible.
  MOCK_POSTS.unshift(newPost);
  currentArticleId = newPost.id;
  return newPost;
}

export function publishArticleToFeed(json: JSONContent, html: string): FeedPost {
  return saveArticle(json, html, 'published');
}

export function saveArticleAsDraft(json: JSONContent, html: string): FeedPost {
  return saveArticle(json, html, 'draft');
}
