import type { JSONContent } from '@tiptap/react';
import { MOCK_POSTS, type FeedPost } from './mockPosts';

// Demo-only, in-memory article store. Tracks which post the editor is
// currently working on (set by ArticlesPanel when an article is loaded or a
// new one is started), and lets the web "Publish Article" flow push the
// editor's current content back into MOCK_POSTS so the Article Feed
// reflects the edit. Everything here is in memory — MOCK_POSTS is a plain
// module-level array we mutate in place, so a browser refresh re-initializes
// the module and resets all of it, which is exactly the intended behavior
// (no persistence, clean slate each session).

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

// Publishes the editor's current content into MOCK_POSTS. Updates the post
// currently being edited if one is loaded; otherwise creates a new post
// (the "New Article" workflow) and makes it the current one, so a repeat
// publish updates rather than duplicates. Returns the affected post.
export function publishArticleToFeed(json: JSONContent, html: string): FeedPost {
  const existing = currentArticleId ? MOCK_POSTS.find((p) => p.id === currentArticleId) : undefined;

  if (existing) {
    existing.json = json;
    existing.html = html;
    existing.title = deriveTitle(json);
    return existing;
  }

  const newPost: FeedPost = {
    id: `post-${Date.now()}`,
    title: deriveTitle(json),
    author: 'You',
    publishedAt: new Date().toISOString().slice(0, 10),
    category: 'Draft',
    html,
    json,
  };
  // Top of the feed so it's immediately visible.
  MOCK_POSTS.unshift(newPost);
  currentArticleId = newPost.id;
  return newPost;
}
