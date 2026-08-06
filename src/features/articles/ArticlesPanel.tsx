import { useState } from 'react';
import type { JSONContent } from '@tiptap/react';
import type { DocumentAdapter } from '../../core/types';
import { MOCK_POSTS } from '../../core/mockPosts';
import './ArticlesPanel.css';

export const BLANK_ARTICLE_HTML = '<p>Start writing your article here…</p>';

interface ArticlesPanelProps {
  adapter: DocumentAdapter;
}

export function ArticlesPanel({ adapter }: ArticlesPanelProps) {
  const [selectedId, setSelectedId] = useState<string | null>(null);

  function handleCreateNew() {
    setSelectedId(null);
    void adapter.setContentHtml(BLANK_ARTICLE_HTML);
  }

  function handleSelectArticle(id: string, json: JSONContent) {
    setSelectedId(id);
    // JSON is the source of truth — setContentJson lets each adapter decide
    // how to consume it (Tiptap natively, Word by converting internally)
    // instead of always pre-converting to HTML here.
    void adapter.setContentJson(json);
  }

  return (
    <div className="articles-panel">
      <ul className="articles-list">
        {MOCK_POSTS.map((post) => (
          <li key={post.id}>
            <button
              type="button"
              className={`articles-item ${selectedId === post.id ? 'active' : ''}`}
              onClick={() => handleSelectArticle(post.id, post.json)}
            >
              <span className="articles-item-title">{post.title}</span>
              <span className="articles-item-meta">
                {post.author} · updated {new Date(post.publishedAt).toLocaleDateString()}
              </span>
            </button>
          </li>
        ))}
      </ul>
      <button type="button" className="publisher-panel-generate" onClick={handleCreateNew}>
        + New Article
      </button>
    </div>
  );
}
