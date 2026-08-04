import { useState } from 'react';
import type { JSONContent } from '@tiptap/react';
import type { DocumentAdapter } from '../../core/types';
import { MOCK_POSTS } from '../../core/mockPosts';
import { tiptapJsonToHtml } from '../../core/tiptap-utils/htmlJsonConversion';
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
    // Proves the stored Tiptap JSON — our source of truth — round-trips
    // through HTML and still renders correctly once fed into the host
    // document, not just back into a Tiptap editor.
    void adapter.setContentHtml(tiptapJsonToHtml(json));
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
