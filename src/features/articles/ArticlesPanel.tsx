import { useState } from 'react';
import type { JSONContent } from '@tiptap/react';
import { MOCK_POSTS } from '../../core/mockPosts';
import './ArticlesPanel.css';

interface ArticlesPanelProps {
  onSelectArticle: (json: JSONContent) => void;
  onCreateNew: () => void;
}

export function ArticlesPanel({ onSelectArticle, onCreateNew }: ArticlesPanelProps) {
  const [selectedId, setSelectedId] = useState<string | null>(null);

  function handleCreateNew() {
    setSelectedId(null);
    onCreateNew();
  }

  return (
    <div className="articles-panel">
      <div className="articles-panel-header">
        <h2>My Articles</h2>
        <button type="button" className="articles-new-button" onClick={handleCreateNew}>
          + New
        </button>
      </div>
      <ul className="articles-list">
        {MOCK_POSTS.map((post) => (
          <li key={post.id}>
            <button
              type="button"
              className={`articles-item ${selectedId === post.id ? 'active' : ''}`}
              onClick={() => {
                setSelectedId(post.id);
                onSelectArticle(post.json);
              }}
            >
              <span className="articles-item-title">{post.title}</span>
              <span className="articles-item-meta">
                {post.author} · updated {new Date(post.publishedAt).toLocaleDateString()}
              </span>
            </button>
          </li>
        ))}
      </ul>
    </div>
  );
}
