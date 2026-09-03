import { useMemo, useState } from 'react';
import { MOCK_POSTS, type FeedPost } from '../../core/mockPosts';
import { tiptapJsonToPlainText } from '../../core/tiptap-utils/htmlJsonConversion';
import { ArticleStatusBadge } from './ArticleStatusBadge';
import './MyArticlesLanding.css';

interface MyArticlesLandingProps {
  onCreate: () => void;
  onOpen: (post: FeedPost) => void;
}

function formatUpdated(post: FeedPost): string {
  return new Date(post.lastUpdatedAt).toLocaleString(undefined, {
    dateStyle: 'medium',
    timeStyle: 'short',
  });
}

// App landing view — every article the user has authored, newest first (that's
// MOCK_POSTS' own order; publishing/saving unshifts). A simple client-side
// filter over title + body text stands in for real search; "History" is a
// placeholder per-article action — version tracking isn't implemented yet, so
// it's disabled rather than wired to anything.
export function MyArticlesLanding({ onCreate, onOpen }: MyArticlesLandingProps) {
  const [query, setQuery] = useState('');

  const filtered = useMemo(() => {
    const q = query.trim().toLowerCase();
    if (!q) {
      return MOCK_POSTS;
    }
    return MOCK_POSTS.filter((post) => {
      if (post.title.toLowerCase().includes(q)) {
        return true;
      }
      return tiptapJsonToPlainText(post.json).toLowerCase().includes(q);
    });
  }, [query]);

  return (
    <div className="my-articles-landing">
      <div className="my-articles-toolbar">
        <input
          type="search"
          className="my-articles-search"
          placeholder="Search your articles by title or content…"
          value={query}
          onChange={(event) => setQuery(event.target.value)}
          aria-label="Search your articles"
        />
        <button type="button" className="my-articles-create" onClick={onCreate}>
          + Create Article
        </button>
      </div>

      {filtered.length === 0 ? (
        <p className="my-articles-empty">
          {query ? `No articles match "${query}".` : 'No articles yet — create your first one.'}
        </p>
      ) : (
        <ul className="my-articles-list">
          {filtered.map((post) => (
            <li key={post.id} className="my-articles-card">
              <button type="button" className="my-articles-card-open" onClick={() => onOpen(post)}>
                <div className="my-articles-card-top">
                  <ArticleStatusBadge status={post.status} />
                  <span className="my-articles-card-title">{post.title}</span>
                </div>
                <p className="my-articles-card-meta">
                  Created by {post.createdBy} · Last updated by {post.lastUpdatedBy} · {formatUpdated(post)}
                </p>
              </button>
              <button
                type="button"
                className="my-articles-history"
                disabled
                title="Version history — coming soon"
              >
                History
              </button>
            </li>
          ))}
        </ul>
      )}
    </div>
  );
}
