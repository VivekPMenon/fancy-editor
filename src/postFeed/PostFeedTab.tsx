import { useState } from 'react';
import { MOCK_POSTS, type FeedPost } from './mockPosts';
import './PostFeedTab.css';

type FeedFormat = 'news' | 'word';

const EXCERPT_LENGTH = 220;

function stripHtml(html: string): string {
  return html
    .replace(/<[^>]+>/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function estimateReadMinutes(plainText: string): number {
  const words = plainText.split(' ').filter(Boolean).length;
  return Math.max(1, Math.round(words / 200));
}

// "News format" strips Word's inline style/class attributes so our own CSS
// fully controls the look, instead of whatever Word happened to embed.
function sanitizeForNewsFormat(html: string): string {
  const container = document.createElement('div');
  container.innerHTML = html;
  container.querySelectorAll('[style]').forEach((el) => el.removeAttribute('style'));
  container.querySelectorAll('[class]').forEach((el) => el.removeAttribute('class'));
  return container.innerHTML;
}

function FeedPostCard({ post, format }: { post: FeedPost; format: FeedFormat }) {
  const [expanded, setExpanded] = useState(false);
  const plainText = stripHtml(post.html);
  const excerpt = plainText.length > EXCERPT_LENGTH ? `${plainText.slice(0, EXCERPT_LENGTH)}…` : plainText;

  return (
    <article className={`feed-post feed-post-${format}`}>
      <span className="feed-post-category">{post.category}</span>
      <button type="button" className="feed-post-title" onClick={() => setExpanded((e) => !e)}>
        {post.title}
      </button>
      <p className="feed-post-meta">
        {post.author} · {new Date(post.publishedAt).toLocaleDateString()} · {estimateReadMinutes(plainText)} min read
      </p>

      {expanded ? (
        <div
          className={`feed-post-body feed-post-body-${format}`}
          dangerouslySetInnerHTML={{ __html: format === 'news' ? sanitizeForNewsFormat(post.html) : post.html }}
        />
      ) : (
        <p className="feed-post-excerpt">{excerpt}</p>
      )}

      <button type="button" className="feed-post-toggle" onClick={() => setExpanded((e) => !e)}>
        {expanded ? 'Show less' : 'Show more'}
      </button>
    </article>
  );
}

export function PostFeedTab() {
  const [format, setFormat] = useState<FeedFormat>('news');

  return (
    <div className="post-feed">
      <div className="post-feed-list">
        <div className="post-feed-toolbar">
          <div className="post-feed-format-toggle">
            <button type="button" className={format === 'news' ? 'active' : ''} onClick={() => setFormat('news')}>
              News format
            </button>
            <button type="button" className={format === 'word' ? 'active' : ''} onClick={() => setFormat('word')}>
              Word format
            </button>
          </div>
        </div>

        {MOCK_POSTS.map((post) => (
          <FeedPostCard key={post.id} post={post} format={format} />
        ))}
      </div>
    </div>
  );
}
