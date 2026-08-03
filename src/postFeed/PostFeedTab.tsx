import { useState } from 'react';
import type { SyntheticEvent } from 'react';
import { MOCK_POSTS, type FeedPost } from '../core/mockPosts';
import { tiptapJsonToHtml, tiptapJsonToPlainText } from '../core/htmlJsonConversion';
import './PostFeedTab.css';

type FeedFormat = 'news' | 'word';

const EXCERPT_LENGTH = 220;
const WORD_FRAME_MIN_HEIGHT = 200;

function estimateReadMinutes(plainText: string): number {
  const words = plainText.split(' ').filter(Boolean).length;
  return Math.max(1, Math.round(words / 200));
}

// "Word format" wants to preserve full fidelity, including that <style>
// block — so instead of injecting it into the page, render the captured
// document inside an isolated iframe where its stylesheet can't leak out.
function WordFormatFrame({ html }: { html: string }) {
  const [height, setHeight] = useState(WORD_FRAME_MIN_HEIGHT);

  function handleLoad(event: SyntheticEvent<HTMLIFrameElement>) {
    const doc = event.currentTarget.contentDocument;
    if (doc?.body) {
      setHeight(Math.max(WORD_FRAME_MIN_HEIGHT, doc.body.scrollHeight + 24));
    }
  }

  return (
    <iframe
      className="feed-post-frame"
      title="Captured Word HTML"
      srcDoc={html}
      style={{ height }}
      onLoad={handleLoad}
    />
  );
}

function FeedPostCard({ post, format }: { post: FeedPost; format: FeedFormat }) {
  const [expanded, setExpanded] = useState(false);
  const plainText = tiptapJsonToPlainText(post.json).replace(/\s+/g, ' ').trim();
  const excerpt = plainText.length > EXCERPT_LENGTH ? `${plainText.slice(0, EXCERPT_LENGTH)}…` : plainText;

  return (
    <div className='articles'>
      <article className={`feed-post feed-post-${format}`}>
        <span className="feed-post-category">{post.category}</span>
        <button type="button" className="feed-post-title" onClick={() => setExpanded((e) => !e)}>
          {post.title}
        </button>
        <p className="feed-post-meta">
          {post.author} · {new Date(post.publishedAt).toLocaleDateString()} · {estimateReadMinutes(plainText)} min read
        </p>

        {expanded ? (
          format === 'word' ? (
            <WordFormatFrame html={post.html} />
          ) : (
            <div
              className="feed-post-body feed-post-body-news"
              dangerouslySetInnerHTML={{ __html: tiptapJsonToHtml(post.json) }}
            />
          )
        ) : (
          <p className="feed-post-excerpt">{excerpt}</p>
        )}

        <button type="button" className="feed-post-toggle" onClick={() => setExpanded((e) => !e)}>
          {expanded ? 'Show less' : 'Show more'}
        </button>
      </article>
    </div>
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
