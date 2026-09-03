import type { ArticleStatus } from '../../core/mockPosts';
import './ArticleStatusBadge.css';

// Small shared pill used both on My Articles' list cards and the editor's
// article info bar, so "what does Draft/Published look like" stays defined
// in exactly one place.
export function ArticleStatusBadge({ status }: { status: ArticleStatus }) {
  return (
    <span className={`article-status-badge article-status-badge--${status}`}>
      {status === 'published' ? 'Published' : 'Draft'}
    </span>
  );
}
