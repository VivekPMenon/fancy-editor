import { useMemo, useState } from 'react';
import { useEditor, EditorContent } from '@tiptap/react';
import { EDITOR_EXTENSIONS } from '../../core/tiptap-utils/editorExtensions';
import { PublisherPlugin } from '../publisher-panel/PublisherPlugin';
import { createTiptapAdapter } from '../../adapters/tiptapAdapter';
import { EditorToolbar } from '../editor-toolbar/EditorToolbar';
import { EditorContextMenu } from '../editor-toolbar/EditorContextMenu';
import { SectionInsertMenu } from './SectionInsertMenu';
import { BLANK_ARTICLE_HTML } from '../articles/ArticlesPanel';
import { ArticleStatusBadge } from '../articles/ArticleStatusBadge';
import type { ArticleStatus, FeedPost } from '../../core/mockPosts';

interface FancyEditorTabProps {
  /** The post being edited, or null for a brand-new, unsaved article. */
  article: FeedPost | null;
  onBack: () => void;
}

interface ArticleMeta {
  status: ArticleStatus;
  title: string;
  lastUpdatedBy: string | null;
  lastUpdatedAt: string | null;
}

function metaFromArticle(article: FeedPost | null): ArticleMeta {
  return article
    ? { status: article.status, title: article.title, lastUpdatedBy: article.lastUpdatedBy, lastUpdatedAt: article.lastUpdatedAt }
    : { status: 'draft', title: 'Untitled Article', lastUpdatedBy: null, lastUpdatedAt: null };
}

export function FancyEditorTab({ article, onBack }: FancyEditorTabProps) {
  const editor = useEditor({
    extensions: EDITOR_EXTENSIONS,
    // JSON is our source of truth for a stored article — passing it directly
    // skips an unnecessary HTML round-trip (see htmlJsonConversion.ts).
    content: article ? article.json : BLANK_ARTICLE_HTML,
  });

  // Snapshot of the article's identity fields, shown in the info bar above
  // the editor so the user always knows which article they're in. Updated
  // (not live-typed) whenever Publish/Save as Draft succeeds — see
  // onArticleSaved below — since the point is "which article/status is this",
  // not a live-typing preview.
  const [articleMeta, setArticleMeta] = useState<ArticleMeta>(() => metaFromArticle(article));

  // Memoized so the adapter's identity stays stable across re-renders;
  // PublisherPanel resubscribes to onContentChange whenever it changes.
  const adapter = useMemo(() => (editor ? createTiptapAdapter(editor) : null), [editor]);

  if (!editor || !adapter) {
    return null;
  }

  return (
    <div className="app-layout">
      <div className="tiptap-editor-column">
        <div className="article-info-bar">
          <button type="button" className="article-info-back" onClick={onBack}>
            ← My Articles
          </button>
          <span className="article-info-divider" aria-hidden="true" />
          <ArticleStatusBadge status={articleMeta.status} />
          <span className="article-info-title">{articleMeta.title}</span>
          <span className="article-info-meta">
            {articleMeta.lastUpdatedAt
              ? `Last updated by ${articleMeta.lastUpdatedBy} · ${new Date(articleMeta.lastUpdatedAt).toLocaleString()}`
              : 'Not saved yet'}
          </span>
        </div>
        <EditorToolbar editor={editor} />
        <SectionInsertMenu editor={editor} />
        <EditorContent className="tiptap-editor" editor={editor} />
        <EditorContextMenu editor={editor} />
      </div>
      <div className="app-right-rail">
        <PublisherPlugin
          adapter={adapter}
          hostLabel="Web"
          onArticleSaved={(post) => setArticleMeta(metaFromArticle(post))}
        />
      </div>
    </div>
  );
}
