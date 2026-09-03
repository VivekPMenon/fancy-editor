import { useState } from 'react';
import { AppHeader } from './AppHeader';
import { FancyEditorTab } from './features/fancy-editor/FancyEditorTab';
import { PostFeedTab } from './features/post-feed/PostFeedTab';
import { MyArticlesLanding } from './features/articles/MyArticlesLanding';
import { setCurrentArticleId } from './core/articleStore';
import type { FeedPost } from './core/mockPosts';
import './App.css';

type Tab = 'editor' | 'feed';
// The "Article Editor" tab has two sub-views: the My Articles landing list
// (shown by default — see the app-launch requirement this implements) and
// the actual Tiptap editor, entered by creating or opening an article.
type EditorView = 'list' | 'edit';

function App() {
  const [activeTab, setActiveTab] = useState<Tab>('editor');
  const [editorView, setEditorView] = useState<EditorView>('list');
  // null = a brand-new, unsaved article; otherwise the post being edited.
  const [editingArticle, setEditingArticle] = useState<FeedPost | null>(null);

  function handleCreateArticle() {
    setCurrentArticleId(null);
    setEditingArticle(null);
    setEditorView('edit');
  }

  function handleOpenArticle(post: FeedPost) {
    setCurrentArticleId(post.id);
    setEditingArticle(post);
    setEditorView('edit');
  }

  function handleBackToList() {
    setEditorView('list');
  }

  return (
    <div className="app-shell">
      <AppHeader />

      <header className="app-header">
        <nav className="app-tabs">
          <button
            type="button"
            className={activeTab === 'editor' ? 'active' : ''}
            onClick={() => setActiveTab('editor')}
          >
          Article Editor
          </button>
          <button type="button" className={activeTab === 'feed' ? 'active' : ''} onClick={() => setActiveTab('feed')}>
          Article Feed
          </button>
        </nav>
      </header>

      {/* The editor tab stays mounted (just hidden) when inactive, so the
          Tiptap document and PublisherPanel state survive switching tabs.
          Within it, the list<->edit sub-view swap DOES unmount/remount the
          editor (via FancyEditorTab's key below) — leaving an article without
          saving and reopening it is expected to reload from the saved copy,
          not resume unsaved changes; that's what Save as Draft is for. */}
      <div className={activeTab === 'editor' ? 'app-tab-panel' : 'app-tab-panel app-tab-panel-hidden'}>
        {editorView === 'list' ? (
          <MyArticlesLanding onCreate={handleCreateArticle} onOpen={handleOpenArticle} />
        ) : (
          <FancyEditorTab
            key={editingArticle?.id ?? 'new'}
            article={editingArticle}
            onBack={handleBackToList}
          />
        )}
      </div>
      {activeTab === 'feed' && <PostFeedTab />}
    </div>
  );
}

export default App;
