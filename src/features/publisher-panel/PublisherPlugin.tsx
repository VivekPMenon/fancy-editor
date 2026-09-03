import { useState } from 'react';
import type { DocumentAdapter } from '../../core/types';
import type { FeedPost } from '../../core/mockPosts';
import { PublisherPanel } from './PublisherPanel';
import { ArticlesPanel } from '../articles/ArticlesPanel';
import './PublisherPanel.css';
import './PublisherPlugin.css';

type PluginTab = 'tools' | 'articles';

interface PublisherPluginProps {
  adapter: DocumentAdapter;
  hostLabel: string;
  /** Web only: fires after Publish/Save as Draft/DOCX-import succeeds, so a
   *  host showing article identity (e.g. FancyEditorTab's info bar) can stay
   *  in sync with the saved status/title/timestamp. */
  onArticleSaved?: (post: FeedPost) => void;
}

export function PublisherPlugin({ adapter, hostLabel, onArticleSaved }: PublisherPluginProps) {
  const [activeTab, setActiveTab] = useState<PluginTab>('tools');

  return (
    <div className="publisher-panel">
      <h2>My Custom Plugin ({hostLabel})</h2>

      <div className="publisher-plugin-tabs" role="tablist">
        <button
          type="button"
          role="tab"
          aria-selected={activeTab === 'tools'}
          className={`publisher-plugin-tab ${activeTab === 'tools' ? 'active' : ''}`}
          onClick={() => setActiveTab('tools')}
        >
          Toolbar
        </button>
        <button
          type="button"
          role="tab"
          aria-selected={activeTab === 'articles'}
          className={`publisher-plugin-tab ${activeTab === 'articles' ? 'active' : ''}`}
          onClick={() => setActiveTab('articles')}
        >
          My Articles
        </button>
      </div>

      {activeTab === 'tools' ? (
        <PublisherPanel adapter={adapter} onArticleSaved={onArticleSaved} />
      ) : (
        <ArticlesPanel adapter={adapter} />
      )}
    </div>
  );
}
