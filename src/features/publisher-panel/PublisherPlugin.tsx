import { useState } from 'react';
import type { DocumentAdapter } from '../../core/types';
import { PublisherPanel } from './PublisherPanel';
import { ArticlesPanel } from '../articles/ArticlesPanel';
import './PublisherPanel.css';
import './PublisherPlugin.css';

type PluginTab = 'tools' | 'articles';

interface PublisherPluginProps {
  adapter: DocumentAdapter;
  hostLabel: string;
}

export function PublisherPlugin({ adapter, hostLabel }: PublisherPluginProps) {
  const [activeTab, setActiveTab] = useState<PluginTab>('tools');

  return (
    <div className="publisher-panel">
      <h2>Publisher Plugin ({hostLabel})</h2>

      <div className="publisher-plugin-tabs" role="tablist">
        <button
          type="button"
          role="tab"
          aria-selected={activeTab === 'tools'}
          className={`publisher-plugin-tab ${activeTab === 'tools' ? 'active' : ''}`}
          onClick={() => setActiveTab('tools')}
        >
          Article Tools
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

      {activeTab === 'tools' ? <PublisherPanel adapter={adapter} /> : <ArticlesPanel adapter={adapter} />}
    </div>
  );
}
