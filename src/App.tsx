import { useState } from 'react';
import { FancyEditorTab } from './FancyEditorTab';
import { PostFeedTab } from './postFeed/PostFeedTab';
import './App.css';

type Tab = 'editor' | 'feed';

function App() {
  const [activeTab, setActiveTab] = useState<Tab>('editor');

  return (
    <div className="app-shell">
      <header className="app-header">
        <nav className="app-tabs">
          <button
            type="button"
            className={activeTab === 'editor' ? 'active' : ''}
            onClick={() => setActiveTab('editor')}
          >
            Fancy Editor
          </button>
          <button type="button" className={activeTab === 'feed' ? 'active' : ''} onClick={() => setActiveTab('feed')}>
            Post Feed
          </button>
        </nav>
      </header>

      {/* The editor tab stays mounted (just hidden) when inactive, so the
          Tiptap document and PublisherPanel state survive switching tabs. */}
      <div className={activeTab === 'editor' ? 'app-tab-panel' : 'app-tab-panel app-tab-panel-hidden'}>
        <FancyEditorTab />
      </div>
      {activeTab === 'feed' && <PostFeedTab />}
    </div>
  );
}

export default App;
