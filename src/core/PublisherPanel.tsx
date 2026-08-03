import { useEffect, useState } from 'react';
import type { DocumentAdapter } from './types';
import type { Article } from './articles';
import { ARTICLES } from './articles';
import { generateMockSummary } from './summary';
import { fancifyText } from './transforms';
import { loadArticle, saveArticle } from './storage';
import './PublisherPanel.css';

interface PublisherPanelProps {
  adapter: DocumentAdapter;
  hostLabel: string;
}

export function PublisherPanel({ adapter, hostLabel }: PublisherPanelProps) {
  const [savedAt, setSavedAt] = useState<string | null>(() => loadArticle()?.savedAt ?? null);
  const [status, setStatus] = useState('');
  const [liveText, setLiveText] = useState('');

  const [query, setQuery] = useState('');
  const [selectedArticles, setSelectedArticles] = useState<Article[]>([]);
  const [summaryMessage, setSummaryMessage] = useState('');

  useEffect(() => adapter.onContentChange(setLiveText), [adapter]);

  const suggestions = query.trim()
    ? ARTICLES.filter(
        (article) =>
          !selectedArticles.some((selected) => selected.id === article.id) &&
          article.title.toLowerCase().includes(query.trim().toLowerCase()),
      ).slice(0, 5)
    : [];

  function handleSelectArticle(article: Article) {
    setSelectedArticles((prev) => [...prev, article]);
    setQuery('');
  }

  function handleRemoveArticle(id: string) {
    setSelectedArticles((prev) => prev.filter((article) => article.id !== id));
  }

  async function handleGenerateSummary() {
    if (selectedArticles.length === 0) {
      return;
    }
    try {
      setSummaryMessage('');
      const html = generateMockSummary(selectedArticles);
      await adapter.insertHtml(html);
      setSummaryMessage('Your summary has been generated and added to the document.');
      setSelectedArticles([]);
    } catch (err) {
      setSummaryMessage(`Summary generation failed: ${(err as Error).message}`);
    }
  }

  async function handleSave() {
    try {
      const html = await adapter.getContentHtml();
      const record = saveArticle(html);
      setSavedAt(record.savedAt);
      setStatus(`Saved ${html.length} chars of HTML.`);
    } catch (err) {
      setStatus(`Save failed: ${(err as Error).message}`);
    }
  }

  async function handleFancify() {
    try {
      const selection = await adapter.getSelectionText();
      if (!selection) {
        setStatus('Select some text first.');
        return;
      }
      await adapter.replaceSelection(fancifyText(selection));
      setStatus(`Replaced "${selection}" with a fancy word.`);
    } catch (err) {
      setStatus(`Fancify failed: ${(err as Error).message}`);
    }
  }

  return (
    <div className="publisher-panel">
      <h2>Publisher Plus ({hostLabel})</h2>

      <section className="publisher-panel-section">
        <h3 className="publisher-panel-section-title">Article summary</h3>

        <div className="publisher-panel-search">
          <input
            type="text"
            value={query}
            onChange={(event) => setQuery(event.target.value)}
            placeholder="Search articles to add…"
          />
          {suggestions.length > 0 && (
            <ul className="publisher-panel-suggestions">
              {suggestions.map((article) => (
                <li key={article.id}>
                  <button type="button" onClick={() => handleSelectArticle(article)}>
                    {article.title}
                  </button>
                </li>
              ))}
            </ul>
          )}
        </div>

        {selectedArticles.length > 0 && (
          <ul className="publisher-panel-chips">
            {selectedArticles.map((article) => (
              <li key={article.id} className="publisher-panel-chip">
                <span>{article.title}</span>
                <button
                  type="button"
                  onClick={() => handleRemoveArticle(article.id)}
                  aria-label={`Remove ${article.title}`}
                >
                  ×
                </button>
              </li>
            ))}
          </ul>
        )}

        <button
          type="button"
          className="publisher-panel-generate"
          onClick={handleGenerateSummary}
          disabled={selectedArticles.length === 0}
        >
          Generate AI Summary
        </button>

        {summaryMessage && <p className="publisher-panel-success">✓ {summaryMessage}</p>}
      </section>

      <section className="publisher-panel-section">
        <h3 className="publisher-panel-section-title">Document tools</h3>
        <button type="button" onClick={handleFancify}>
          Fancify selection
        </button>
        <button type="button" onClick={handleSave}>
          Save article as HTML
        </button>
        {savedAt && <p className="publisher-panel-meta">Last saved: {new Date(savedAt).toLocaleString()}</p>}
        {status && <p className="publisher-panel-status">{status}</p>}
      </section>

      <div className="publisher-panel-preview">
        <p className="publisher-panel-preview-label">Live uppercase preview</p>
        <p className="publisher-panel-preview-text">{liveText ? liveText.toUpperCase() : '…'}</p>
      </div>
    </div>
  );
}
