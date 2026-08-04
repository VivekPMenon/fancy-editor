import { useEffect, useMemo, useState } from 'react';
import type { DocumentAdapter } from '../../core/types';
import type { Article } from './articles';
import { ARTICLES } from './articles';
import { generateMockSummary } from './summary';
// import { fancifyText } from './transforms';
import { loadArticle, saveArticle } from './storage';
import { detectEntities } from './entities';
import { AI_TAGS } from './tags';
import { htmlToTiptapJson } from '../../core/tiptap-utils/htmlJsonConversion';
import './PublisherPanel.css';

interface PublisherPanelProps {
  adapter: DocumentAdapter;
  hostLabel: string;
}

export function PublisherPanel({ adapter, hostLabel }: PublisherPanelProps) {
  const [savedAt, setSavedAt] = useState<string | null>(() => loadArticle()?.savedAt ?? null);
  const [status, setStatus] = useState('');
  const [warnings, setWarnings] = useState<string[]>([]);
  const [liveText, setLiveText] = useState('');

  const [query, setQuery] = useState('');
  const [selectedArticles, setSelectedArticles] = useState<Article[]>([]);
  const [summaryMessage, setSummaryMessage] = useState('');

  useEffect(() => adapter.onContentChange(setLiveText), [adapter]);

  const identifiedEntities = useMemo(() => detectEntities(liveText), [liveText]);
  const aiTags = useMemo(() => (liveText.trim() ? AI_TAGS : []), [liveText]);

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

  async function saveCurrentArticle() {
    const contentWarnings = await adapter.getContentWarnings();
    setWarnings(contentWarnings);
    const html = await adapter.getContentHtml();
    const json = htmlToTiptapJson(html);
    const record = saveArticle(html, json);
    setSavedAt(record.savedAt);
    return { html, json };
  }

  async function handleSaveHtml() {
    try {
      const { html } = await saveCurrentArticle();
      try {
        await navigator.clipboard.writeText(html);
        setStatus(`Saved and copied ${html.length} chars of HTML to clipboard.`);
      } catch {
        setStatus(`Saved ${html.length} chars of HTML — clipboard copy failed, copy manually instead.`);
      }
    } catch (err) {
      setStatus(`Save failed: ${(err as Error).message}`);
    }
  }

  async function handleSaveJson() {
    try {
      const { html, json } = await saveCurrentArticle();
      try {
        await navigator.clipboard.writeText(JSON.stringify(json, null, 2));
        setStatus(`Saved (from ${html.length} chars of HTML) and copied JSON to clipboard.`);
      } catch {
        setStatus('Saved, but JSON clipboard copy failed — copy manually instead.');
      }
    } catch (err) {
      setStatus(`Save failed: ${(err as Error).message}`);
    }
  }

  async function handleScanFlaggedTerms() {
    try {
      const count = await adapter.highlightFlaggedTerms();
      setStatus(
        count > 0
          ? `Flagged ${count} term${count === 1 ? '' : 's'} — highlighted and commented in the document.`
          : 'No flagged terms found.',
      );
    } catch (err) {
      setStatus(`Scan failed: ${(err as Error).message}`);
    }
  }

  // async function handleFancify() {
  //   try {
  //     const selection = await adapter.getSelectionText();
  //     if (!selection) {
  //       setStatus('Select some text first.');
  //       return;
  //     }
  //     await adapter.replaceSelection(fancifyText(selection));
  //     setStatus(`Replaced "${selection}" with a fancy word.`);
  //   } catch (err) {
  //     setStatus(`Fancify failed: ${(err as Error).message}`);
  //   }
  // }

  return (
    <div className="publisher-panel">
      <h2>Publisher Plugin ({hostLabel})</h2>

      <section className="publisher-panel-section">
        <div className="publisher-panel-section-header">
          <h3 className="publisher-panel-section-title">AI Tagging</h3>
          <span className="publisher-panel-live-indicator" aria-hidden="true" />
        </div>
        {aiTags.length > 0 ? (
          <ul className="publisher-panel-tags">
            {aiTags.map((tag) => (
              <li key={tag} className="publisher-panel-tag">
                {tag}
              </li>
            ))}
          </ul>
        ) : (
          <p className="publisher-panel-empty-hint">Tags will appear here as you write.</p>
        )}
      </section>

      <section className="publisher-panel-section">
        <div className="publisher-panel-section-header">
          <h3 className="publisher-panel-section-title">Identified Entities</h3>
          <span className="publisher-panel-live-indicator" aria-hidden="true" />
        </div>
        {identifiedEntities.length > 0 ? (
          <ul className="publisher-panel-entities">
            {identifiedEntities.map((entity) => (
              <li key={entity.symbol} className="publisher-panel-entity-card">
                <div className="publisher-panel-entity-header">
                  <span className="publisher-panel-entity-symbol">
                    {entity.symbol} <span className="publisher-panel-entity-suffix">{entity.exchangeSuffix}</span>
                  </span>
                  <span
                    className={`publisher-panel-entity-rating publisher-panel-entity-rating--${entity.rating.toLowerCase()}`}
                  >
                    {entity.rating}
                  </span>
                </div>
                <p className="publisher-panel-entity-name">{entity.name}</p>
                <div className="publisher-panel-entity-footer">
                  <span className="publisher-panel-entity-target">{entity.priceTarget}</span>
                  <span className="publisher-panel-entity-check" title="Verified">
                    ✓
                  </span>
                </div>
              </li>
            ))}
          </ul>
        ) : (
          <p className="publisher-panel-empty-hint">
            No instruments identified yet
          </p>
        )}
      </section>

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
        {/* <button type="button" onClick={handleFancify}>
          Fancify selection
        </button> */}
        <button type="button" onClick={handleSaveHtml}>
          Save article as HTML
        </button>
        <button type="button" onClick={handleSaveJson}>
          Save article as JSON
        </button>
        <button type="button" onClick={handleScanFlaggedTerms}>
          Scan for flagged terms
        </button>
        {savedAt && <p className="publisher-panel-meta">Last saved: {new Date(savedAt).toLocaleString()}</p>}
        {status && <p className="publisher-panel-status">{status}</p>}
        {warnings.length > 0 && (
          <ul className="publisher-panel-warnings">
            {warnings.map((warning) => (
              <li key={warning}>⚠ {warning}</li>
            ))}
          </ul>
        )}
      </section>
{/*
      <div className="publisher-panel-preview">
        <p className="publisher-panel-preview-label">Live uppercase preview</p>
        <p className="publisher-panel-preview-text">{liveText ? liveText.toUpperCase() : '…'}</p>
      </div> */}
    </div>
  );
}
