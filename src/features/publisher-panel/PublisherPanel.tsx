import { useEffect, useMemo, useState } from 'react';
import type { DocumentAdapter } from '../../core/types';
import type { Article } from './articles';
import { ARTICLES } from './articles';
import { generateMockSummary } from './summary';
// import { fancifyText } from './transforms';
import { detectEntities } from './entities';
import { AI_TAGS } from './tags';
import { convertOoxmlStringToTiptapJson } from '../../core/tiptap-utils/ooxml/ooxmlToTiptapJson';
import { publishArticleToFeed } from '../../core/articleStore';
import './PublisherPanel.css';

interface PublisherPanelProps {
  adapter: DocumentAdapter;
}

// Word exposes getOoxml(); the Tiptap web adapter doesn't. That single
// capability is the host discriminator — Word gets the Data Format choice
// (publish to clipboard as HTML / OOXML / JSON), web just publishes the
// edited article straight into the in-memory feed.
type DataFormat = 'html' | 'ooxml' | 'json';

export function PublisherPanel({ adapter }: PublisherPanelProps) {
  const isWordHost = Boolean(adapter.getContentOoxml);

  const [status, setStatus] = useState('');
  const [warnings, setWarnings] = useState<string[]>([]);
  const [liveText, setLiveText] = useState('');
  const [lastPublishedAt, setLastPublishedAt] = useState<string | null>(null);

  const [query, setQuery] = useState('');
  const [selectedArticles, setSelectedArticles] = useState<Article[]>([]);
  const [summaryMessage, setSummaryMessage] = useState('');

  // Word host only: JSON (default) publishes via the OOXML->Tiptap-JSON
  // route, HTML via getHtml(), OOXML as the raw captured XML.
  const [dataFormat, setDataFormat] = useState<DataFormat>('json');

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

  // Word: capture the document in the selected Data Format and copy it to
  // the clipboard — JSON goes through the OOXML->Tiptap-JSON converter,
  // HTML through getHtml(), OOXML is the raw captured flat-OPC XML.
  async function publishFromWord() {
    if (dataFormat === 'html') {
      // The floating-image warning only applies to the HTML path — getHtml()
      // genuinely can't capture a floating image's bytes. The OOXML path
      // (JSON/OOXML formats) does capture them, so those don't warn.
      setWarnings(await adapter.getContentWarnings());
      const html = await adapter.getContentHtml();
      await copyToClipboard(html, `Published — copied ${html.length} chars of HTML to clipboard.`);
      return;
    }

    setWarnings([]);

    // OOXML and JSON both start from getOoxml(); if Word itself fails to
    // produce it (office-js#998/#5394), surface that plainly.
    const ooxml = await adapter.getContentOoxml!();
    if (dataFormat === 'ooxml') {
      await copyToClipboard(ooxml, `Published — copied ${ooxml.length} chars of raw OOXML to clipboard.`);
      return;
    }

    const json = convertOoxmlStringToTiptapJson(ooxml);
    const blockCount = json.content?.length ?? 0;
    await copyToClipboard(
      JSON.stringify(json, null, 2),
      `Published — converted OOXML to Tiptap JSON (${blockCount} blocks) and copied to clipboard.`,
    );
  }

  // Web: push the editor's current content straight into the in-memory feed
  // (updating the loaded article, or creating a new one) so the Article Feed
  // reflects the edit. No clipboard, no format choice — the web editor is
  // already Tiptap JSON.
  async function publishFromWeb() {
    const json = adapter.getContentJson ? await adapter.getContentJson() : undefined;
    if (!json) {
      setStatus('Publish is only available in the web editor for now.');
      return;
    }
    const html = await adapter.getContentHtml();
    const post = publishArticleToFeed(json, html);
    setStatus(`Published "${post.title}" to the Article Feed.`);
  }

  async function copyToClipboard(text: string, successMessage: string) {
    try {
      await navigator.clipboard.writeText(text);
      setStatus(successMessage);
    } catch {
      setStatus('Published, but copying to the clipboard failed — try again or copy manually.');
    }
  }

  async function handlePublish() {
    try {
      if (isWordHost) {
        await publishFromWord();
      } else {
        await publishFromWeb();
      }
      setLastPublishedAt(new Date().toISOString());
    } catch (err) {
      setStatus(`Publish failed: ${(err as Error).message}`);
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
    <>
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

        {isWordHost && (
          <div className="publisher-panel-data-format">
            <span className="publisher-panel-data-format-label">Data Format</span>
            <div className="publisher-panel-segment" role="radiogroup" aria-label="Data Format">
              {(['html', 'json', 'ooxml'] as const).map((format) => (
                <button
                  key={format}
                  type="button"
                  role="radio"
                  aria-checked={dataFormat === format}
                  className={dataFormat === format ? 'active' : ''}
                  onClick={() => setDataFormat(format)}
                >
                  {format.toUpperCase()}
                </button>
              ))}
            </div>
          </div>
        )}

        <button type="button" className="publisher-panel-publish" onClick={handlePublish}>
          Publish Article
        </button>

        <button type="button" onClick={handleScanFlaggedTerms}>
          Scan for flagged terms
        </button>
        {lastPublishedAt && (
          <p className="publisher-panel-meta">Last published: {new Date(lastPublishedAt).toLocaleString()}</p>
        )}
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
    </>
  );
}
