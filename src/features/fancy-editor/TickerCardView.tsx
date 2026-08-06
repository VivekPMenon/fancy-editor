import { useEffect, useState } from 'react';
import { NodeViewWrapper, type ReactNodeViewProps } from '@tiptap/react';
import { ArrowSyncRegular } from '@fluentui/react-icons';
import { TICKER_DATABASE, getTickerInfo, type TickerInfo } from '../../core/tickerDatabase';
import './TickerCardView.css';

const AUTO_REFRESH_MS = 3000;

function jitterPrice(ticker: TickerInfo): string {
  const base = Number(ticker.priceTarget.replace(/[^0-9.]/g, ''));
  const jittered = base * (1 + (Math.random() - 0.5) * 0.04);
  return jittered.toFixed(2);
}

// The "Word can't do this" node — a real, interactive React component living
// inside the document flow (not a pasted screenshot). Switching the symbol
// persists to the document (a real node attribute, saved/loaded like any
// other content); the price jitter is deliberately ephemeral — it never
// touches the document — just to prove the card is genuinely alive, not a
// static image, while it's open in the editor.
export function TickerCardView({ node, updateAttributes, selected }: ReactNodeViewProps) {
  const symbol = node.attrs.symbol as string;
  const ticker = getTickerInfo(symbol);
  const [livePrice, setLivePrice] = useState<string | null>(null);
  const [lastUpdated, setLastUpdated] = useState<Date | null>(null);

  // Auto-refresh while the card is mounted — getTickerInfo() returns the
  // same object reference for a given symbol (TICKER_DATABASE.find()), so
  // this only re-fires when the symbol actually changes, not on every
  // unrelated re-render.
  useEffect(() => {
    if (!ticker) {
      return;
    }
    const tick = () => {
      setLivePrice(jitterPrice(ticker));
      setLastUpdated(new Date());
    };
    tick();
    const intervalId = setInterval(tick, AUTO_REFRESH_MS);
    return () => clearInterval(intervalId);
  }, [ticker]);

  function handleRefresh() {
    if (ticker) {
      setLivePrice(jitterPrice(ticker));
      setLastUpdated(new Date());
    }
  }

  return (
    <NodeViewWrapper className={`ticker-card-view ${selected ? 'selected' : ''}`} data-drag-handle>
      <div className="ticker-card-view-header">
        <select
          className="ticker-card-view-symbol-select"
          value={symbol}
          onChange={(event) => {
            // No need to reset livePrice/lastUpdated here — the auto-refresh
            // effect re-fires as soon as `ticker` changes and immediately
            // ticks a fresh value for the new symbol.
            updateAttributes({ symbol: event.target.value });
          }}
          contentEditable={false}
        >
          {TICKER_DATABASE.map((t) => (
            <option key={t.symbol} value={t.symbol}>
              {t.symbol}
            </option>
          ))}
        </select>
        {ticker && (
          <span className={`ticker-card-view-rating ticker-card-view-rating--${ticker.rating.toLowerCase()}`}>
            {ticker.rating}
          </span>
        )}
        <button type="button" className="ticker-card-view-refresh" title="Refresh live price" onClick={handleRefresh}>
          <ArrowSyncRegular />
        </button>
      </div>

      {ticker ? (
        <>
          <p className="ticker-card-view-name">{ticker.name}</p>
          <div className="ticker-card-view-footer">
            <span className="ticker-card-view-target">{ticker.priceTarget}</span>
            {livePrice && (
              <span className="ticker-card-view-live">
                Live: ${livePrice} <span className="ticker-card-view-live-dot" />
              </span>
            )}
          </div>
          {lastUpdated && (
            <p className="ticker-card-view-updated">Updated {lastUpdated.toLocaleTimeString()}</p>
          )}
        </>
      ) : (
        <p className="ticker-card-view-name">No data for {symbol}</p>
      )}
    </NodeViewWrapper>
  );
}
