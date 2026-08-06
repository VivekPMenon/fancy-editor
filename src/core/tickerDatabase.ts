export interface TickerInfo {
  symbol: string;
  exchangeSuffix: string;
  name: string;
  rating: 'Buy' | 'Neutral' | 'Sell';
  priceTarget: string;
}

// Shared hardcoded ticker data — used by the Identified Entities panel
// (features/publisher-panel/entities.ts) and the live Ticker Card node
// (core/tiptap-utils/tickerCardExtension.ts). One list so both stay
// consistent, same reasoning as core/flaggedTerms.ts.
export const TICKER_DATABASE: TickerInfo[] = [
  { symbol: 'AAPL', exchangeSuffix: 'US', name: 'Apple Inc', rating: 'Neutral', priceTarget: 'PT 296.00 USD' },
  // { symbol: 'MSFT', exchangeSuffix: 'US', name: 'Microsoft Corp', rating: 'Buy', priceTarget: 'PT 520.00 USD' },
  // { symbol: 'META', exchangeSuffix: 'US', name: 'Meta Platforms Inc', rating: 'Buy', priceTarget: 'PT 715.00 USD' },
];

export function getTickerInfo(symbol: string): TickerInfo | undefined {
  return TICKER_DATABASE.find((ticker) => ticker.symbol === symbol.toUpperCase());
}
