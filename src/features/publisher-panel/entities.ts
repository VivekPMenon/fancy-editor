import { TICKER_DATABASE, type TickerInfo } from '../../core/tickerDatabase';

export type EntityInfo = TickerInfo;

// Surface an instrument only when the article actually mentions it — a
// case-insensitive substring match on the ticker symbol, the full company
// name, or the company's leading word (so "Apple" matches "Apple Inc"). Same
// simple "includes" approach as the AI tags; not a real NER model.
export function detectEntities(text: string): EntityInfo[] {
  const haystack = text.toLowerCase();
  if (!haystack.trim()) {
    return [];
  }
  return TICKER_DATABASE.filter((ticker) => {
    const leadingNameWord = ticker.name.split(' ')[0].toLowerCase();
    return (
      haystack.includes(ticker.symbol.toLowerCase()) ||
      haystack.includes(ticker.name.toLowerCase()) ||
      haystack.includes(leadingNameWord)
    );
  });
}
