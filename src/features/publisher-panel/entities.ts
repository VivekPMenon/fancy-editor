export interface EntityInfo {
  symbol: string;
  exchangeSuffix: string;
  name: string;
  rating: 'Buy' | 'Neutral' | 'Sell';
  priceTarget: string;
}

const ENTITY_DB: EntityInfo[] = [
  { symbol: 'AAPL', exchangeSuffix: 'US', name: 'Apple Inc', rating: 'Neutral', priceTarget: 'PT 296.00 USD' },
  { symbol: 'MSFT', exchangeSuffix: 'US', name: 'Microsoft Corp', rating: 'Buy', priceTarget: 'PT 520.00 USD' },
  // { symbol: 'META', exchangeSuffix: 'US', name: 'Meta Platforms Inc', rating: 'Buy', priceTarget: 'PT 715.00 USD' },
];

export function detectEntities(text: string): EntityInfo[] {
  if (!text.trim()) {
    return [];
  }
  return ENTITY_DB;
}
