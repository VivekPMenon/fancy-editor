import { TICKER_DATABASE, type TickerInfo } from '../../core/tickerDatabase';

export type EntityInfo = TickerInfo;

export function detectEntities(text: string): EntityInfo[] {
  if (!text.trim()) {
    return [];
  }
  return TICKER_DATABASE;
}
