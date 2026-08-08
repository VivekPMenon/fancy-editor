export interface TagRule {
  label: string;
  // A tag is applied when ANY of these appears as a case-insensitive
  // substring of the article text — a deliberately simple keyword "includes"
  // check standing in for a real classifier, enough to demo tags that track
  // what's actually written rather than a fixed list.
  keywords: string[];
}

export const TAG_RULES: TagRule[] = [
  {
    label: 'Equities',
    keywords: ['equit', 'stock', 'share', 'nasdaq', 'ticker', 'aapl', 'apple', 'msft', 'microsoft'],
  },
  { label: 'USD', keywords: ['usd', 'dollar', 'united states', 'u.s.'] },
  { label: 'GBP', keywords: ['gbp', 'pound', 'sterling'] },
  {
    label: 'Research',
    keywords: ['research', 'analyst', 'analysis', 'coverage', 'rating', 'price target', 'forecast', 'outlook', 'valuation'],
  },
];

export function detectTags(text: string): string[] {
  const haystack = text.toLowerCase();
  if (!haystack.trim()) {
    return [];
  }
  return TAG_RULES.filter((rule) => rule.keywords.some((keyword) => haystack.includes(keyword))).map(
    (rule) => rule.label,
  );
}
