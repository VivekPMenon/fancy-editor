import type { Article } from './articles';

export function generateMockSummary(articles: Article[]): string {
  const items = articles
    .map((article) => `<li><strong>${article.title}</strong> — ${article.snippet}</li>`)
    .join('');

  return (
    `<h2>Combined Summary</h2>` +
    `<p>This summary was generated from ${articles.length} selected article${articles.length === 1 ? '' : 's'}:</p>` +
    `<ul>${items}</ul>` +
    `<p>Taken together, these pieces point to a market environment balancing near-term caution with structural tailwinds worth monitoring closely.</p>`
  );
}
