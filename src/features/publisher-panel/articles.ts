export interface Article {
  id: string;
  title: string;
  snippet: string;
}

export const ARTICLES: Article[] = [
  {
    id: 'art-1',
    title: 'Global Supply Chains Show Signs of Stabilization',
    snippet:
      'Freight rates and lead times have eased over the past two quarters as manufacturers diversify sourcing and rebuild inventory buffers.',
  },
  {
    id: 'art-2',
    title: 'Renewable Energy Investment Hits Record High',
    snippet:
      'Global investment in solar and wind projects surpassed prior-year levels, driven by falling technology costs and new policy incentives.',
  },
  {
    id: 'art-3',
    title: 'Central Banks Signal Cautious Approach to Rate Cuts',
    snippet:
      'Policymakers reiterated a data-dependent stance, citing persistent services inflation as the key obstacle to faster easing.',
  },
  {
    id: 'art-4',
    title: 'AI Regulation Frameworks Take Shape Across Major Markets',
    snippet:
      'New guidelines aim to balance innovation with accountability, focusing on transparency requirements for high-risk automated systems.',
  },
  {
    id: 'art-5',
    title: 'Consumer Spending Resilient Despite Rate Pressures',
    snippet:
      'Retail sales data points to steady discretionary spending, though savings rates continue to decline from pandemic-era highs.',
  },
  {
    id: 'art-6',
    title: 'Semiconductor Demand Rebounds on AI Infrastructure Buildout',
    snippet:
      'Chipmakers report improving order books as hyperscalers ramp capital expenditure on data center capacity.',
  },
];
