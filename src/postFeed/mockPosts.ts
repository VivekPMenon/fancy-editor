export interface FeedPost {
  id: string;
  title: string;
  author: string;
  publishedAt: string;
  category: string;
  html: string;
}

// Placeholder bodies below approximate what Word's `getHtml()` export tends to
// produce (headings, bold/italic/underline, lists, tables, links). Once
// you've used "Save article as HTML" and copied real captured HTML, replace
// the `html` field of an entry with the pasted content to see how it renders.
export const MOCK_POSTS: FeedPost[] = [
  {
    id: 'post-1',
    title: 'Global Supply Chains Show Signs of Stabilization',
    author: 'Priya Nair',
    publishedAt: '2026-07-28',
    category: 'Trade',
    html: `
      <h2>Global Supply Chains Show Signs of Stabilization</h2>
      <p>Freight rates and lead times have <strong>eased over the past two quarters</strong> as manufacturers diversify sourcing and rebuild inventory buffers.</p>
      <p><em>Analysts note</em> that the recovery remains <u>uneven across regions</u>, with Southeast Asia leading the improvement.</p>
      <ul>
        <li>Ocean freight rates down 18% year-over-year</li>
        <li>Average lead times back to pre-2021 levels</li>
        <li>Inventory-to-sales ratios normalizing across retail</li>
      </ul>
      <p>Read the full methodology on the <a href="https://example.com/supply-chain-index">Supply Chain Index</a> page.</p>
    `,
  },
  {
    id: 'post-2',
    title: 'Renewable Energy Investment Hits Record High',
    author: 'Daniel Osei',
    publishedAt: '2026-07-25',
    category: 'Energy',
    html: `
      <h2>Renewable Energy Investment Hits Record High</h2>
      <p>Global investment in solar and wind projects <strong>surpassed prior-year levels</strong>, driven by falling technology costs and new policy incentives.</p>
      <table>
        <thead>
          <tr><th>Region</th><th>2025 Investment</th><th>2026 Investment</th></tr>
        </thead>
        <tbody>
          <tr><td>Asia-Pacific</td><td>$412B</td><td>$488B</td></tr>
          <tr><td>Europe</td><td>$298B</td><td>$331B</td></tr>
          <tr><td>Americas</td><td>$205B</td><td>$249B</td></tr>
        </tbody>
      </table>
      <p>Battery storage co-located with generation projects was the <em>fastest-growing</em> segment.</p>
    `,
  },
  {
    id: 'post-3',
    title: 'Central Banks Signal Cautious Approach to Rate Cuts',
    author: 'Marta Kowalski',
    publishedAt: '2026-07-21',
    category: 'Rates',
    html: `
      <h2>Central Banks Signal Cautious Approach to Rate Cuts</h2>
      <p>Policymakers reiterated a <strong>data-dependent stance</strong>, citing persistent services inflation as the key obstacle to faster easing.</p>
      <ol>
        <li>Headline inflation continues to trend toward target</li>
        <li>Services inflation remains sticky above 4%</li>
        <li>Labor markets are gradually loosening but remain tight</li>
      </ol>
      <p style="color: #666666;"><em>Markets are now pricing in two cuts before year-end, down from four at the start of the quarter.</em></p>
    `,
  },
  {
    id: 'post-4',
    title: 'AI Regulation Frameworks Take Shape Across Major Markets',
    author: 'Kenji Watanabe',
    publishedAt: '2026-07-18',
    category: 'Policy',
    html: `
      <h2>AI Regulation Frameworks Take Shape Across Major Markets</h2>
      <p>New guidelines aim to <strong>balance innovation with accountability</strong>, focusing on transparency requirements for high-risk automated systems.</p>
      <p>Key provisions include:</p>
      <ul>
        <li><strong>Disclosure requirements</strong> for AI-generated content</li>
        <li><u>Mandatory risk assessments</u> for high-impact deployments</li>
        <li>Independent audits for systems used in <em>financial services and healthcare</em></li>
      </ul>
      <p>Compliance deadlines vary by jurisdiction; see the <a href="https://example.com/ai-policy-tracker">policy tracker</a> for details.</p>
    `,
  },
  {
    id: 'post-5',
    title: 'Semiconductor Demand Rebounds on AI Infrastructure Buildout',
    author: 'Sofia Bianchi',
    publishedAt: '2026-07-14',
    category: 'Technology',
    html: `
      <h2>Semiconductor Demand Rebounds on AI Infrastructure Buildout</h2>
      <p>Chipmakers report <strong>improving order books</strong> as hyperscalers ramp capital expenditure on data center capacity.</p>
      <p>Foundry utilization rates are back above 90% for advanced nodes, with lead times <em>stretching into 2027</em> for the most in-demand packaging technologies.</p>
      <blockquote>"We're seeing the strongest bookings visibility in three years," one industry executive noted on a recent earnings call.</blockquote>
      <p>Capital spending guidance was raised across the sector for the third consecutive quarter.</p>
    `,
  },
];
