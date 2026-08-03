import type { JSONContent } from '@tiptap/react';
import { htmlToTiptapJson } from './htmlJsonConversion';

export interface FeedPost {
  id: string;
  title: string;
  author: string;
  publishedAt: string;
  category: string;
  html: string;
  json: JSONContent;
}

type RawFeedPost = Omit<FeedPost, 'json'>;

// Placeholder bodies below approximate what Word's `getHtml()` export tends to
// produce (headings, bold/italic/underline, lists, tables, links). Once
// you've used "Save article as HTML" and copied real captured HTML, replace
// the `html` field of an entry with the pasted content to see how it renders.
// `json` is derived from `html` below rather than hand-written — it's what
// News format renders and what the JSON would look like fed back into the
// editor, normalized to whatever EDITOR_EXTENSIONS' schema can represent.
const RAW_POSTS: RawFeedPost[] = [
  {
    id: 'post-1',
    title: 'Global Supply Chains Show Signs of Stabilization',
    author: 'Priya Nair',
    publishedAt: '2026-07-28',
    category: 'Trade',
    html: `<html>

<head>
<meta http-equiv=Content-Type content="text/html; charset=windows-1252">
<meta name=Generator content="Microsoft Word 15 (filtered)">
<style>
<!--
 /* Font Definitions */
 @font-face
	{font-family:Wingdings;
	panose-1:5 0 0 0 0 0 0 0 0 0;}
@font-face
	{font-family:"Cambria Math";
	panose-1:2 4 5 3 5 4 6 3 2 4;}
@font-face
	{font-family:Calibri;
	panose-1:2 15 5 2 2 2 4 3 2 4;}
 /* Style Definitions */
 p.MsoNormal, li.MsoNormal, div.MsoNormal
	{margin-top:0in;
	margin-right:0in;
	margin-bottom:8.0pt;
	margin-left:0in;
	text-align:justify;
	text-justify:inter-ideograph;
	line-height:105%;
	font-size:11.0pt;
	font-family:"Calibri",sans-serif;}
.MsoPapDefault
	{margin-bottom:8.0pt;
	text-align:justify;
	text-justify:inter-ideograph;
	line-height:105%;}
@page WordSection1
	{size:8.5in 11.0in;
	margin:1.0in 1.0in 1.0in 1.0in;}
div.WordSection1
	{page:WordSection1;}
 /* List Definitions */
 ol
	{margin-bottom:0in;}
ul
	{margin-bottom:0in;}
-->
</style>

</head>

<body lang=EN-US link="#0563C1" vlink="#954F72" style='word-wrap:break-word'>

<div class=WordSection1>

<p id=41B66C78-0EED-4EAB-BC14-FD576975A9F0 class=MsoNormal><b>Donna&nbsp;Robbins&nbsp;</b></p>

<p id=9786B70D-37EB-4019-9D89-D151A850970C class=MsoNormal>Accountant&nbsp;</p>

<p id=A062D723-9E89-4D67-AFD9-59470395D3FB class=MsoNormal><b>313.555.0100&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;
donna@example.com&nbsp;&nbsp;&nbsp;&nbsp;|&nbsp;&nbsp;&nbsp;&nbsp;www.greatsiteaddress.com&nbsp;</b></p>

<p id=D2D49BE8-57F5-4A76-A98D-DD21527D171D class=MsoNormal><b>&nbsp;</b></p>

<p id=2B6C9ABB-C36D-4139-B5DB-98F85851CE84 class=MsoNormal><b>4 April
20XX&nbsp;<br>
Hiring Manager&nbsp;<br>
ESG&nbsp;Financial Services&nbsp;<br>
5678 Mountain Drive&nbsp;<br>
Denver, CO 12345&nbsp;</b></p>

<p id=10EFC261-E7A1-48B6-9C4E-0892C62C89C0 class=MsoNormal>&nbsp;</p>

<p id=2C671257-40D5-4689-81FC-CEF4440B7318 class=MsoNormal><b>DEAR LILLI
ALLIK,&nbsp;</b></p>

<p id=3BA2CFC0-F51F-4BE9-A68B-34B758439A4D class=MsoNormal>&nbsp;</p>

<p id=5D278D92-191F-49D5-8834-3E65E63EDCD5 class=MsoNormal>I am writing to
express my interest in the accountant position at ESG Financial Services. With
a Bachelor of Science in Accounting, a minor in Business Administration, and
several years of hands-on experience in both public accounting and private
sectors, I am confident in my ability to provide valuable financial solutions
to your organization. My&nbsp;expertise&nbsp;spans income tax preparation,
financial statement preparation, and general ledger accounting, with a strong
background in GAAP compliance.&nbsp;</p>

<p id=AB109CCA-0E44-4238-99CF-479E900B8530 class=MsoNormal>Currently, as an
Accountant at Trey Research, I support a wide range of clients, from
individuals to businesses, delivering&nbsp;accurate&nbsp;and efficient
accounting services. My attention to detail, ownership mentality, and
collaborative approach have enabled me to&nbsp;identify&nbsp;and implement
strategic solutions that enhance financial processes and support business
growth. Additionally, my&nbsp;proficiency&nbsp;in Microsoft NAV Dynamics,
QuickBooks, and cashflow management ensures that I can handle the financial
complexities of a diverse client base.&nbsp;</p>

<p id=04807CA2-00EF-488B-8FDB-BDC11E491E41 class=MsoNormal>I am particularly
drawn to ESG Financial Services because of its commitment to supporting women
and minority owned businesses.&nbsp;I would be thrilled to bring my analytical
skills, organizational abilities, and passion for accounting to contribute to your
team’s success.&nbsp;</p>

<p id=8399AED4-9AA9-45A2-ACA7-0EA82854B4E6 class=MsoNormal>Thank you for
considering my application. I look forward to the opportunity to discuss how my
qualifications align with the needs of your organization.&nbsp;</p>

<p id=66D921EC-07A2-42D8-9526-AC6BE79CBA76 class=MsoNormal>&nbsp;</p>

<p id=C251B117-3ACF-43D5-B896-1F619005E892 class=MsoNormal>Sincerely,&nbsp;<br>
Donna Robbins&nbsp;</p>

<p id=A733BB31-786E-4923-A182-3AAA08CA2587 class=MsoNormal>&nbsp;</p>

<p id=6D86174F-3858-434E-AF77-4A0DFC203F2B class=MsoNormal>&nbsp;</p>

</div>

</body>

</html>
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

export const MOCK_POSTS: FeedPost[] = RAW_POSTS.map((post) => ({
  ...post,
  json: htmlToTiptapJson(post.html),
}));
