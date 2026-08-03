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
      <html>

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
@font-face
	{font-family:"Calibri Light";
	panose-1:2 15 3 2 2 2 4 3 2 4;}
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
h1
	{mso-style-link:"Heading 1 Char";
	margin-top:16.0pt;
	margin-right:0in;
	margin-bottom:2.0pt;
	margin-left:0in;
	text-align:justify;
	text-justify:inter-ideograph;
	line-height:105%;
	page-break-after:avoid;
	font-size:14.0pt;
	font-family:"Calibri Light",sans-serif;
	text-transform:uppercase;
	letter-spacing:.2pt;
	font-weight:bold;}
a:link, span.MsoHyperlink
	{color:#0563C1;
	text-decoration:underline;}
p.MsoListParagraph, li.MsoListParagraph, div.MsoListParagraph
	{margin-top:0in;
	margin-right:0in;
	margin-bottom:8.0pt;
	margin-left:.5in;
	text-align:justify;
	text-justify:inter-ideograph;
	line-height:105%;
	font-size:11.0pt;
	font-family:"Calibri",sans-serif;}
p.MsoListParagraphCxSpFirst, li.MsoListParagraphCxSpFirst, div.MsoListParagraphCxSpFirst
	{margin-top:0in;
	margin-right:0in;
	margin-bottom:0in;
	margin-left:.5in;
	text-align:justify;
	text-justify:inter-ideograph;
	line-height:105%;
	font-size:11.0pt;
	font-family:"Calibri",sans-serif;}
p.MsoListParagraphCxSpMiddle, li.MsoListParagraphCxSpMiddle, div.MsoListParagraphCxSpMiddle
	{margin-top:0in;
	margin-right:0in;
	margin-bottom:0in;
	margin-left:.5in;
	text-align:justify;
	text-justify:inter-ideograph;
	line-height:105%;
	font-size:11.0pt;
	font-family:"Calibri",sans-serif;}
p.MsoListParagraphCxSpLast, li.MsoListParagraphCxSpLast, div.MsoListParagraphCxSpLast
	{margin-top:0in;
	margin-right:0in;
	margin-bottom:8.0pt;
	margin-left:.5in;
	text-align:justify;
	text-justify:inter-ideograph;
	line-height:105%;
	font-size:11.0pt;
	font-family:"Calibri",sans-serif;}
span.Heading1Char
	{mso-style-name:"Heading 1 Char";
	mso-style-link:"Heading 1";
	font-family:"Calibri Light",sans-serif;
	text-transform:uppercase;
	letter-spacing:.2pt;
	font-weight:bold;}
.MsoChpDefault
	{font-family:"Calibri",sans-serif;}
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

<h1 id=BE901453-7DC2-43AF-942C-FB072A20EC05>Story of Me. Kunchi…</h1>

<p id=9BE71B6A-97C8-47F9-BE81-F67D9F9E2FD9 class=MsoNormal>&nbsp;</p>

<p id=9CA2BC05-E329-401E-A018-276D149EC812 class=MsoNormal>My name is Kunchi. I
am a nice beautiful girl. My friend's name is Michael. I go to bubble school
everyday. I like to play in the school park. I also like to play in the Swing
and the ship park. My favorite food is Ice cream.</p>

<p id=23D39F10-5E10-4B37-B3F2-3001021711E3 class=MsoNormal>My favorite food is
nja mummum. Below are the list of fruits I like</p>

<p id=80A763AF-7F00-4950-A841-5D1220F899AC class=MsoListParagraphCxSpFirst
style='margin-left:38.25pt;text-indent:-.25in'>1.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Strawberry</p>

<p id=83D2C5FC-5D47-44EC-B3E8-8B194D1E5F68 class=MsoListParagraphCxSpMiddle
style='margin-left:38.25pt;text-indent:-.25in'>2.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Banana</p>

<p id=610D1BB6-51AA-479A-99F7-E3CD96210FF8 class=MsoListParagraphCxSpMiddle
style='margin-left:38.25pt;text-indent:-.25in'>3.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Apple</p>

<p id=CC427361-F4D8-4DFD-8FC8-2E72F39210B4 class=MsoListParagraphCxSpMiddle
style='margin-left:38.25pt;text-indent:-.25in'>4.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Cucumber</p>

<p id=B0CDDBF0-0D0D-4AD8-9478-3C71BF75B4AC class=MsoListParagraphCxSpMiddle
style='margin-left:38.25pt;text-indent:-.25in'>5.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Blueberry</p>

<p id=C89105D3-D58E-4FCE-9D55-185D39BF67E0 class=MsoListParagraphCxSpLast
style='margin-left:38.25pt;text-indent:-.25in'>6.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Tomato</p>

<p id=BD66814C-6D44-442C-BA12-7D171A813AB1 class=MsoNormal><br>
<br>
</p>

<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0
 style='border-collapse:collapse;border:none'>
 <tr>
  <td width=208 valign=top style='width:155.8pt;border:solid windowtext 1.0pt;
  padding:0in 5.4pt 0in 5.4pt'>
  <p id=3A12436E-DCC1-4224-8465-92C97FBCC47A class=MsoNormal style='margin-bottom:
  0in;line-height:normal'><b>Member Name</b></p>
  </td>
  <td width=208 valign=top style='width:155.85pt;border:solid windowtext 1.0pt;
  border-left:none;padding:0in 5.4pt 0in 5.4pt'>
  <p id=465436E8-218A-43B6-81AE-A44CD561EEA9 class=MsoNormal style='margin-bottom:
  0in;line-height:normal'><b>Relationship</b></p>
  </td>
  <td width=208 valign=top style='width:155.85pt;border:solid windowtext 1.0pt;
  border-left:none;padding:0in 5.4pt 0in 5.4pt'>
  <p id=00D75778-FC08-4F70-8605-5561D3923940 class=MsoNormal style='margin-bottom:
  0in;line-height:normal'><b>Like Level</b></p>
  </td>
 </tr>
 <tr>
  <td width=208 valign=top style='width:155.8pt;border:solid windowtext 1.0pt;
  border-top:none;padding:0in 5.4pt 0in 5.4pt'>
  <p id=DE8B4178-3E08-4EC9-B80D-AC8EAA9752FC class=MsoNormal style='margin-bottom:
  0in;line-height:normal'>Vivek Menon</p>
  </td>
  <td width=208 valign=top style='width:155.85pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:0in 5.4pt 0in 5.4pt'>
  <p id=6BD68329-5274-4E10-B173-7E807A026886 class=MsoNormal style='margin-bottom:
  0in;line-height:normal'>Daddy</p>
  </td>
  <td width=208 valign=top style='width:155.85pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:0in 5.4pt 0in 5.4pt'>
  <p id=BD929026-914F-4CFD-9D51-F1E10CCEC560 class=MsoNormal style='margin-bottom:
  0in;line-height:normal'>10</p>
  </td>
 </tr>
 <tr>
  <td width=208 valign=top style='width:155.8pt;border:solid windowtext 1.0pt;
  border-top:none;padding:0in 5.4pt 0in 5.4pt'>
  <p id=AA3F7BDE-B29F-4BD1-8B53-F8470D72065A class=MsoNormal style='margin-bottom:
  0in;line-height:normal'>Sukanya Thekke </p>
  </td>
  <td width=208 valign=top style='width:155.85pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:0in 5.4pt 0in 5.4pt'>
  <p id=1E34E66F-73E6-4D96-B7A5-425D4ECF7408 class=MsoNormal style='margin-bottom:
  0in;line-height:normal'>Mommy</p>
  </td>
  <td width=208 valign=top style='width:155.85pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:0in 5.4pt 0in 5.4pt'>
  <p id=86A6BF2B-9E2C-4C11-A985-E20C07D70B01 class=MsoNormal style='margin-bottom:
  0in;line-height:normal'>10</p>
  </td>
 </tr>
 <tr>
  <td width=208 valign=top style='width:155.8pt;border:solid windowtext 1.0pt;
  border-top:none;padding:0in 5.4pt 0in 5.4pt'>
  <p id=FFDDFCC0-B018-4192-B398-8F726899060D class=MsoNormal style='margin-bottom:
  0in;line-height:normal'>Chakki</p>
  </td>
  <td width=208 valign=top style='width:155.85pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:0in 5.4pt 0in 5.4pt'>
  <p id=8521ECDF-B9AA-48AF-AD8D-C9A16866B57C class=MsoNormal style='margin-bottom:
  0in;line-height:normal'>Sister</p>
  </td>
  <td width=208 valign=top style='width:155.85pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:0in 5.4pt 0in 5.4pt'>
  <p id=43337BEC-3FF0-464C-9E76-0810E5D3FF18 class=MsoNormal style='margin-bottom:
  0in;line-height:normal'>10</p>
  </td>
 </tr>
 <tr>
  <td width=208 valign=top style='width:155.8pt;border:solid windowtext 1.0pt;
  border-top:none;padding:0in 5.4pt 0in 5.4pt'>
  <p id=160F30BD-EF22-4F9C-9FCF-0689D8B73640 class=MsoNormal style='margin-bottom:
  0in;line-height:normal'>Kunchi</p>
  </td>
  <td width=208 valign=top style='width:155.85pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:0in 5.4pt 0in 5.4pt'>
  <p id=5D9BBDB2-52B8-4372-AF54-B63078B7B075 class=MsoNormal style='margin-bottom:
  0in;line-height:normal'>Me</p>
  </td>
  <td width=208 valign=top style='width:155.85pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:0in 5.4pt 0in 5.4pt'>
  <p id=8B087F6B-AD39-4C89-B4B2-E6D3B2E4334F class=MsoNormal style='margin-bottom:
  0in;line-height:normal'>10</p>
  </td>
 </tr>
</table>

<p id=A733BB31-786E-4923-A182-3AAA08CA2587 class=MsoNormal align=left
style='text-align:left'>Here is my nice little photograph taken from <b>Kerala</b><br>
<br>
<img width=624 height=616 id="Picture 9"
src="~WRS%7b0FDD8BDC-95A3-48A8-9D42-DD48E20F7EDC%7d_files/image001.jpg"><br>
<br>
<br>
Here is a video that l like <br>
<br>
<a href="https://www.youtube.com/watch?v=VclGFzAtALQ&amp;t=2558s">https://www.youtube.com/watch?v=VclGFzAtALQ&amp;t=2558s</a></p>

<p id=4387EF4E-4BA2-4FF5-8D14-082C79A9D236 class=MsoNormal>&nbsp;</p>

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
