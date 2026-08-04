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
    html: `
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="Generator" content="Microsoft Word 15 (filtered)">
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
	{font-family:Tahoma;
	panose-1:2 11 6 4 3 5 4 4 2 4;}
@font-face
	{font-family:"Lao UI";}
@font-face
	{font-family:Ubuntu;}
 /* Style Definitions */
 p.MsoNormal, li.MsoNormal, div.MsoNormal
	{margin:0in;
	text-indent:.3in;
	line-height:115%;
	font-size:12.0pt;
	font-family:"Ubuntu",sans-serif;}
h1
	{mso-style-link:"Heading 1 Char";
	margin-top:24.0pt;
	margin-right:0in;
	margin-bottom:0in;
	margin-left:0in;
	text-align:center;
	line-height:115%;
	page-break-before:always;
	page-break-after:avoid;
	font-size:14.0pt;
	font-family:"Ubuntu",sans-serif;
	color:#365F91;}
h2
	{mso-style-link:"Heading 2 Char";
	margin-top:10.0pt;
	margin-right:0in;
	margin-bottom:0in;
	margin-left:0in;
	line-height:115%;
	page-break-after:avoid;
	font-size:13.0pt;
	font-family:"Ubuntu",sans-serif;
	color:#4F81BD;}
p.MsoToc1, li.MsoToc1, div.MsoToc1
	{margin-top:12.0pt;
	margin-right:0in;
	margin-bottom:6.0pt;
	margin-left:0in;
	text-indent:.3in;
	line-height:115%;
	font-size:10.0pt;
	font-family:"Ubuntu",sans-serif;
	font-weight:bold;}
p.MsoToc2, li.MsoToc2, div.MsoToc2
	{margin-top:6.0pt;
	margin-right:0in;
	margin-bottom:0in;
	margin-left:12.0pt;
	text-indent:.3in;
	line-height:115%;
	font-size:10.0pt;
	font-family:"Ubuntu",sans-serif;
	font-style:italic;}
p.MsoFootnoteText, li.MsoFootnoteText, div.MsoFootnoteText
	{mso-style-link:"Footnote Text Char";
	margin:0in;
	font-size:10.0pt;
	font-family:"Ubuntu",sans-serif;}
span.MsoFootnoteReference
	{vertical-align:super;}
span.MsoEndnoteReference
	{vertical-align:super;}
p.MsoEndnoteText, li.MsoEndnoteText, div.MsoEndnoteText
	{mso-style-link:"Endnote Text Char";
	margin:0in;
	text-indent:.3in;
	font-size:10.0pt;
	font-family:"Ubuntu",sans-serif;}
a:link, span.MsoHyperlink
	{color:blue;
	text-decoration:underline;}
p.MsoListParagraph, li.MsoListParagraph, div.MsoListParagraph
	{margin-top:0in;
	margin-right:0in;
	margin-bottom:0in;
	margin-left:.5in;
	text-indent:.3in;
	line-height:115%;
	font-size:12.0pt;
	font-family:"Ubuntu",sans-serif;}
p.MsoListParagraphCxSpFirst, li.MsoListParagraphCxSpFirst, div.MsoListParagraphCxSpFirst
	{margin-top:0in;
	margin-right:0in;
	margin-bottom:0in;
	margin-left:.5in;
	text-indent:.3in;
	line-height:115%;
	font-size:12.0pt;
	font-family:"Ubuntu",sans-serif;}
p.MsoListParagraphCxSpMiddle, li.MsoListParagraphCxSpMiddle, div.MsoListParagraphCxSpMiddle
	{margin-top:0in;
	margin-right:0in;
	margin-bottom:0in;
	margin-left:.5in;
	text-indent:.3in;
	line-height:115%;
	font-size:12.0pt;
	font-family:"Ubuntu",sans-serif;}
p.MsoListParagraphCxSpLast, li.MsoListParagraphCxSpLast, div.MsoListParagraphCxSpLast
	{margin-top:0in;
	margin-right:0in;
	margin-bottom:0in;
	margin-left:.5in;
	text-indent:.3in;
	line-height:115%;
	font-size:12.0pt;
	font-family:"Ubuntu",sans-serif;}
span.MsoSubtleEmphasis
	{color:gray;
	font-style:italic;}
span.MsoIntenseEmphasis
	{color:#4F81BD;
	font-weight:bold;
	font-style:italic;}
span.Heading1Char
	{mso-style-name:"Heading 1 Char";
	mso-style-link:"Heading 1";
	font-family:"Ubuntu",sans-serif;
	color:#365F91;
	font-weight:bold;}
span.Heading2Char
	{mso-style-name:"Heading 2 Char";
	mso-style-link:"Heading 2";
	font-family:"Ubuntu",sans-serif;
	color:#4F81BD;
	font-weight:bold;}
p.DecimalAligned, li.DecimalAligned, div.DecimalAligned
	{mso-style-name:"Decimal Aligned";
	margin-top:0in;
	margin-right:0in;
	margin-bottom:10.0pt;
	margin-left:0in;
	line-height:115%;
	font-size:11.0pt;
	font-family:"Ubuntu",sans-serif;}
span.FootnoteTextChar
	{mso-style-name:"Footnote Text Char";
	mso-style-link:"Footnote Text";
	font-family:"Times New Roman",serif;}
span.EndnoteTextChar
	{mso-style-name:"Endnote Text Char";
	mso-style-link:"Endnote Text";
	font-family:"Ubuntu",sans-serif;}
.MsoChpDefault
	{font-family:"Ubuntu",sans-serif;}
.MsoPapDefault
	{line-height:115%;}
 /* Page Definitions */
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





<div class="WordSection1">

<h1 id="61F4958F-B3BC-45E2-89D1-1596344F5926"><a name="_Toc359077852"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Text Formatting</span></a></h1>

<h2 id="929909A7-1889-4D28-8E83-DDA0FAD84A79"><a name="_Toc359077853"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Inline formatting</span></a></h2>

<p id="B336535F-ACEB-457D-A889-E6781C454011" class="MsoNormal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Here, we demonstrate various types of
inline text formatting and the use of embedded fonts.</span></p>

<p id="0168CCAF-5A28-40A7-9224-330633A7A2E6" class="MsoNormal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>

<p id="213ADEF1-91D8-4B52-842D-4890EDB716DB" class="MsoNormal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Here is some <b>bold, </b><i>italic, <b>bold-italic,
</b></i><u>underlined </u>and <s>struck out </s>&nbsp;text. Then, we have a super<sup>script</sup>
and a sub<sub>script</sub>. Now we see some <span style="color:red">red</span>,
<span style="color:#92D050">green</span> and <span style="color:#0070C0">blue</span>
text. Some text with a <span style="background:yellow">yellow highlight</span>.
Some text in a <span style="border:solid windowtext 1.0pt;padding:0in">box</span>.
Some text in <span style="color:white;background:black">inverse video</span>.</span></p>

<p id="3561EE72-A440-4F36-9489-BD73760833CE" class="MsoNormal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>

<p id="E0B4361B-EF50-4BC6-9EC3-1D6E8747EF23" class="MsoNormal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">A paragraph with styled text: <span class="MsoSubtleEmphasis">subtle emphasis&nbsp; </span>followed by <strong><span style="font-family:&quot;Lao UI&quot;,sans-serif">strong text </span></strong>and <span class="MsoIntenseEmphasis">intense emphasis</span>. This paragraph uses document
wide styles for styling rather than inline text properties as demonstrated in
the previous paragraph — calibre can handle both with equal ease.</span></p>

<h2 id="6085FB06-C20E-4E41-BE20-EF82F9F2118E"><a name="_Toc359077854"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Fun with fonts</span></a></h2>

<p id="C4C2252B-6AFA-4D66-B05D-E59B8232B0CD" class="MsoNormal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>

<p id="CC74018E-4DE8-417D-B8F2-BA99FA7E8C69" class="MsoNormal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">This document has embedded the Ubuntu
font family. The body text is in the Ubuntu typeface, here is some text in the
Ubuntu Mono typeface, notice how every letter has the same width, even i and m.
Every embedded font will automatically be embedded in the output ebook during
conversion. </span></p>

<h2 id="80C7D3F8-0BC7-408D-9C40-3D873A7C9B68"><a name="_Toc359077855"></a><a name="_Paragraph_level_formatting"></a><span class="MsoIntenseEmphasis"><span style="font-family:&quot;Lao UI&quot;,sans-serif;font-style:normal">Paragraph level
formatting</span></span></h2>

<p id="6A8A95E6-89FA-45BE-B97A-CD47ED05F271" class="MsoNormal">&nbsp;</p>

<div style="border:none;border-right:solid windowtext 1.0pt;padding:0in 4.0pt 0in 0in;
background:#DDDDDD">

<p id="B39CC181-5538-4619-91DE-D4C17ECE4BCB" class="MsoNormal" align="right" style="text-align:right;background:#DDDDDD;border:none;padding:0in"><span style="font-family:&quot;Lao UI&quot;,sans-serif;color:black">You can do crazy things
with paragraphs, if the urge strikes you. For instance this paragraph is right
aligned and has a right border. It has also been given a light gray background.</span></p>

</div>

<p id="FF76FBFF-4A88-4250-9FDF-813C01948D89" class="MsoNormal" style="margin-top:
30.0pt;margin-right:0in;margin-bottom:0in;margin-left:.5in;margin-bottom:.0001pt;
text-indent:-.5in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">For the lovers
of poetry amongst you, paragraphs with hanging indents, like this often come in
handy. You can use hanging indents to ensure that a line of poetry retains its
individual identity as a line even when the screen is&nbsp; too narrow to display it
as a single line. Not only does this paragraph have a hanging indent, it is
also has an extra top margin, setting it apart from the preceding paragraph.</span></p>

<h1 id="16003DE9-8276-4DB4-9B90-CF593F79B648" align="left" style="text-align:left"><a name="_Toc359077856"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Tables</span></a></h1>

<p id="9034425D-5CB3-401F-9F07-ADAB2E08CE08" class="MsoNormal">&nbsp;</p>

<table class="MsoTableLightListAccent3" border="1" cellspacing="0" cellpadding="0" align="left" style="border-collapse:collapse;border:none;margin-left:-2.25pt;
 margin-right:7.1pt;margin-bottom:1.35pt">
 <tbody><tr>
  <td width="121" valign="top" style="width:90.9pt;border-top:solid #9BBB59 1.0pt;
  border-left:solid #9BBB59 1.0pt;border-bottom:none;border-right:none;
  background:#9BBB59;padding:0in 5.4pt 0in 5.4pt">
  <p id="A8B77302-C660-4D27-9143-063EBBBF7424" class="MsoNormal" style="line-height:
  normal"><b><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif;
  color:white">ITEM</span></b></p>
  </td>
  <td width="108" valign="top" style="width:81.0pt;border-top:solid #9BBB59 1.0pt;
  border-left:none;border-bottom:none;border-right:solid #9BBB59 1.0pt;
  background:#9BBB59;padding:0in 5.4pt 0in 5.4pt">
  <p id="A91921A5-AAB0-4A7D-9102-AA446606B8BB" class="MsoNormal" style="line-height:
  normal"><b><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif;
  color:white">NEEDED</span></b></p>
  </td>
 </tr>
 <tr>
  <td width="121" valign="top" style="width:90.9pt;border:none;border-left:solid #9BBB59 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="2A908BCB-8695-4525-9291-436CB6E2CDA4" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif">Books</span></p>
  </td>
  <td width="108" valign="top" style="width:81.0pt;border:none;border-right:solid #9BBB59 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="81E72DE7-A1E3-4B31-B094-EEB16E15F161" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif">1</span></p>
  </td>
 </tr>
 <tr>
  <td width="121" valign="top" style="width:90.9pt;border:none;border-left:solid #9BBB59 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="6E6BC14D-31B1-4DA2-8652-BAE67F0D4265" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif">Pens</span></p>
  </td>
  <td width="108" valign="top" style="width:81.0pt;border:none;border-right:solid #9BBB59 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="007141F3-062E-4764-ABE3-594E29AC922F" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif">3</span></p>
  </td>
 </tr>
 <tr>
  <td width="121" valign="top" style="width:90.9pt;border:none;border-left:solid #9BBB59 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="F163B614-342D-4258-BCB8-D6DC118D4740" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif">Pencils</span></p>
  </td>
  <td width="108" valign="top" style="width:81.0pt;border:none;border-right:solid #9BBB59 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="D76EC6FC-CAC7-4792-8195-D6E44514D6B2" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif">2</span></p>
  </td>
 </tr>
 <tr>
  <td width="121" valign="top" style="width:90.9pt;border:none;border-left:solid #9BBB59 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="DB6B7383-184D-44A5-93FF-023ACF3FA8CE" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif">Highlighter</span></p>
  </td>
  <td width="108" valign="top" style="width:81.0pt;border:none;border-right:solid #9BBB59 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="803C289C-042B-4634-A03D-85832A076E41" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif">2
  colors</span></p>
  </td>
 </tr>
 <tr>
  <td width="121" valign="top" style="width:90.9pt;border-top:none;border-left:
  solid #9BBB59 1.0pt;border-bottom:solid #9BBB59 1.0pt;border-right:none;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="3341FDBD-CA40-463E-AC75-62729875F6E2" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif">Scissors</span></p>
  </td>
  <td width="108" valign="top" style="width:81.0pt;border-top:none;border-left:
  none;border-bottom:solid #9BBB59 1.0pt;border-right:solid #9BBB59 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="211C2B9D-0422-44FF-B5A5-6CCA37CE22B0" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif">1 pair</span></p>
  </td>
 </tr>
</tbody></table>

<p id="3FEC18E6-50EA-44EF-980A-851A5E00DE13" class="MsoNormal" style="text-indent:
0in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Tables in Word can vary from
the extremely simple to the extremely complex. calibre tries to do its best
when converting tables. While you may run into trouble with the occasional
table, the vast majority of common cases should be converted very well, as
demonstrated in this section. Note that for optimum results, when creating
tables in Word, you should set their widths using percentages, rather than
absolute units. &nbsp;To the left of this paragraph is a floating two column table
with a nice green border and header row.</span></p>

<p id="2C30A0CA-4C01-4F84-A495-9112C070335F" class="MsoNormal" style="text-indent:
0in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>

<p id="12607223-A005-4253-87E4-D8FDFA3D2F1E" class="MsoNormal" style="text-indent:
0in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Now let’s look at a fancier
table—one with alternating row colors and partial borders. This table is
stretched out to take 100% of the available width.</span></p>

<p id="4BAF7AD6-4710-4C48-B187-F5F60AE93068" class="MsoNormal" style="text-indent:
0in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>

<table class="MsoTableMediumList2Accent1" border="1" cellspacing="0" cellpadding="0" width="100%" style="width:100.0%;border-collapse:collapse;border:none">
 <tbody><tr>
  <td width="20%" nowrap="" valign="top" style="width:20.86%;border:none;border-bottom:
  solid #4F81BD 3.0pt;background:white;padding:0in 5.4pt 0in 5.4pt">
  <p id="A8DB85DA-FE3C-4D66-BCA1-5CB00D45590B" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif;
  color:black">City or Town</span></p>
  </td>
  <td width="12%" valign="top" style="width:12.46%;border:none;border-bottom:
  solid #4F81BD 3.0pt;background:white;padding:0in 5.4pt 0in 5.4pt">
  <p id="EA9B8D60-9AE8-458A-A786-51EBCFFDA9C0" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:black">Point A</span></p>
  </td>
  <td width="16%" valign="top" style="width:16.98%;border:none;border-bottom:
  solid #4F81BD 3.0pt;background:white;padding:0in 5.4pt 0in 5.4pt">
  <p id="4CC7A2DA-473E-44DD-A601-258F094F793C" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:black">Point B</span></p>
  </td>
  <td width="17%" valign="top" style="width:17.0%;border:none;border-bottom:solid #4F81BD 3.0pt;
  background:white;padding:0in 5.4pt 0in 5.4pt">
  <p id="05869A6D-9D3D-4D8C-B859-1DEFB0F6A3DD" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:black">Point C</span></p>
  </td>
  <td width="17%" valign="top" style="width:17.02%;border:none;border-bottom:
  solid #4F81BD 3.0pt;background:white;padding:0in 5.4pt 0in 5.4pt">
  <p id="177A1F85-6E4A-44D1-BCB5-95BB8C8FEA3C" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:black">Point D</span></p>
  </td>
  <td width="15%" valign="top" style="width:15.66%;border:none;border-bottom:
  solid #4F81BD 3.0pt;background:white;padding:0in 5.4pt 0in 5.4pt">
  <p id="20B6FBD3-740E-4009-9A7B-34F3CA6FB144" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:black">Point E</span></p>
  </td>
 </tr>
 <tr>
  <td width="20%" nowrap="" valign="top" style="width:20.86%;border:none;border-right:
  solid #4F81BD 1.0pt;background:white;padding:0in 5.4pt 0in 5.4pt">
  <p id="2697B1EB-2310-4198-9076-8CECF030CBEC" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif;
  color:black">Point A</span></p>
  </td>
  <td width="12%" valign="top" style="width:12.46%;border:none;background:#D3DFEE;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="71632D79-68AE-42A0-828B-0DDF22675631" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:black">—</span></p>
  </td>
  <td width="16%" valign="top" style="width:16.98%;border:none;background:#D3DFEE;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="1C324FA4-CB55-4F25-9D7F-146205ADDDBB" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>
  </td>
  <td width="17%" valign="top" style="width:17.0%;border:none;background:#D3DFEE;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="B6836442-5BAD-4673-A7B5-473245779F1D" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>
  </td>
  <td width="17%" valign="top" style="width:17.02%;border:none;background:#D3DFEE;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="DA9B1CA9-067C-4550-BD8C-4E01F5318D17" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>
  </td>
  <td width="15%" valign="top" style="width:15.66%;border:none;border-right:solid #4F81BD 1.0pt;
  background:#D3DFEE;padding:0in 5.4pt 0in 5.4pt">
  <p id="963B7FAA-4F4E-49D0-ADBA-1BAA220FCC5C" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>
  </td>
 </tr>
 <tr>
  <td width="20%" nowrap="" valign="top" style="width:20.86%;border:none;border-right:
  solid #4F81BD 1.0pt;background:white;padding:0in 5.4pt 0in 5.4pt">
  <p id="035DD4BF-37B1-4A17-A258-5402D0BE4F8E" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif;
  color:black">Point B</span></p>
  </td>
  <td width="12%" valign="top" style="width:12.46%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="07228840-0877-4576-AE4B-E9A401E324B4" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif">87</span></p>
  </td>
  <td width="16%" valign="top" style="width:16.98%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="30FF2CE3-A845-4A06-8DE3-EEE5550DF9A8" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif">—</span></p>
  </td>
  <td width="17%" valign="top" style="width:17.0%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="030847DF-1935-40A2-BA37-C8C5E2B67CB0" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>
  </td>
  <td width="17%" valign="top" style="width:17.02%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="ED66A1F3-BA78-4503-B103-14342474B76D" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>
  </td>
  <td width="15%" valign="top" style="width:15.66%;border:none;border-right:solid #4F81BD 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="2598C274-862B-4362-A9ED-EC29CF77DB3C" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>
  </td>
 </tr>
 <tr>
  <td width="20%" nowrap="" valign="top" style="width:20.86%;border:none;border-right:
  solid #4F81BD 1.0pt;background:white;padding:0in 5.4pt 0in 5.4pt">
  <p id="AE1D315C-43D8-4C64-A25C-AB5EBB494B30" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif;
  color:black">Point C</span></p>
  </td>
  <td width="12%" valign="top" style="width:12.46%;border:none;background:#D3DFEE;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="F894646F-745F-4565-AC49-3595364BAF83" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:black">64</span></p>
  </td>
  <td width="16%" valign="top" style="width:16.98%;border:none;background:#D3DFEE;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="8A19658C-B0B2-44F0-AF23-BCB767218E6B" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:black">56</span></p>
  </td>
  <td width="17%" valign="top" style="width:17.0%;border:none;background:#D3DFEE;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="124BA0AA-3788-48C3-84C6-11E7E148A80C" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:black">—</span></p>
  </td>
  <td width="17%" valign="top" style="width:17.02%;border:none;background:#D3DFEE;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="03FAA796-C072-49AB-863C-040E370E6615" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>
  </td>
  <td width="15%" valign="top" style="width:15.66%;border:none;border-right:solid #4F81BD 1.0pt;
  background:#D3DFEE;padding:0in 5.4pt 0in 5.4pt">
  <p id="C8522F52-307D-4BE8-92B2-33C1D449EB4F" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>
  </td>
 </tr>
 <tr>
  <td width="20%" nowrap="" valign="top" style="width:20.86%;border:none;border-right:
  solid #4F81BD 1.0pt;background:white;padding:0in 5.4pt 0in 5.4pt">
  <p id="44B07CD1-6640-4E6C-B5EC-1E3A2B0839DD" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif;
  color:black">Point D</span></p>
  </td>
  <td width="12%" valign="top" style="width:12.46%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="EF273CB2-B9F7-4FD7-BAD6-12A079A22DE3" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif">37</span></p>
  </td>
  <td width="16%" valign="top" style="width:16.98%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="034E0FDB-A308-4120-97F5-A21E21C85323" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif">32</span></p>
  </td>
  <td width="17%" valign="top" style="width:17.0%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="573E0108-8E10-46CE-AC4F-388C35A28803" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif">91</span></p>
  </td>
  <td width="17%" valign="top" style="width:17.02%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="52ADF126-5CBB-4158-912E-2B17C00C621A" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif">—</span></p>
  </td>
  <td width="15%" valign="top" style="width:15.66%;border:none;border-right:solid #4F81BD 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="4E8ECE77-9470-4245-B6FF-7AA78A891CAA" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>
  </td>
 </tr>
 <tr>
  <td width="20%" nowrap="" valign="top" style="width:20.86%;border:none;border-right:
  solid #4F81BD 1.0pt;background:white;padding:0in 5.4pt 0in 5.4pt">
  <p id="05BF9C75-0F73-49A9-A537-1C53E20D7539" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif;
  color:black">Point E</span></p>
  </td>
  <td width="12%" valign="top" style="width:12.46%;border:none;background:#D3DFEE;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="C25B526E-FC2F-4CB2-9C8E-C79C937BDD67" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:black">93</span></p>
  </td>
  <td width="16%" valign="top" style="width:16.98%;border:none;background:#D3DFEE;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="ED4549A7-1944-42CA-B626-3DFE69AF8A9F" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:black">35</span></p>
  </td>
  <td width="17%" valign="top" style="width:17.0%;border:none;background:#D3DFEE;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="11A95A8B-C7A5-4B88-9DA5-B71312F56CFD" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:black">54</span></p>
  </td>
  <td width="17%" valign="top" style="width:17.02%;border:none;background:#D3DFEE;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="FABDD8B0-86C7-4BC4-9524-C14E07C36FF4" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:black">43</span></p>
  </td>
  <td width="15%" valign="top" style="width:15.66%;border:none;border-right:solid #4F81BD 1.0pt;
  background:#D3DFEE;padding:0in 5.4pt 0in 5.4pt">
  <p id="DDB6948E-290B-4736-91C3-237BE3A35F06" class="MsoNormal" align="center" style="text-align:center;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:black">—</span></p>
  </td>
 </tr>
</tbody></table>

<p id="E2C6B00A-BD08-4294-B10E-CE7AAB066597" class="MsoNormal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>

<p id="05C859D2-30A6-4210-8C3D-E9DF2778AE0B" class="MsoNormal" style="margin-top:
12.0pt;margin-right:0in;margin-bottom:12.0pt;margin-left:0in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Next, we see a table with special
formatting in various locations. Notice how the formatting for the header row
and sub header rows is preserved.</span></p>

<table class="MsoTableMediumShading2Accent5" border="1" cellspacing="0" cellpadding="0" width="100%" style="width:100.0%;border-collapse:collapse;
 border:none">
 <tbody><tr>
  <td width="25%" nowrap="" valign="top" style="width:25.0%;border-top:solid windowtext 2.25pt;
  border-left:none;border-bottom:solid windowtext 2.25pt;border-right:none;
  background:#4BACC6;padding:0in 5.4pt 0in 5.4pt">
  <p id="D69627EF-2823-4AF3-93F6-444F58472550" class="MsoNormal" style="line-height:
  normal"><b><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif;
  color:white">College</span></b></p>
  </td>
  <td width="23%" valign="top" style="width:23.12%;border-top:solid windowtext 2.25pt;
  border-left:none;border-bottom:solid windowtext 2.25pt;border-right:none;
  background:#4BACC6;padding:0in 5.4pt 0in 5.4pt">
  <p id="1A8E8EAF-262B-4EA9-A85A-A878449F5F7A" class="MsoNormal" style="text-indent:
  0in;line-height:normal"><b><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif;
  color:white">New students</span></b></p>
  </td>
  <td width="26%" valign="top" style="width:26.88%;border-top:solid windowtext 2.25pt;
  border-left:none;border-bottom:solid windowtext 2.25pt;border-right:none;
  background:#4BACC6;padding:0in 5.4pt 0in 5.4pt">
  <p id="FFBC2018-AC43-4B69-96C5-1EE5C1ED1269" class="MsoNormal" style="text-indent:
  0in;line-height:normal"><b><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif;
  color:white">Graduating Students</span></b></p>
  </td>
  <td width="25%" valign="top" style="width:25.0%;border-top:solid windowtext 2.25pt;
  border-left:none;border-bottom:solid windowtext 2.25pt;border-right:none;
  background:#4BACC6;padding:0in 5.4pt 0in 5.4pt">
  <p id="2D71B8E4-2479-4F97-8F58-06745C6D06D5" class="MsoNormal" style="line-height:
  normal"><b><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif;
  color:white">Change</span></b></p>
  </td>
 </tr>
 <tr>
  <td width="25%" nowrap="" valign="top" style="width:25.0%;border:none;padding:
  0in 5.4pt 0in 5.4pt">
  <p id="D03EE14C-101D-4B8C-8CD2-15D41D3CD5F7" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>
  </td>
  <td width="23%" valign="top" style="width:23.12%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="C1BA5CC0-9D74-490F-A2D4-EF7E486EA1A7" class="MsoNormal" style="line-height:
  normal"><span class="MsoSubtleEmphasis"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif">Undergraduate</span></span></p>
  </td>
  <td width="26%" valign="top" style="width:26.88%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="127768A2-E421-4CDC-9527-5B3EC102F732" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>
  </td>
  <td width="25%" valign="top" style="width:25.0%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="DF48C229-C1E9-4A49-938B-57E6C41873C8" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>
  </td>
 </tr>
 <tr>
  <td width="25%" nowrap="" valign="top" style="width:25.0%;border:none;padding:
  0in 5.4pt 0in 5.4pt">
  <p id="8258F6E6-5BAE-4763-A27C-DBA8183AEB1C" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif">Cedar
  University</span></p>
  </td>
  <td width="23%" valign="top" style="width:23.12%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="E37C9263-6800-45CC-A62D-F6C9434C4601" class="DecimalAligned" style="line-height:normal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">110</span></p>
  </td>
  <td width="26%" valign="top" style="width:26.88%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="AA64FAB4-0203-4A3C-B0D7-C850355C70EE" class="DecimalAligned" style="line-height:normal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">103</span></p>
  </td>
  <td width="25%" valign="top" style="width:25.0%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="C3FB7849-4855-471F-9D27-782A2DFACAB4" class="DecimalAligned" style="line-height:normal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">+7</span></p>
  </td>
 </tr>
 <tr>
  <td width="25%" nowrap="" valign="top" style="width:25.0%;border:none;padding:
  0in 5.4pt 0in 5.4pt">
  <p id="3FB36F62-3168-4959-902C-4B29052BAC4D" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif">Oak
  Institute</span></p>
  </td>
  <td width="23%" valign="top" style="width:23.12%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="42ABD6B3-8E2F-402D-BEBC-9ACFA59C769D" class="DecimalAligned" style="line-height:normal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">202</span></p>
  </td>
  <td width="26%" valign="top" style="width:26.88%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="D02EE1BB-6BC5-497C-B458-E35925530D18" class="DecimalAligned" style="line-height:normal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">210</span></p>
  </td>
  <td width="25%" valign="top" style="width:25.0%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="AE8038F9-570B-4F56-8B77-F316F0409129" class="DecimalAligned" style="line-height:normal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">-8</span></p>
  </td>
 </tr>
 <tr>
  <td width="25%" nowrap="" valign="top" style="width:25.0%;border:none;padding:
  0in 5.4pt 0in 5.4pt">
  <p id="17864D22-0824-4678-A7C1-A5955CA01B00" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>
  </td>
  <td width="23%" valign="top" style="width:23.12%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="81C14F5C-B757-4639-88BA-9F1581262002" class="MsoNormal" style="line-height:
  normal"><span class="MsoSubtleEmphasis"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif">Graduate</span></span></p>
  </td>
  <td width="26%" valign="top" style="width:26.88%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="2D75F36D-8D9D-4B47-A099-66B40CEDD87C" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>
  </td>
  <td width="25%" valign="top" style="width:25.0%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="E07A4786-980E-4681-B0D4-1EF6FBC6BE5C" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>
  </td>
 </tr>
 <tr>
  <td width="25%" nowrap="" valign="top" style="width:25.0%;border:none;padding:
  0in 5.4pt 0in 5.4pt">
  <p id="6F13A384-4541-46AC-AC0A-406C7E117739" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif">Cedar
  University</span></p>
  </td>
  <td width="23%" valign="top" style="width:23.12%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="F139BA23-2EA7-4AE9-974C-574919E40A9F" class="DecimalAligned" style="line-height:normal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">24</span></p>
  </td>
  <td width="26%" valign="top" style="width:26.88%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="99E9F5DF-8ADA-46D3-8E82-AD6B5C79D6CE" class="DecimalAligned" style="line-height:normal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">20</span></p>
  </td>
  <td width="25%" valign="top" style="width:25.0%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="F1C453A9-2E44-4410-BA7F-A06B6C63475A" class="DecimalAligned" style="line-height:normal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">+4</span></p>
  </td>
 </tr>
 <tr>
  <td width="25%" nowrap="" valign="top" style="width:25.0%;border:none;padding:
  0in 5.4pt 0in 5.4pt">
  <p id="79F1024B-D50F-4D49-842E-1F53AFEA6F25" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif">Elm
  College</span></p>
  </td>
  <td width="23%" valign="top" style="width:23.12%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="9DB957BF-0419-4F36-A1E9-D409BCAB6E75" class="DecimalAligned" style="line-height:normal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">43</span></p>
  </td>
  <td width="26%" valign="top" style="width:26.88%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="95BF70FA-3404-4143-B219-4AC512C5D452" class="DecimalAligned" style="line-height:normal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">53</span></p>
  </td>
  <td width="25%" valign="top" style="width:25.0%;border:none;padding:0in 5.4pt 0in 5.4pt">
  <p id="58C47458-FDF8-49AD-8D59-9D3E199DD20E" class="DecimalAligned" style="line-height:normal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">-10</span></p>
  </td>
 </tr>
 <tr>
  <td width="25%" nowrap="" valign="top" style="width:25.0%;border-top:double windowtext 2.25pt;
  border-left:none;border-bottom:solid windowtext 2.25pt;border-right:none;
  background:white;padding:0in 5.4pt 0in 5.4pt">
  <p id="3596EC3A-6BBF-4136-A711-157AA70B3913" class="MsoNormal" style="line-height:
  normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif;
  color:black">Total</span></p>
  </td>
  <td width="23%" valign="top" style="width:23.12%;border-top:double windowtext 2.25pt;
  border-left:none;border-bottom:solid windowtext 2.25pt;border-right:none;
  background:white;padding:0in 5.4pt 0in 5.4pt">
  <p id="D24F801E-BBD0-4686-B9A8-4B51B4072C26" class="DecimalAligned" style="line-height:normal"><span style="font-family:&quot;Lao UI&quot;,sans-serif;
  color:black">998</span></p>
  </td>
  <td width="26%" valign="top" style="width:26.88%;border-top:double windowtext 2.25pt;
  border-left:none;border-bottom:solid windowtext 2.25pt;border-right:none;
  background:white;padding:0in 5.4pt 0in 5.4pt">
  <p id="5DF339DC-23EE-4EC1-8FF5-02A4D2F6C3E8" class="DecimalAligned" style="line-height:normal"><span style="font-family:&quot;Lao UI&quot;,sans-serif;
  color:black">908</span></p>
  </td>
  <td width="25%" valign="top" style="width:25.0%;border-top:double windowtext 2.25pt;
  border-left:none;border-bottom:solid windowtext 2.25pt;border-right:none;
  background:white;padding:0in 5.4pt 0in 5.4pt">
  <p id="CD5DAFF6-E6BA-4288-B1F5-5D015891B060" class="DecimalAligned" style="line-height:normal"><span style="font-family:&quot;Lao UI&quot;,sans-serif;
  color:black">90</span></p>
  </td>
 </tr>
</tbody></table>

<p id="99B19E0F-B60A-4A88-8FA5-277ACAD67F3E" class="MsoFootnoteText"><span class="MsoSubtleEmphasis"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Source:</span></span><span style="font-family:&quot;Lao UI&quot;,sans-serif"> Fictitious data, for illustration
purposes only</span></p>

<p id="5D0F9038-84A9-4C45-945A-B1C7C24A09B9" class="MsoNormal" style="margin-top:
12.0pt"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Next, we have something a
little more complex, a nested table, i.e. a table inside another table.
Additionally, the inner table has some of its cells merged. The table is
displayed horizontally centered.</span></p>

<p id="EF193876-FFF3-4DB8-9E55-88D4FC6A8C0A" class="MsoNormal" style="margin-top:
12.0pt"><span style="font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>

<div align="center">

<table class="MsoTableGrid" border="1" cellspacing="0" cellpadding="0" width="70%" style="width:70.0%;border-collapse:collapse;border:none">
 <tbody><tr style="height:53.95pt">
  <td width="319" valign="top" style="width:239.4pt;border:solid black 1.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:53.95pt">
  <div align="center">
  <table class="MsoTableGrid" border="1" cellspacing="0" cellpadding="0" width="80%" style="width:80.0%;border-collapse:collapse;border:none">
   <tbody><tr>
    <td width="112" rowspan="2" valign="top" style="width:83.9pt;border:solid black 1.0pt;
    padding:0in 5.4pt 0in 5.4pt">
    <p id="F138B705-92B3-472B-8BBF-DE5115B0107B" class="MsoNormal" style="margin-top:12.0pt;text-indent:0in;line-height:normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif">One</span></p>
    <p id="58A39288-50A8-4B12-BCFF-87A379540F99" class="MsoNormal" style="margin-top:12.0pt;text-indent:0in;line-height:normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif">Three</span></p>
    </td>
    <td width="109" valign="top" style="width:81.45pt;border:solid black 1.0pt;
    border-left:none;padding:0in 5.4pt 0in 5.4pt">
    <p id="470B9D22-A400-48C8-9F97-1461FDA0A598" class="MsoNormal" style="margin-top:12.0pt;text-indent:0in;line-height:normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif">Two</span></p>
    </td>
   </tr>
   <tr>
    <td width="109" valign="top" style="width:81.45pt;border-top:none;border-left:
    none;border-bottom:solid black 1.0pt;border-right:solid black 1.0pt;
    padding:0in 5.4pt 0in 5.4pt">
    <p id="3FE22FBD-4C8F-44AB-AA6D-2837319BCA9E" class="MsoNormal" style="margin-top:12.0pt;text-indent:0in;line-height:normal"><span style="font-size:11.0pt;font-family:&quot;Lao UI&quot;,sans-serif">Four</span></p>
    </td>
   </tr>
  </tbody></table>
  </div>
  <p id="CB4437E9-0117-4D6C-849C-2387A4306AE5" class="MsoNormal" style="margin-top:
  12.0pt;text-indent:0in;line-height:normal"></p>
  </td>
  <td width="319" valign="top" style="width:239.4pt;border:solid black 1.0pt;
  border-left:none;padding:0in 5.4pt 0in 5.4pt;height:53.95pt">
  <p id="91C6922C-DB89-44E5-873A-2B337D1DC8B2" class="MsoNormal" style="margin-top:
  12.0pt;text-indent:0in;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif">To the left is a table inside a table, with
  some cells merged.</span></p>
  </td>
 </tr>
</tbody></table>

</div>

<p id="3C392836-8334-40C2-8FB4-92EFAA1A96B7" class="MsoNormal" style="margin-top:
12.0pt;text-indent:.5in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>

<p id="964DB43E-13BF-4BBF-BD8C-68EA649846BF" class="MsoNormal" style="margin-top:
12.0pt;text-indent:.5in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">We end
with a fancy calendar, note how much of the original formatting is preserved. Note
that this table will only display correctly on relatively wide screens. In
general, very wide tables or tables whose cells have fixed width requirements
don’t fare well in ebooks.</span></p>

<div align="center">

<table class="Calendar3" border="0" cellspacing="0" cellpadding="0" style="border-collapse:
 collapse">
 <tbody><tr style="height:40.5pt">
  <td width="500" colspan="13" valign="top" style="width:375.0pt;padding:0in 5.4pt 0in 5.4pt;
  height:40.5pt">
  <p id="BA084723-6139-4F11-98C3-F718E5B98E0A" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:22.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">December 2007</span></p>
  </td>
  <td style="border:none;padding:0in 0in 0in 0in" width="1"><p class="MsoNormal">&nbsp;</p></td>
 </tr>
 <tr style="height:.4in">
  <td width="58" nowrap="" valign="bottom" style="width:.6in;border:none;border-bottom:
  solid #365F91 1.0pt;padding:0in 2.15pt 0in 2.15pt;height:.4in">
  <p id="F1B3796B-EFBD-48A7-98B3-FAC3A70CEC65" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">Sun</span></p>
  </td>
  <td width="15" nowrap="" valign="bottom" style="width:11.1pt;padding:0in 2.15pt 0in 2.15pt;
  height:.4in">
  <p id="15F0D6D0-A4E1-441B-ABDD-D788EE43CCBE" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="66" nowrap="" valign="bottom" style="width:49.2pt;border:none;border-bottom:
  solid #7F7F7F 1.0pt;padding:0in 2.15pt 0in 2.15pt;height:.4in">
  <p id="313E252A-6D54-4599-AE55-B6865931535D" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">Mon</span></p>
  </td>
  <td width="15" nowrap="" valign="bottom" style="width:11.1pt;padding:0in 2.15pt 0in 2.15pt;
  height:.4in">
  <p id="3045616E-71B1-45BA-919E-75E351497B46" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" nowrap="" valign="bottom" style="width:.6in;border:none;border-bottom:
  solid #7F7F7F 1.0pt;padding:0in 2.15pt 0in 2.15pt;height:.4in">
  <p id="6FBC5496-F635-4DCB-A3D7-03ED0FFAE72A" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">Tue</span></p>
  </td>
  <td width="15" nowrap="" valign="bottom" style="width:11.1pt;padding:0in 2.15pt 0in 2.15pt;
  height:.4in">
  <p id="03143E94-5536-472B-B00F-769D9634C779" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" nowrap="" valign="bottom" style="width:.6in;border:none;border-bottom:
  solid #7F7F7F 1.0pt;padding:0in 2.15pt 0in 2.15pt;height:.4in">
  <p id="12DF4981-8135-4696-8645-56B63CCF0924" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">Wed</span></p>
  </td>
  <td width="15" nowrap="" valign="bottom" style="width:11.1pt;padding:0in 2.15pt 0in 2.15pt;
  height:.4in">
  <p id="8F6E0EAD-78FC-4AA6-BBF3-3F5AC8DFB61A" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" nowrap="" valign="bottom" style="width:.6in;border:none;border-bottom:
  solid #7F7F7F 1.0pt;padding:0in 2.15pt 0in 2.15pt;height:.4in">
  <p id="81E04565-5FCF-4417-83C2-46AEC3C7BE23" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">Thu</span></p>
  </td>
  <td width="15" nowrap="" valign="bottom" style="width:11.1pt;padding:0in 2.15pt 0in 2.15pt;
  height:.4in">
  <p id="583AECE2-9DA6-49BB-817E-E1ACEEE653C5" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" nowrap="" valign="bottom" style="width:.6in;border:none;border-bottom:
  solid #7F7F7F 1.0pt;padding:0in 2.15pt 0in 2.15pt;height:.4in">
  <p id="A2FA48B0-2098-407E-975F-DA948A5F3A21" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">Fri</span></p>
  </td>
  <td width="16" nowrap="" valign="bottom" style="width:11.8pt;padding:0in 2.15pt 0in 2.15pt;
  height:.4in">
  <p id="297FD42D-BAF2-43F1-97D9-66A7320A4F5A" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" nowrap="" colspan="2" valign="bottom" style="width:.6in;border:none;
  border-bottom:solid #365F91 1.0pt;padding:0in 2.15pt 0in 2.15pt;height:.4in">
  <p id="9F0AA8EE-160C-4646-BA58-0072D34079F3" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">Sat</span></p>
  </td>
 </tr>
 <tr style="height:.4in">
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #365F91 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="7CF74925-09DC-4CF7-B65E-AB13D63177D4" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="F358EEF7-F791-4EC5-A562-5D52AE5AB27C" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="66" valign="top" style="width:49.2pt;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="8D922458-30EB-47D8-85BA-CB809BB20DB5" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="BCA822C9-1AD4-439D-86C3-23859799903C" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="1F42DE94-2E06-42AF-A802-B306E6E153DE" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="1D9953FF-85E9-4014-ABD9-2E28ED53B553" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="1AEC2366-90D9-4E53-BB29-899BA44BBE4C" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="23D2E514-BA5F-4C88-8F41-123115211DCB" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="C76DC7D4-B152-43C5-827A-08147CAF795E" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="56160CA0-EC5A-4886-AB1D-F6416F08862D" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="83708D79-3EBA-4717-8C5B-5182476C090D" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="16" valign="top" style="width:11.8pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="FF3C22B3-78D9-4195-8741-91F0B7E2F598" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" colspan="2" valign="top" style="width:.6in;border:none;border-right:
  solid #365F91 3.0pt;padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="E7E50382-E91D-48E5-9900-D9D15973A065" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">1</span></p>
  </td>
 </tr>
 <tr>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #365F91 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="E4063E7C-DA8A-4012-83E9-1F1A0ABD7F77" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="6D548438-DD7C-4BC6-ADEB-67EC59D959FD" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="66" valign="top" style="width:49.2pt;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="8C5C76D2-9ED5-4924-A1EF-3E9C5691E2A3" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="74D022A7-C7DB-45F5-9A45-2B85C4B2BF89" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="E5772049-92DB-4601-AC99-0F51C5002AC9" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="0C9BD033-2EFE-4A79-8F38-3FFF70611842" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="89793CA9-0156-4188-A55A-ACCF2CD5B77A" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="99C72FFD-9F5D-43E9-B55C-B8408AA2156E" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="21AE4359-170E-49DA-8DC7-9A966C01DB4D" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="7A50F973-18BC-4D6B-836E-01EC3FC02134" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="ABCD284B-A43D-4770-8718-1E2722AC58B3" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="16" valign="top" style="width:11.8pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="9DBC6A22-7516-44A1-95F2-59EE01CF62FF" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" colspan="2" valign="top" style="width:.6in;border:none;border-bottom:
  solid #365F91 1.0pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="90CFF9AD-0646-43B6-AFA2-DBC93FF1FA73" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">&nbsp;</span></p>
  </td>
 </tr>
 <tr style="height:.4in">
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #365F91 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="A72E4E8C-199B-4764-BD1B-D2C1DAC9A8EE" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">2</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="B738F24D-6ABC-4235-A848-C134BDD2F971" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="66" valign="top" style="width:49.2pt;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="2C8CD59C-0C51-406C-8CDB-7E285ADF2164" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">3</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="750EC43E-3D1F-481E-B3C8-7C85D050F50B" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="B34665A3-DF51-42E4-B182-F06F6303225A" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">4</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="ACB81A7B-8FE5-4B19-88C4-834062182384" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="7EB6C004-D0C5-4F96-B76A-D62D528A971A" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">5</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="25F1242E-A5C7-4734-A042-4DFE08CA6F6E" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="DF366CD5-D90A-4BD9-8405-83FF20EA5C31" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">6</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="6B10CF59-97BF-45A1-957E-0AAB4C8ED3DC" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="7CBF66DC-AA92-4782-BC1B-861C1BAEDFBA" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">7</span></p>
  </td>
  <td width="16" valign="top" style="width:11.8pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="3D62F075-2F78-4A46-A476-2E5BB88E2BBA" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" colspan="2" valign="top" style="width:.6in;border:none;border-right:
  solid #365F91 3.0pt;padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="ECBB2523-AFF4-4D19-A746-CA23094000E9" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">8</span></p>
  </td>
 </tr>
 <tr>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #365F91 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="2B7FC3ED-0511-4EE8-AC33-A79E2E2FA3CB" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="016BB693-560B-478A-8CB2-FAC0F08A0A57" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="66" valign="top" style="width:49.2pt;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="968314D7-D386-47BA-A4FD-7FFF9D8EA98E" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="97B124FC-5EFF-4075-85AF-0D2BF93478F7" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="20872836-5D47-4187-BD66-140DBB40D0B2" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="6C496153-29D0-4C8A-989F-71E2880AA177" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="E08F82C5-C76E-4217-A6EA-8552C23AD98F" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="8DFCD831-EB89-4E1F-A7C5-F85449238CE4" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="01802387-404A-4333-A398-9FAAD0777F17" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="2FE34DF8-035A-41A9-B7FB-71FE1C58A869" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="3E9B919E-922F-4196-B418-70B5FE2CAA8B" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="16" valign="top" style="width:11.8pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="EED2EB3E-E321-41CE-9524-8073932DA9D2" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" colspan="2" valign="top" style="width:.6in;border:none;border-bottom:
  solid #365F91 1.0pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="119CB2FD-E7C8-4B03-A49A-1095F3B3B15D" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">&nbsp;</span></p>
  </td>
 </tr>
 <tr style="height:.4in">
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #365F91 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="B327142C-2F90-4CC7-8EC2-4CBDB3149A0A" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">9</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="6D990D72-88D5-4DB0-A333-BD54F379C3EA" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="66" valign="top" style="width:49.2pt;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="839BD779-1B1C-48DD-BEE3-51BF41DFED83" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">10</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="F332BD7E-3CE6-4B1D-91BA-A3F073A5434D" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="6DB8814B-76AE-4E00-946A-99C381D57AE4" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">11</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="4101B5EE-CEAE-4301-9472-C0B03E67D0AE" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="22776CB6-7E84-44B8-8A45-7C286BFB9E88" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">12</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="FF9D66B5-1EAE-4A66-A5FF-15C84D12FA81" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="24EC5BD2-44D5-4A78-B05E-66B10F3687F5" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">13</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="D5B84633-F423-4CC4-8816-E033348E82BB" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="8E1353FD-FE4A-4339-B68E-1442A990AFDA" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">14</span></p>
  </td>
  <td width="16" valign="top" style="width:11.8pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="F63764E9-BEE5-49AD-B2D0-C84FE8F06834" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" colspan="2" valign="top" style="width:.6in;border:none;border-right:
  solid #365F91 3.0pt;padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="BAFBB143-5DE5-4929-9FB6-F24A578ADB4A" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">15</span></p>
  </td>
 </tr>
 <tr>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #365F91 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="8EE445EB-8CA7-4588-AF2F-A15F6DD26BE8" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="D4C771B6-A4BA-4765-AD33-51C3C61CCF72" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="66" valign="top" style="width:49.2pt;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="FC9A7311-5AD0-41F3-AD38-471B9123C3EB" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="41E5C753-2891-4C93-8177-55CE929E10C6" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="3B91ABF5-810C-4957-BE01-6616221D41C8" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="BB399730-62D8-4FBD-9B55-355B36E855DC" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="5FC2525F-D506-4624-88AC-20D3A1177458" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="777079C1-9D8F-4DBE-AEBA-121A5CACC26E" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="1AB86736-7861-41C3-86F3-C279868758C8" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="CA00F224-A196-44C3-94DD-FB6D65A5AD3A" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="AC0E950A-A218-4BAC-BDFE-2A6B4E9F38CC" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="16" valign="top" style="width:11.8pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="C7B74F93-A15C-4712-A95F-2CEFED3B5554" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" colspan="2" valign="top" style="width:.6in;border:none;border-bottom:
  solid #365F91 1.0pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="E53A96CB-A188-4419-9336-C7CA0CF201E6" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">&nbsp;</span></p>
  </td>
 </tr>
 <tr style="height:.4in">
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #365F91 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="0CBA6BB8-9DC3-4AF5-9283-7E53601460F5" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">16</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="7455DB44-2BBB-47F1-80C7-0274D5C16885" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="66" valign="top" style="width:49.2pt;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="62EC85C1-D55A-4996-8775-49A205F4ED9E" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">17</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="E7707972-A56E-470E-BEDB-8DA6B43492F9" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="D95EF1F7-52AF-454B-AB48-4745E1BAFE1C" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">18</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="60969112-1780-4EA1-AC08-7A4ABDDA7E07" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="DCCA0535-19D2-4B07-BD4E-89D83FA8E56C" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">19</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="08691729-2078-41CE-A379-0C1F0764CA2D" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="7F9B71C4-E12A-41E9-88C8-606A4E3308F6" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">20</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="B0C54154-C84E-4C53-8D90-92FDE0B35DBC" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="973D1B44-D352-4BE0-84FC-DACAD261EC8B" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">21</span></p>
  </td>
  <td width="16" valign="top" style="width:11.8pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="8856D3B5-627F-4B5D-94D6-B6EDF860BD1F" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" colspan="2" valign="top" style="width:.6in;border:none;border-right:
  solid #365F91 3.0pt;padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="9896252C-EEF4-4A3E-A129-973945C8AA3E" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">22</span></p>
  </td>
 </tr>
 <tr>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #365F91 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="6652EA2D-0E07-4F8C-842C-4239BC71B73B" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="F938FDCC-A224-4473-B70A-8A613149FC36" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="66" valign="top" style="width:49.2pt;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="DF359587-BAE7-4C29-93D3-E63AC0D4A24D" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="21FB8574-A593-4B0C-B2EE-70FF6044F012" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="D6D27438-8957-4AF6-A4EA-ED7B2D6BA124" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="2D2CC779-9909-4C6F-BDF0-898926C9B0B6" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="1677BE4B-CDA6-4962-AC16-3C92DA5D686F" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="D24F4352-227B-47FF-AD13-9BC9D445DA34" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="5F852569-896B-4E1E-ABF7-2BEE7BADC601" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="8BF0BD2D-2D81-4103-B5F6-D079E2EA6D9C" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="A27D94B4-0BE3-429E-9EA6-03B224B45E70" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="16" valign="top" style="width:11.8pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="F31B5A11-49CD-4FA6-8D57-ADCE7EE51A42" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" colspan="2" valign="top" style="width:.6in;border:none;border-bottom:
  solid #365F91 1.0pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="9EE0AAF5-FBF0-450E-A457-25B4D606908F" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">&nbsp;</span></p>
  </td>
 </tr>
 <tr style="height:.4in">
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #365F91 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="A6A439D7-B29A-45D9-8137-D890B3A3DBD2" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">23</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="6DD0CFB8-FE55-44EF-B9DE-3AB1F3D8C531" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="66" valign="top" style="width:49.2pt;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="717F6B69-277D-4F94-953B-17DFD46F796F" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">24</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="51E39A7A-DC00-49E7-9FDB-CAB48DF9EF2A" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="C7BF2DAA-9F54-4D4A-9E37-623AE17C2B91" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">25</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="A3208505-FB6E-482A-A82F-E8ED5AD40760" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="A7E19C42-CEAA-4AE0-8235-1A6BEC950C67" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">26</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="AF2930D2-C3BE-4F6B-9243-1F05FF4052B6" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="FFDACBD0-39A6-41D1-AD31-2127B8A584CE" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">27</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="EF96D74B-428A-4E5C-A658-39279211E7A8" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="CAAB907C-0C68-4181-82E3-96F14769825B" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">28</span></p>
  </td>
  <td width="16" valign="top" style="width:11.8pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="153788BA-B94A-46B1-B5D4-8513B29B1F23" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" colspan="2" valign="top" style="width:.6in;border:none;border-right:
  solid #365F91 3.0pt;padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="F209B7DC-50B4-4BE0-9B06-054F6EED25B7" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">29</span></p>
  </td>
 </tr>
 <tr>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #365F91 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="5432247A-5BA9-4819-9BCB-BCBA3572AA1B" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="ED922339-1356-4B22-8E68-D41910242C4D" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="66" valign="top" style="width:49.2pt;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="C3A9C380-41F5-4C82-BA5D-64CCE1AB6FCB" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="3150D54A-A247-4B44-BAC4-A12CFDDCA751" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="558F59C5-5007-4576-ADBE-0AA81CFAE1A8" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="A73BF2FD-68E6-434E-8B16-7B84F9AFE08B" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="8A22E9DA-6C6D-4491-9D88-1E6080087BF3" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="15B2C3A2-7A42-4817-ADCD-42399A70E74D" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="C597014E-317F-499C-ADDE-BA6C15F7FD68" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="C2E95BA2-8E73-401E-96C7-527FA20597AF" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-bottom:solid #7F7F7F 1.0pt;
  padding:0in 5.4pt 0in 5.4pt">
  <p id="1E56FBD0-63B1-41D4-AF78-E864F04BBBDF" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="16" valign="top" style="width:11.8pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="6DD9AA6A-B5A4-4D79-90C8-99FD80288035" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" colspan="2" valign="top" style="width:.6in;border:none;border-bottom:
  solid #365F91 1.0pt;padding:0in 5.4pt 0in 5.4pt">
  <p id="A180DF86-79D8-4C35-A20C-36D8399267A0" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">&nbsp;</span></p>
  </td>
 </tr>
 <tr style="height:.4in">
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #365F91 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="EFD7FB64-67CB-484C-BBC2-EC9FEDDE9E79" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">30</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="153CAAE4-4A2F-4249-906F-06D8EB4D32E2" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="66" valign="top" style="width:49.2pt;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="B7759810-50E5-4193-B2BB-C41C40A64664" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">31</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="289691D0-E67E-4CAE-9957-4CCC21F1CBEA" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="D5CDB2B8-C81A-44B0-9CED-2A69757E0636" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="086C4AC5-57C3-4347-B1AB-29BFC82614AF" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="219B51A4-069C-45A7-A333-E3974EEF304C" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="3460B9F5-C71F-414F-B020-B540C3918670" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="7B81272D-FFF5-4570-A297-7DF1748875DE" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="15" valign="top" style="width:11.1pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="32D98226-E149-4651-8B6A-31F274B6F7D2" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" valign="top" style="width:.6in;border:none;border-right:solid #7F7F7F 3.0pt;
  padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="7FFD99D1-5187-4197-9BF8-FB44947C6083" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="16" valign="top" style="width:11.8pt;border:none;padding:0in 5.4pt 0in 5.4pt;
  height:.4in">
  <p id="15EE4AED-D969-4B73-AC11-B791CCBE74D5" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#7F7F7F">&nbsp;</span></p>
  </td>
  <td width="58" colspan="2" valign="top" style="width:.6in;border:none;border-right:
  solid #365F91 3.0pt;padding:0in 5.4pt 0in 5.4pt;height:.4in">
  <p id="6D9F2D53-FD85-49CD-8EFB-45A894E52DB4" class="MsoNormal" align="right" style="text-align:right;line-height:normal"><span style="font-size:11.0pt;
  font-family:&quot;Lao UI&quot;,sans-serif;color:#365F91">&nbsp;</span></p>
  </td>
 </tr>
 <tr height="0">
  <td width="59" style="border:none"></td>
  <td width="15" style="border:none"></td>
  <td width="66" style="border:none"></td>
  <td width="15" style="border:none"></td>
  <td width="59" style="border:none"></td>
  <td width="15" style="border:none"></td>
  <td width="65" style="border:none"></td>
  <td width="15" style="border:none"></td>
  <td width="59" style="border:none"></td>
  <td width="15" style="border:none"></td>
  <td width="59" style="border:none"></td>
  <td width="16" style="border:none"></td>
  <td width="58" style="border:none"></td>
  <td width="1" style="border:none"></td>
 </tr>
</tbody></table>

</div>

<h1 id="379020F9-B5BD-40B3-86A6-39EE44955D5D"><a name="_Toc359077857"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Structural Elements</span></a></h1>

<p id="5C118C4F-AF1E-468F-B622-4299AE015AA0" class="MsoNormal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>

<p id="AE962FFD-12D4-4CC4-A094-510348AE611E" class="MsoNormal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Miscellaneous structural elements you
can add to your document, like footnotes, endnotes, dropcaps and the like. </span></p>

<h2 id="DF5A7B6F-856B-4BA9-8791-4A702CE4AB06"><a name="_Toc359077858"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Footnotes &amp; Endnotes</span></a></h2>

<p id="2EFEABE7-5DF7-4E3D-9B86-A1402401967B" class="MsoNormal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>

<p id="37235022-3228-495E-81FB-4196FFA81F36" class="MsoNormal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Footnotes<a href="#_ftn1" name="_ftnref1" title=""><span class="MsoFootnoteReference"><span class="MsoFootnoteReference"><span style="font-size:12.0pt;line-height:115%;
font-family:&quot;Lao UI&quot;,sans-serif">[1]</span></span></span></a> and endnotes<a href="#_edn1" name="_ednref1" title=""><span class="MsoEndnoteReference"><span class="MsoEndnoteReference"><span style="font-size:12.0pt;line-height:115%;
font-family:&quot;Lao UI&quot;,sans-serif">[i]</span></span></span></a> are automatically
recognized and both are converted to endnotes, with backlinks for maximum ease
of use in ebook devices.</span></p>

<h2 id="F41B701B-664A-4296-BB5C-FAD2C5A6E993"><a name="_Toc359077859"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Dropcaps</span></a></h2>

<p id="3B647D1E-0436-4E7A-8594-0DC4F97D4F38" class="MsoNormal">&nbsp;</p>

<div>

<table cellspacing="0" cellpadding="0" hspace="0" vspace="0" align="left">
 <tbody><tr>
  <td valign="top" align="left" style="padding-top:0in;padding-right:0in;
  padding-bottom:0in;padding-left:0in">
  <p id="79946E3C-75D6-47D6-9941-BC45C1732170" class="MsoNormal" style="text-indent:
  0in;line-height:47.55pt;page-break-after:avoid;vertical-align:baseline"><span style="font-size:58.5pt;font-family:&quot;Lao UI&quot;,sans-serif">D</span></p>
  </td>
 </tr>
</tbody></table>

</div>

<p id="A36F37F4-F471-44D1-A0C4-DE4058EA32B9" class="MsoNormal" style="text-indent:
0in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">rop caps are used to
emphasize the leading paragraph at the start of a section. In Word it is
possible to specify how many lines of text a drop-cap should use. Because of
limitations in ebook technology, this is not possible when converting.&nbsp;
Instead, the converted drop cap will use font size and line height to simulate
the effect as well as possible. While not as good as the original, the result
is usually tolerable. This paragraph has a “D” dropcap set to occupy three
lines of text with a font size of 58.5 pts. Depending on the screen width and
capabilities of the device you view the book on, this dropcap can look anything
from perfect to ugly.</span></p>

<h2 id="1E5690FE-9423-4948-BB33-E9A94B7DAF0A"><a name="_Toc359077860"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Links</span></a></h2>

<p id="8684A6B9-6F21-4026-A797-60BB28ED5025" class="MsoNormal" style="text-indent:
0in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>

<p id="975F2584-D2D5-4647-A9CE-B0ECDBF4C5B9" class="MsoNormal" style="text-indent:
.5in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Two kinds of links are
possible, those that refer to an external website and those that refer to locations
inside the document itself. Both are supported by calibre. For example, here is
a link pointing to the <a href="http://calibre-ebook.com/download">calibre
download page</a>. Then we have a link that points back to the section on <a href="#_Paragraph_level_formatting">paragraph level formatting</a> in this
document.</span></p>

<h2 id="09B0EEDF-79C0-4405-BAF7-094F051BE3EA"><a name="_Toc359077861"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Table of Contents</span></a></h2>

<p id="D671D8AA-51A3-40D3-A648-617CC374691A" class="MsoNormal">&nbsp;</p>

<p id="ACFCC1E1-F47F-46CF-99E8-5CA523BB5377" class="MsoNormal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">There are two approaches that calibre
takes when generating a Table of Contents. The first is if the Word document
has a Table of Contents itself. Provided that the Table of Contents uses
hyperlinks, calibre will automatically use it. The levels of the Table of
Contents are identified by their left indent, so if you want the ebook to have
a multi-level Table of Contents, make sure you create a properly indented Table
of Contents in Word.</span></p>

<p id="70EEDDCF-C8F6-44E6-B211-3D5C9FAAD19B" class="MsoNormal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">If no Table of Contents is found in the
document, then a table of contents is automatically generated from the headings
in the document. A heading is identified as something that has the Heading 1 or
Heading 2, etc. style applied to it. These headings are turned into a Table of
Contents with Heading 1 being the topmost level, Heading 2 the second level and
so on.</span></p>

<p id="78B99C20-FB61-4602-B85C-48850B5D9E8A" class="MsoNormal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>

<p id="3ADAE6D3-EAB3-4C47-84CB-32F44F9E44E8" class="MsoNormal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;You can see the Table of Contents
created by calibre by clicking the Table of Contents button in whatever viewer
you are using to view the converted ebook. </span></p>

<p id="36E2A6FB-979B-4E92-AFB4-400580F95330" class="MsoToc1"><span style="font-family:&quot;Lao UI&quot;,sans-serif"><a href="#_Toc359077851"><span style="font-size:12.0pt;line-height:115%">Demonstration of DOCX support in
calibre</span><span style="font-size:12.0pt;line-height:115%;color:windowtext;
display:none;text-decoration:none">. </span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">1</span></a></span></p>

<p id="50CE4DB9-321C-49E9-8473-E7D13B149520" class="MsoToc1"><span style="font-family:&quot;Lao UI&quot;,sans-serif"><a href="#_Toc359077852"><span style="font-size:12.0pt;line-height:115%">Text Formatting</span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">.. </span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">2</span></a></span></p>

<p id="29DE676C-9981-4A2E-BACF-CA0B5A23A07C" class="MsoToc2"><span style="font-family:&quot;Lao UI&quot;,sans-serif"><a href="#_Toc359077853"><span style="font-size:12.0pt;line-height:115%">Inline formatting</span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">. </span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">2</span></a></span></p>

<p id="B7EE25D1-EE5C-404D-98A8-2F783F6E2925" class="MsoToc2"><span style="font-family:&quot;Lao UI&quot;,sans-serif"><a href="#_Toc359077854"><span style="font-size:12.0pt;line-height:115%">Fun with fonts</span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">. </span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">2</span></a></span></p>

<p id="4BCCABB5-224C-4E86-91F8-5D77DB8488E6" class="MsoToc2"><span style="font-family:&quot;Lao UI&quot;,sans-serif"><a href="#_Toc359077855"><span style="font-size:12.0pt;line-height:115%">Paragraph level formatting</span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">. </span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">2</span></a></span></p>

<p id="7AA6CB4C-85E5-4B37-A313-536DF67B9F45" class="MsoToc1"><span style="font-family:&quot;Lao UI&quot;,sans-serif"><a href="#_Toc359077856"><span style="font-size:12.0pt;line-height:115%">Tables</span><span style="font-size:
12.0pt;line-height:115%;color:windowtext;display:none;text-decoration:none">. </span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">3</span></a></span></p>

<p id="0D67B310-A2F0-4DA7-A3AB-0F98C7E5E5EF" class="MsoToc1"><span style="font-family:&quot;Lao UI&quot;,sans-serif"><a href="#_Toc359077857"><span style="font-size:12.0pt;line-height:115%">Structural Elements</span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">. </span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">5</span></a></span></p>

<p id="F436632D-D987-4A65-87F9-053463461707" class="MsoToc2"><span style="font-family:&quot;Lao UI&quot;,sans-serif"><a href="#_Toc359077858"><span style="font-size:12.0pt;line-height:115%">Footnotes &amp; Endnotes</span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">. </span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">5</span></a></span></p>

<p id="ACC5EFE6-EFB0-4985-B446-D02C8A951D5B" class="MsoToc2"><span style="font-family:&quot;Lao UI&quot;,sans-serif"><a href="#_Toc359077859"><span style="font-size:12.0pt;line-height:115%">Dropcaps</span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">. </span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">5</span></a></span></p>

<p id="E4A3EE5E-3536-46FA-AB6E-026BA108D8E7" class="MsoToc2"><span style="font-family:&quot;Lao UI&quot;,sans-serif"><a href="#_Toc359077860"><span style="font-size:12.0pt;line-height:115%">Links</span><span style="font-size:
12.0pt;line-height:115%;color:windowtext;display:none;text-decoration:none">. </span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">5</span></a></span></p>

<p id="9A360DB0-569F-4011-9B15-7D7CAA428EE3" class="MsoToc2"><span style="font-family:&quot;Lao UI&quot;,sans-serif"><a href="#_Toc359077861"><span style="font-size:12.0pt;line-height:115%">Table of Contents</span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">. </span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">5</span></a></span></p>

<p id="A2AD3EE4-CCC0-40A9-8476-D0516C6A9963" class="MsoToc1"><span style="font-family:&quot;Lao UI&quot;,sans-serif"><a href="#_Toc359077862"><span style="font-size:12.0pt;line-height:115%">Images</span><span style="font-size:
12.0pt;line-height:115%;color:windowtext;display:none;text-decoration:none">. </span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">7</span></a></span></p>

<p id="E81E460A-6D2B-4468-A315-B49ECC9F0BC3" class="MsoToc1"><span style="font-family:&quot;Lao UI&quot;,sans-serif"><a href="#_Toc359077863"><span style="font-size:12.0pt;line-height:115%">Lists</span><span style="font-size:
12.0pt;line-height:115%;color:windowtext;display:none;text-decoration:none">. </span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">8</span></a></span></p>

<p id="4EA5D251-C21A-4716-B653-9CC8A5C7AAE7" class="MsoToc2"><span style="font-family:&quot;Lao UI&quot;,sans-serif"><a href="#_Toc359077864"><span style="font-size:12.0pt;line-height:115%">Bulleted List</span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none"> </span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">8</span></a></span></p>

<p id="CBD376BD-55A6-4F4B-960B-32614D9BBA15" class="MsoToc2"><span style="font-family:&quot;Lao UI&quot;,sans-serif"><a href="#_Toc359077865"><span style="font-size:12.0pt;line-height:115%">Numbered List</span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none"> </span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">8</span></a></span></p>

<p id="380E9F49-491E-4501-978F-69F009FE2B53" class="MsoToc2"><span style="font-family:&quot;Lao UI&quot;,sans-serif"><a href="#_Toc359077866"><span style="font-size:12.0pt;line-height:115%">Multi-level Lists</span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">. </span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">8</span></a></span></p>

<p id="445085D9-9F2C-4661-899A-C6E9710B5ECB" class="MsoToc2"><span style="font-family:&quot;Lao UI&quot;,sans-serif"><a href="#_Toc359077867"><span style="font-size:12.0pt;line-height:115%">Continued Lists</span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">. </span><span style="font-size:12.0pt;line-height:115%;color:windowtext;display:none;
text-decoration:none">8</span></a></span></p>

<p id="EBA4EBDB-F59C-4264-AE5C-E21AAB919050" class="MsoNormal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>

<p id="2A78B801-862E-4AC8-A939-A4AC295A998A" class="MsoNormal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>

<h1 id="294ADFC6-8C37-470B-A32D-4B7249531530"><a name="_Toc359077862"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Images</span></a></h1>

<p id="BD31D73D-EE19-4F6C-86BA-E982FAB3EF3E" class="MsoNormal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Images can be of three main types.
Inline images are images that are part of the normal text flow, like this image
of a green dot <img width="14" height="14" id="Picture 0" src="data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAACAAAAAgCAYAAABzenr0AAAFvUlEQVR42u2WfVBUVRTAz/vYt2/ZBYRVTCAEP0gFv3IEJTUFRBmpNHXGUccUGxPLsskv1JJKpbAmR0vSSSVHzRk0rWQ0FCQrERwTFdT8AmXle0Xc7/d23+u8R4sw7jM/dqb+6DJnz+Pce8/53XPuu/cRoijCv9mI/wH+UwCZR1T/ND4EZRLKBJQuKIP+tpehNKLkoRxAuf0wJ+njec8Aaw8rAgxBSdOqA+dGdI6D7oExoFH5Q2ddD7mzyXwDbHwL3LxTCpVNJ8HiuLMNzdkoZzw5W5msALAmrxVg6tC5bbbc09uytOqAJUPDp0OvrsPB6WoCHsUpmEAUHa1OCDXQpC+oqM5Ao1yrL4bTVXsQpHk9+lrazpesV01QAPjwJ0bW02JTYW/JdumfDeH6wWnx/eaBS2gAzlUPgmBHs9K+IYAkWWCorkCRQVB4cStUGc9KmViEPjn0KY9a/RLnGWD1wVaA6XGpsOfk9s3RoaPTYnumgMNpAF5oeUjgB0FUpD+o6VAouX4Iyg1F2ehzAfpsXehEBYBV3zPux6wwfd8lidEzgXPW4upNjxi4Y6OwLAzdDY6V74Jbxkvr0SSXY82rCgArctWSivFR+5ZMjp0PJGHHet99ouDupqI6gSCysL/ka7A6TLFoKl031eEZYOleGWDHiD5jZ0d26481b3yMtCuXg6G6wJXaC/Db5aM5aJiTNU0B4L3d6ggd63dj8rCZuJk4TL3tKYO7S6HBzcvA/lO7wGy/1+PzGY5KjwCLdrLLo8OiMmN6v4BpM3th9fezQBI6KL36O5TfqkjfMMv+iUeAhTvYgsRBo+LDg3oisXdW724kZqGq4TocKztRuGmOPcEjQNo3rGHKyAkhAdpAPGS4JwqkmAOCgWbLHdj3a97t7NftoR4B5m3ROGYlTGRYhsHke/eSIvDPznGws+Agt/UNm9ojQOpXGsfspFcYkMN7H0D6zcn/gdv+pgLAaxt9DNMTxoewjAo3ocurACRBYQZ42FNw5Pa3b1s9l2DGFz75KcPjxgbr9XjZ8E8USKnRpApqjEY4VHzy6O53rUkeAaZ95rP8+cjIzJi+ffAItnsVgKFZKL10Gf64ciV972Kr59dw6qfaEJ2GNcxOTsIjmMMsOL20ehqPZAZyDueD2WYPzV1maftg6QAwaZ1WUpvjhwxMi44IAxtnBW8cxRrGB8orb0HhmXPS1bzgwArL/d72AC9/LAP017Ls+dSUROx1gYN/ulKoVSyugYIdeQW4etsANF348X0FgJQMnfsxs0dw1+WTxwzDTy0r7gfHYwV1N4ZW46ebD+w/fgpu1NRLdU+X7IcyzJ4Bkj/QtZ+/aVDviLfGDx8ol0LKxKOeDdI7L61cSv2R4nNQdrXySzQvdPcf/kgBYNxKHeSvsxAZGRlETU0NVaX/bmOvZ4PmT4kfBi7gEcSCm/Phr6eKUmFgLVCggn2Fp+BadcMWqmzEOzqdzpmbmyskrdCKP69VAEhM94WR7GLSaDSq6urqpA3h3xJ2YhWj41OTYgfA4Oe640cpJ2eDc0nXtSDPo0gS73xGXjWN+uyfNyG/5DxwFiYnoPrFTEEQ7rEsa6qurnYUFRW5xHZBOwAkLPMFfWUyhYNZnucDse8ZgiCC7DpDlE1XNU7dyR7fLyIUJPH1YUHv7yvPM7aYwGS1w8VKgyyOu+xxjTn8KGsOvYzBjeijDnUTDjVjFnhFgDFLfWG0tmMGSJL0Q61BTdvZ+iCrtmqUk7kXJxJCoEDw3aR5pKiqJUSqmeb8irXW8F8Ya5cGDOpE4Vwul+mRMyABFK03t+2B2tpalZ+fH41jKMwIQdM06XQ6aXRKMwxDoCbccymKkhyhycXjGEGj0Qg4R6qRy2w28+49MHqJTjyedf8j9wEAd5NAJC3BuG0VFRVEY2MjERkZKduam5vb+gICAsTg4GARx4hRUVFtTnG+/CwFdtvaA/wF5UDY3/sPE+wAAAAASUVORK5CYII=" alt="dot_green.png">. Inline images do not cause breaks in the text and are
usually small in size.</span><img width="102" height="102" src="~WRS%7b2BA62031-08BC-4A90-AD18-D031F4C1F354%7d_files/image003.gif" align="left" hspace="12" alt="back.png"><span style="font-family:&quot;Lao UI&quot;,sans-serif">
The next category of image is a floating image, one that “floats “ on the page
and is surrounded by text. Word supports more types of floating images than are
possible with current ebook technology, so the conversion maps floating images
to simple left and right floats, as you can see with the left and right arrow
images on the sides of this paragraph.</span></p>

<p id="331E9094-A759-4B5A-BD7A-0C4F168FD915" class="MsoNormal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>

<p id="D9D4CCAF-F3DC-447E-BE6B-48CA7730D00B" class="MsoNormal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">The final type of image is a “block”
image, one that becomes a paragraph on its own and has no text on either side.
Below is a centered green dot.</span></p>

<p id="97B2AFF6-B153-4CA1-9AB0-A82C93E0F289" class="MsoNormal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>

<p id="902BAF4B-0D9E-4B2C-BB71-DA8244C8C433" class="MsoNormal"><img width="26" height="26" src="~WRS%7b2BA62031-08BC-4A90-AD18-D031F4C1F354%7d_files/image004.gif" alt="dot_green.png"><br clear="ALL">
<img width="102" height="102" src="~WRS%7b2BA62031-08BC-4A90-AD18-D031F4C1F354%7d_files/image005.gif" align="right" hspace="12" alt="forward.png"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Centered
images like this are useful for large pictures that should be a focus of
attention. Generally, it is not possible to translate the exact positioning of
images from a Word document to an ebook. That is because in Word, image
positioning is specified in absolute units from the page boundaries.&nbsp; There is
no analogous technology in ebooks, so the conversion will usually end up
placing the image either centered or floating close to the point in the text
where it was <span class="MsoSubtleEmphasis">inserted</span>, not necessarily
where it appears on the page in Word.</span></p>

<h1 id="0979AE6C-BF3C-447B-B2FB-A52846FC5E7D"><a name="_Toc359077863"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Lists</span></a></h1>

<p id="F5A74C27-A0A0-4862-8909-8EB44A055E88" class="MsoNormal"><span style="font-family:&quot;Lao UI&quot;,sans-serif">All types of lists are supported by the
conversion, with the exception of lists that use fancy bullets, these get
converted to regular bullets.</span></p>

<h2 id="DDC65CB2-E58B-4FC4-A790-94ACC2277CC9"><a name="_Toc359077864"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Bulleted List</span></a></h2>

<p id="6B340EF5-263F-4B85-9716-BB9B3B70EC31" class="MsoNormal">&nbsp;</p>

<p id="5AE6136F-6407-46AF-A461-E13B6C6D3CA3" class="MsoListParagraphCxSpFirst" style="margin-left:.8in;text-indent:-.25in"><span style="font-family:Symbol">·<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span></span><span style="font-family:&quot;Lao UI&quot;,sans-serif">One</span></p>

<p id="A6C27E16-7A34-4853-ACCB-46C49CBDAE64" class="MsoListParagraphCxSpLast" style="margin-left:.8in;text-indent:-.25in"><span style="font-family:Symbol">·<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span></span><span style="font-family:&quot;Lao UI&quot;,sans-serif">Two</span></p>

<h2 id="1C38A49C-44C5-4103-88C6-2BF84300A396"><a name="_Toc359077865"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Numbered List</span></a></h2>

<p id="D7C9E81E-1759-4E51-A07A-C1C677A440D2" class="MsoNormal">&nbsp;</p>

<p id="9FA40E88-7113-45A6-911A-F45BA9A7A695" class="MsoListParagraphCxSpFirst" style="margin-left:.8in;text-indent:-.25in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">1.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp; </span></span><span style="font-family:&quot;Lao UI&quot;,sans-serif">One, with a very long line to
demonstrate that the hanging indent for the list is working correctly</span></p>

<p id="9E2C95A4-5AF5-4054-B646-45001E4871B7" class="MsoListParagraphCxSpLast" style="margin-left:.8in;text-indent:-.25in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">2.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp; </span></span><span style="font-family:&quot;Lao UI&quot;,sans-serif">Two</span></p>

<h2 id="E929D2B3-071F-4687-BA97-CBF4C9ED4CEE"><a name="_Toc359077866"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Multi-level Lists</span></a></h2>

<p id="11BEEF38-D808-4DAE-9069-431AB7354668" class="MsoNormal">&nbsp;</p>

<p id="F411FF7E-C46E-426A-9941-B81893BDFD22" class="MsoListParagraphCxSpFirst" style="text-indent:-.25in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">1.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp; </span></span><span style="font-family:&quot;Lao UI&quot;,sans-serif">One</span></p>

<p id="B477F476-38E8-4DE6-8B3F-8213BD057E00" class="MsoListParagraphCxSpMiddle" style="margin-left:.8in;text-indent:-.3in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">1.1.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp; </span></span><span style="font-family:&quot;Lao UI&quot;,sans-serif">Two</span></p>

<p id="45485FD2-3623-4920-B13A-738D4B7FB8D9" class="MsoListParagraphCxSpMiddle" style="margin-left:1.1in;text-indent:-.35in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">1.1.1.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span></span><span style="font-family:&quot;Lao UI&quot;,sans-serif">Three</span></p>

<p id="797C0A93-7773-45CC-A302-F9494AFC5AF1" class="MsoListParagraphCxSpMiddle" style="margin-left:1.1in;text-indent:-.35in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">1.1.2.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span></span><span style="font-family:&quot;Lao UI&quot;,sans-serif">Four with a very
long line to demonstrate that the hanging indent for the list is working
correctly.</span></p>

<p id="D2F87CAC-F26A-45F7-9077-CC252845EFD4" class="MsoListParagraphCxSpMiddle" style="margin-left:1.1in;text-indent:-.35in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">1.1.3.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span></span><span style="font-family:&quot;Lao UI&quot;,sans-serif">Five</span></p>

<p id="15798DD8-BE66-47D4-8EA2-92256947BFB1" class="MsoListParagraphCxSpMiddle" style="text-indent:-.25in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">2.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp; </span></span><span style="font-family:&quot;Lao UI&quot;,sans-serif">Six</span></p>

<p id="D9810042-03FB-400F-9DFF-BB39F79A419A" class="MsoListParagraphCxSpMiddle" style="text-indent:0in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>

<p id="07577C5A-778D-4F31-AAF3-3FE052F5C698" class="MsoListParagraphCxSpMiddle" style="margin-left:.25in;text-indent:0in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">A
Multi-level list with bullets:</span></p>

<p id="FBC40E21-F037-42A8-842C-128A81E7F093" class="MsoListParagraphCxSpMiddle" style="margin-left:.25in;text-indent:0in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>

<p id="449AF2C1-5840-4EC0-82F3-359C19BB17EE" class="MsoListParagraphCxSpMiddle" style="text-indent:-.25in"><span style="font-family:Wingdings">§<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp; </span></span><span style="font-family:&quot;Lao UI&quot;,sans-serif">One</span></p>

<p id="0CA6CD6B-240E-4DBB-B14B-0C78A77FBDE0" class="MsoListParagraphCxSpMiddle" style="margin-left:.75in;text-indent:-.25in"><span style="font-family:Symbol">·<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span></span><span style="font-family:&quot;Lao UI&quot;,sans-serif">Two</span></p>

<p id="3BE39F0E-A377-4B57-BB61-CA5E26B1130C" class="MsoListParagraphCxSpMiddle" style="margin-left:1.0in;text-indent:-.25in"><span style="font-family:Symbol"><img width="13" height="13" src="~WRS%7b2BA62031-08BC-4A90-AD18-D031F4C1F354%7d_files/image001.gif" alt="*"><span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; </span></span><span style="font-family:&quot;Lao UI&quot;,sans-serif">This bullet uses an image as the bullet
item</span></p>

<p id="54A0BD6A-3949-4846-BF09-F35CF415F10E" class="MsoListParagraphCxSpMiddle" style="margin-left:1.25in;text-indent:-.25in"><span style="font-family:&quot;Courier New&quot;">o<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp; </span></span><span style="font-family:&quot;Lao UI&quot;,sans-serif">Four</span></p>

<p id="9E99634D-1926-47E8-841A-E64AD625E458" class="MsoListParagraphCxSpMiddle" style="text-indent:-.25in"><span style="font-family:Wingdings">§<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp; </span></span><span style="font-family:&quot;Lao UI&quot;,sans-serif">Five</span></p>

<p id="0A5DED0C-8306-4951-8520-FB8EA21D19DA" class="MsoListParagraphCxSpLast" style="text-indent:0in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>

<h2 id="3B236B27-5756-4D52-B067-AEDF54826B8A"><a name="_Toc359077867"><span style="font-family:&quot;Lao UI&quot;,sans-serif">Continued Lists</span></a></h2>

<p id="F6D1CA4F-2FA7-48C9-88D1-2CD2675EDF75" class="MsoListParagraphCxSpFirst" style="margin-left:.8in;text-indent:-.8in"><span style="font-family:&quot;Lao UI&quot;,sans-serif"><span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>i.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span></span><span style="font-family:&quot;Lao UI&quot;,sans-serif">One</span></p>

<p id="3119C747-F2AE-4CF9-B1A5-357FD7DE2BA9" class="MsoListParagraphCxSpMiddle" style="margin-left:.8in;text-indent:-.8in"><span style="font-family:&quot;Lao UI&quot;,sans-serif"><span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>ii.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span></span><span style="font-family:&quot;Lao UI&quot;,sans-serif">Two</span></p>

<p id="CFFD3115-4866-4795-AD6F-9015E2B2665C" class="MsoListParagraphCxSpMiddle" style="margin-left:.8in;text-indent:0in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>

<p id="2A9CEC86-FB16-46A9-AE2C-B481919AF159" class="MsoListParagraphCxSpMiddle" style="text-indent:0in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">An
interruption in our regularly scheduled listing, for this essential and very
relevant public service announcement.</span></p>

<p id="2EC1CCA6-7DC6-4030-9F31-D3FBC1AB5F63" class="MsoListParagraphCxSpMiddle" style="text-indent:0in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>

<p id="D7A2ACAA-04AD-4AD1-9298-882993CF25F9" class="MsoListParagraphCxSpMiddle" style="margin-left:.8in;text-indent:-.8in"><span style="font-family:&quot;Lao UI&quot;,sans-serif"><span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>iii.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span></span><span style="font-family:&quot;Lao UI&quot;,sans-serif">We now resume our
normal programming</span></p>

<p id="C5AF022B-DDCF-4022-8DFF-41B9CF0CACF4" class="MsoListParagraphCxSpLast" style="margin-left:.8in;text-indent:-.8in"><span style="font-family:&quot;Lao UI&quot;,sans-serif"><span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>iv.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span></span><span style="font-family:&quot;Lao UI&quot;,sans-serif">Four</span></p>

<p id="FCC5E32B-DBB6-4F50-AC71-44BE4EBD2A3E" class="MsoNormal" style="margin-left:
1.0in;text-indent:0in"><span style="font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></p>

<p id="A6B927D8-47C7-4CAF-9599-214F573A6FB3" class="MsoNormal"><span class="MsoSubtleEmphasis"><span style="font-family:&quot;Lao UI&quot;,sans-serif">&nbsp;</span></span></p>

<p id="4D71A0D0-6C36-41EC-9DF6-5932FE90EE58" class="MsoNormal">&nbsp;</p>

</div>

<div><br clear="all">

<hr align="left" size="1" width="33%">

<div id="ftn1">

<p id="85869A0E-4E67-4C72-A520-D2EE2AE9D9BC" class="MsoFootnoteText"><a href="#_ftnref1" name="_ftn1" title=""><span class="MsoFootnoteReference"><span class="MsoFootnoteReference"><span style="font-size:10.0pt;line-height:115%;
font-family:&quot;Ubuntu&quot;,sans-serif">[1]</span></span></span></a> In paged media, footnotes
are usually displayed at the bottom of the text. However, in ebooks, a better
paradigm is to make them clickable endnotes that the user can browse at her
pleasure. This conversion is handled automatically by calibre.</p>

</div>

</div>

<div><br clear="all">

<hr align="left" size="1" width="33%">

<div id="edn1">

<p id="D603C590-D942-410F-A35F-F5587D7AA629" class="MsoEndnoteText"><a href="#_ednref1" name="_edn1" title=""><span class="MsoEndnoteReference"><span class="MsoEndnoteReference"><span style="font-size:10.0pt;line-height:115%;
font-family:&quot;Ubuntu&quot;,sans-serif">[i]</span></span></span></a> Endnotes are
typically used for longer notes, they remain endnotes when converted into ebook
form, except that they have an additional backlink to make it easy to return to
the current position after reading the note.</p>

</div>

</div>
`,
  },
  {
    id: 'post-2',
    title: 'Renewable Energy Investment Hits Record High',
    author: 'Daniel Osei',
    publishedAt: '2026-07-25',
    category: 'Energy',
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
	{font-family:Tahoma;
	panose-1:2 11 6 4 3 5 4 4 2 4;}
 /* Style Definitions */
 p.MsoNormal, li.MsoNormal, div.MsoNormal
	{margin-top:6.0pt;
	margin-right:0in;
	margin-bottom:6.0pt;
	margin-left:0in;
	line-height:115%;
	font-size:11.0pt;
	font-family:"Calibri",sans-serif;}
h1
	{mso-style-link:"Heading 1 Char";
	margin-top:6.0pt;
	margin-right:0in;
	margin-bottom:6.0pt;
	margin-left:0in;
	line-height:115%;
	font-size:24.0pt;
	font-family:"Calibri",sans-serif;
	font-weight:bold;}
h2
	{mso-style-link:"Heading 2 Char";
	margin-top:6.0pt;
	margin-right:0in;
	margin-bottom:6.0pt;
	margin-left:0in;
	line-height:115%;
	font-size:18.0pt;
	font-family:"Calibri",sans-serif;
	font-weight:bold;}
h3
	{mso-style-link:"Heading 3 Char";
	margin-top:6.0pt;
	margin-right:0in;
	margin-bottom:6.0pt;
	margin-left:0in;
	line-height:115%;
	font-size:14.0pt;
	font-family:"Calibri",sans-serif;
	font-weight:bold;}
a:link, span.MsoHyperlink
	{color:blue;
	text-decoration:underline;}
p.MsoListParagraph, li.MsoListParagraph, div.MsoListParagraph
	{margin-top:6.0pt;
	margin-right:0in;
	margin-bottom:6.0pt;
	margin-left:.5in;
	line-height:115%;
	font-size:11.0pt;
	font-family:"Calibri",sans-serif;}
p.MsoListParagraphCxSpFirst, li.MsoListParagraphCxSpFirst, div.MsoListParagraphCxSpFirst
	{margin-top:6.0pt;
	margin-right:0in;
	margin-bottom:0in;
	margin-left:.5in;
	line-height:115%;
	font-size:11.0pt;
	font-family:"Calibri",sans-serif;}
p.MsoListParagraphCxSpMiddle, li.MsoListParagraphCxSpMiddle, div.MsoListParagraphCxSpMiddle
	{margin-top:0in;
	margin-right:0in;
	margin-bottom:0in;
	margin-left:.5in;
	line-height:115%;
	font-size:11.0pt;
	font-family:"Calibri",sans-serif;}
p.MsoListParagraphCxSpLast, li.MsoListParagraphCxSpLast, div.MsoListParagraphCxSpLast
	{margin-top:0in;
	margin-right:0in;
	margin-bottom:6.0pt;
	margin-left:.5in;
	line-height:115%;
	font-size:11.0pt;
	font-family:"Calibri",sans-serif;}
span.Heading1Char
	{mso-style-name:"Heading 1 Char";
	mso-style-link:"Heading 1";
	font-weight:bold;}
span.Heading2Char
	{mso-style-name:"Heading 2 Char";
	mso-style-link:"Heading 2";
	font-weight:bold;}
span.Heading3Char
	{mso-style-name:"Heading 3 Char";
	mso-style-link:"Heading 3";
	font-weight:bold;}
.MsoChpDefault
	{font-family:"Calibri",sans-serif;}
.MsoPapDefault
	{margin-bottom:10.0pt;
	line-height:115%;}
@page WordSection1
	{size:8.5in 11.0in;
	margin:1.0in 1.0in .5in 1.0in;}
div.WordSection1
	{page:WordSection1;}
@page WordSection2
	{size:8.5in 11.0in;
	margin:1.0in 1.0in 1.0in 1.0in;}
div.WordSection2
	{page:WordSection2;}
 /* List Definitions */
 ol
	{margin-bottom:0in;}
ul
	{margin-bottom:0in;}
-->
</style>

</head>

<body lang=EN-US link=blue vlink=purple style='word-wrap:break-word'>

<div class=WordSection1>

<h1 id=50428A28-C88F-4F74-812D-867DECDD2420 style='margin-top:0in'><span
style='color:#C00000'>Sample Document</span></h1>

<p id=A772888C-8FD4-4627-BEE8-54D7BFAC4C2D class=MsoNormal>This document was
created using accessibility techniques for headings, lists, image alternate
text, tables, and columns. It should be completely accessible using assistive
technologies such as screen readers.</p>

<h2 id=2F1A4FCA-2BFF-4474-B72F-0FE35FD9F276>Headings</h2>

<p id=E0066C0A-94A2-4684-83EC-721F1E1DD4FB class=MsoNormal>There are eight
section headings in this document. At the beginning, &quot;Sample
Document&quot; is a level 1 heading. The main section headings, such as
&quot;Headings&quot; and &quot;Lists&quot; are level 2 headings. The Tables
section contains two sub-headings, &quot;Simple Table&quot; and &quot;Complex
Table,&quot; which are both level 3 headings.</p>

<h2 id=A18F5209-50F0-4F78-8FDD-4F2E9F9088FB>Lists</h2>

<p id=8E64EB22-8C1F-4CA4-A558-390E4A12B373 class=MsoNormal>The following
outline of the sections of this document is an ordered (numbered) list with six
items. The fifth item, &quot;Tables,&quot; contains a nested unordered
(bulleted) list with two items.</p>

<p id=DABBBBD8-E222-4FD3-9F1E-B022051D0382 class=MsoListParagraphCxSpFirst
style='text-indent:-.25in'>1.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Headings </p>

<p id=CF124995-10CA-4D2B-9897-4AAA820B7F05 class=MsoListParagraphCxSpMiddle
style='text-indent:-.25in'>2.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Lists </p>

<p id=2B95CCB9-13F0-46DE-90CD-4A40661CCDC8 class=MsoListParagraphCxSpMiddle
style='text-indent:-.25in'>3.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Links </p>

<p id=4A89FA65-5A50-4CC0-B4B7-EA8BD857421D class=MsoListParagraphCxSpMiddle
style='text-indent:-.25in'>4.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Images </p>

<p id=ABD39AF5-188D-4A26-9E19-BB5F7ABC931F class=MsoListParagraphCxSpMiddle
style='text-indent:-.25in'>5.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Tables </p>

<p id=B833AE5B-3B90-4F4A-ACBE-9EF87ABB27AA class=MsoListParagraphCxSpMiddle
style='margin-left:.75in;text-indent:-.25in'><span style='font-family:Symbol'>·<span
style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span></span>Simple Tables </p>

<p id=640ED023-E74F-45BF-A3FA-EAA557607D97 class=MsoListParagraphCxSpMiddle
style='margin-left:.75in;text-indent:-.25in'><span style='font-family:Symbol'>·<span
style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span></span>Complex Tables </p>

<p id=81620E9F-98F2-4527-9437-85353E7B1291 class=MsoListParagraphCxSpLast
style='text-indent:-.25in'>6.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Columns </p>

<h2 id=7456C615-42A8-4404-908F-19238EDBB2FA>Links</h2>

<p id=67F62CA5-A724-4D4B-BF5F-300F82F5C8B3 class=MsoNormal>In web documents,
links can point different locations on the page, different pages, or even
downloadable documents, such as Word documents or PDFs:</p>

<p id=ECAFBE42-39A2-4AF6-8877-404DA8F442D8 class=MsoNormal><a href="#_top">Top
of this Page</a><br>
<a href="http://www.dhs.state.il.us/page.aspx?item=67072">Sample Document</a><br>
<a
href="http://www.dhs.state.il.us/OneNetLibrary/27897/documents/Initiatives/IITAA/Sample-Document.docx">Sample
Document (docx)</a></p>

<h2 id=F4D321CE-F707-4DCE-AFE8-BDCD6C5633D4>Images</h2>

<p id=9A877AB8-9B79-4ED5-AF9C-932C5EB21852 class=MsoNormal><img width=141
height=131
src="~WRS%7bF7BAA60F-D2A8-4492-AEDB-80F340C79276%7d_files/image001.gif"
align=left hspace=12 alt="Web Access Symbol">Documents may contain images. For
example, there is an image of the web accessibility symbol to the left of this
paragraph. Its alternate text is &quot;Web Access Symbol&quot;.</p>

<p id=9D9068A5-8CC3-49E8-BCF7-2F50FB18BD51 class=MsoNormal>Alt text should
communicate what an image means, not how it looks.</p>

<br clear=all style='page-break-before:always'>

<p id=511C03C0-7920-4DD6-A89A-8F29AD662233 class=MsoNormal style='margin-top:
0in;margin-right:0in;margin-bottom:10.0pt;margin-left:0in'>&nbsp;</p>

<p id=BD7EC9F9-102E-4476-848A-4D14D756880D class=MsoNormal><img width=284
height=140
src="~WRS%7bF7BAA60F-D2A8-4492-AEDB-80F340C79276%7d_files/image002.png"
align=left hspace=12
alt="Chart of Screen Reader Market Share.&#10;(Unfortunately, there isn't a way in Word or PDF to include rich formatting, such as a table, in alternate text.)">Some
images, such as charts or graphs, require long descriptions, but not all
document types allow that. In web pages, long descriptions may be provided in
several ways: on the page below the image, via a link below the image, or via a
link on the image.</p>

<p id=99964281-EB10-4E2B-83F9-A75E9ECE0720 class=MsoNormal>&nbsp;</p>

<h2 id=4E3E49B7-304A-4963-AA97-882B91CDA400>Tables</h2>

<h3 id=88A23E37-F0B2-4FE7-AAD7-DBF9430B1BB6>Simple Tables</h3>

<p id=21DBCD66-13D6-4797-A537-5209C89DA845 class=MsoNormal>Simple tables have a
uniform number of columns and rows, without any merged cells:</p>

<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0
 style='margin-left:5.4pt;border-collapse:collapse;border:none'>
 <thead>
  <tr>
   <td width=108 valign=top style='width:80.85pt;border:solid windowtext 1.0pt;
   padding:2.9pt 5.75pt 2.9pt 5.75pt'>
   <p id=2E45A501-2DE3-4EA8-9212-D20DA509631F class=MsoNormal style='margin:
   0in;line-height:normal'><b>Screen Reader</b></p>
   </td>
   <td width=85 valign=top style='width:63.5pt;border:solid windowtext 1.0pt;
   border-left:none;padding:2.9pt 5.75pt 2.9pt 5.75pt'>
   <p id=6E76F3B0-41C7-4254-997D-A31146939A1D class=MsoNormal style='margin:
   0in;line-height:normal'><b>Responses</b></p>
   </td>
   <td width=59 valign=top style='width:44.2pt;border:solid windowtext 1.0pt;
   border-left:none;padding:2.9pt 5.75pt 2.9pt 5.75pt'>
   <p id=6DF63D59-18A2-4F05-B6C4-C1A15A373557 class=MsoNormal style='margin:
   0in;line-height:normal'><b>Share </b></p>
   </td>
  </tr>
  <tr>
   <td width=108 valign=top style='width:80.85pt;border:solid windowtext 1.0pt;
   border-top:none;padding:2.9pt 5.75pt 2.9pt 5.75pt'>
   <p id=CF013A9A-D318-4E1A-9952-2541ADC8A5C5 class=MsoNormal style='margin:
   0in;line-height:normal'>JAWS</p>
   </td>
   <td width=85 valign=top style='width:63.5pt;border-top:none;border-left:
   none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
   padding:2.9pt 5.75pt 2.9pt 5.75pt'>
   <p id=F9598005-F514-44B7-A795-52B034114D2D class=MsoNormal style='margin:
   0in;line-height:normal'>853</p>
   </td>
   <td width=59 valign=top style='width:44.2pt;border-top:none;border-left:
   none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
   padding:2.9pt 5.75pt 2.9pt 5.75pt'>
   <p id=B56EE26A-11E2-4084-BA36-6664A88200A4 class=MsoNormal style='margin:
   0in;line-height:normal'>49% </p>
   </td>
  </tr>
  <tr>
   <td width=108 valign=top style='width:80.85pt;border:solid windowtext 1.0pt;
   border-top:none;padding:2.9pt 5.75pt 2.9pt 5.75pt'>
   <p id=0C609F67-9821-49D6-8929-1E408250F089 class=MsoNormal style='margin:
   0in;line-height:normal'>NVDA</p>
   </td>
   <td width=85 valign=top style='width:63.5pt;border-top:none;border-left:
   none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
   padding:2.9pt 5.75pt 2.9pt 5.75pt'>
   <p id=5FA88E88-8D76-470F-833D-B58AA99D3904 class=MsoNormal style='margin:
   0in;line-height:normal'>238</p>
   </td>
   <td width=59 valign=top style='width:44.2pt;border-top:none;border-left:
   none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
   padding:2.9pt 5.75pt 2.9pt 5.75pt'>
   <p id=537692BB-EF56-4640-BA6E-5CC597836BBD class=MsoNormal style='margin:
   0in;line-height:normal'>14% </p>
   </td>
  </tr>
  <tr>
   <td width=108 valign=top style='width:80.85pt;border:solid windowtext 1.0pt;
   border-top:none;padding:2.9pt 5.75pt 2.9pt 5.75pt'>
   <p id=F3548C53-8097-4C7D-A179-BDEB1C037A80 class=MsoNormal style='margin:
   0in;line-height:normal'>Window-Eyes</p>
   </td>
   <td width=85 valign=top style='width:63.5pt;border-top:none;border-left:
   none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
   padding:2.9pt 5.75pt 2.9pt 5.75pt'>
   <p id=39A3F5E7-F101-416D-9057-9E1BAD73E03F class=MsoNormal style='margin:
   0in;line-height:normal'>214</p>
   </td>
   <td width=59 valign=top style='width:44.2pt;border-top:none;border-left:
   none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
   padding:2.9pt 5.75pt 2.9pt 5.75pt'>
   <p id=5B8750C7-41D6-463F-B32F-3227F2EFDD98 class=MsoNormal style='margin:
   0in;line-height:normal'>12% </p>
   </td>
  </tr>
  <tr>
   <td width=108 valign=top style='width:80.85pt;border:solid windowtext 1.0pt;
   border-top:none;padding:2.9pt 5.75pt 2.9pt 5.75pt'>
   <p id=2C164937-3B0A-4667-90F6-FB292431DD78 class=MsoNormal style='margin:
   0in;line-height:normal'>System Access</p>
   </td>
   <td width=85 valign=top style='width:63.5pt;border-top:none;border-left:
   none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
   padding:2.9pt 5.75pt 2.9pt 5.75pt'>
   <p id=0F155A6F-97FF-4841-A872-C0EFF9942DFF class=MsoNormal style='margin:
   0in;line-height:normal'>181</p>
   </td>
   <td width=59 valign=top style='width:44.2pt;border-top:none;border-left:
   none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
   padding:2.9pt 5.75pt 2.9pt 5.75pt'>
   <p id=97E0D35D-F3FB-436D-B56D-DA2C14C04398 class=MsoNormal style='margin:
   0in;line-height:normal'>10% </p>
   </td>
  </tr>
  <tr>
   <td width=108 valign=top style='width:80.85pt;border:solid windowtext 1.0pt;
   border-top:none;padding:2.9pt 5.75pt 2.9pt 5.75pt'>
   <p id=434E74D8-B117-4E6E-A02E-92ADC34282E0 class=MsoNormal style='margin:
   0in;line-height:normal'>VoiceOver</p>
   </td>
   <td width=85 valign=top style='width:63.5pt;border-top:none;border-left:
   none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
   padding:2.9pt 5.75pt 2.9pt 5.75pt'>
   <p id=4ACADE63-1F8C-46AB-BB02-29C739D0DB36 class=MsoNormal style='margin:
   0in;line-height:normal'>159</p>
   </td>
   <td width=59 valign=top style='width:44.2pt;border-top:none;border-left:
   none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
   padding:2.9pt 5.75pt 2.9pt 5.75pt'>
   <p id=3A9306A4-6ED4-4043-AA2C-FBDCF4F81940 class=MsoNormal style='margin:
   0in;line-height:normal'>9% </p>
   </td>
  </tr>
 </thead>
</table>

<p id=9B3BE42C-62FA-4B90-853C-A55B092EC48E class=MsoNormal style='margin:0in;
line-height:50%'>&nbsp;</p>

<h3 id=FCA1678D-0397-4731-AF41-CD21BCF9FE67>Complex Tables</h3>

<p id=D1C89594-C89B-4B7E-872F-08E44E6091BF class=MsoNormal>The following is a
complex table, using merged cells as headers for sections within the table.
This can't be made accessible in all types of documents:</p>

<table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0
 style='margin-left:5.75pt;border-collapse:collapse;border:none'>
 <tr>
  <td width=120 valign=top style='width:1.25in;border:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=0DA60F6A-ED46-4FBF-BC04-3678122DF419 class=MsoNormal style='margin:
  0in;line-height:normal'><b>&nbsp;</b></p>
  </td>
  <td width=255 colspan=2 valign=top style='width:191.5pt;border:solid windowtext 1.0pt;
  border-left:none;padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=DED9F2B0-4824-49B0-A4EA-65C0006D7B8F class=MsoNormal style='margin:
  0in;line-height:normal'><b>May 2012</b></p>
  </td>
  <td width=255 colspan=2 valign=top style='width:191.55pt;border:solid windowtext 1.0pt;
  border-left:none;padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=B2FCD5AD-A7E8-4628-BA81-015C6A1895B3 class=MsoNormal style='margin:
  0in;line-height:normal'><b> September 2010</b></p>
  </td>
 </tr>
 <tr>
  <td width=120 valign=top style='width:1.25in;border:solid windowtext 1.0pt;
  border-top:none;padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=3F00E4CE-6C28-4326-9EC3-94666694DFAD class=MsoNormal style='margin:
  0in;line-height:normal'><b>Screen Reader</b></p>
  </td>
  <td width=128 valign=top style='width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=92CB501E-DC2F-427D-AF76-E77848DEB2B9 class=MsoNormal style='margin:
  0in;line-height:normal'><b>Responses</b></p>
  </td>
  <td width=128 valign=top style='width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=DA8B16A7-EAB1-4077-A1BE-72BB30A3336B class=MsoNormal style='margin:
  0in;line-height:normal'><b>Share</b></p>
  </td>
  <td width=128 valign=top style='width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=458172AB-FB03-4357-80C0-D1DD0EDBFB40 class=MsoNormal style='margin:
  0in;line-height:normal'><b>Responses</b></p>
  </td>
  <td width=128 valign=top style='width:95.8pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=FB1DBA79-E563-422F-A848-7174DCB9E3C7 class=MsoNormal style='margin:
  0in;line-height:normal'><b>Share </b></p>
  </td>
 </tr>
 <tr>
  <td width=120 valign=top style='width:1.25in;border:solid windowtext 1.0pt;
  border-top:none;padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=A4054563-F74D-4443-9B40-B3AECEB3797C class=MsoNormal style='margin:
  0in;line-height:normal'>JAWS</p>
  </td>
  <td width=128 valign=top style='width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=6CCDEFB1-AF01-4696-A039-451318F517B6 class=MsoNormal style='margin:
  0in;line-height:normal'>853</p>
  </td>
  <td width=128 valign=top style='width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=0FF2AF8A-B481-419E-887C-880A5BA1EC49 class=MsoNormal style='margin:
  0in;line-height:normal'>49%</p>
  </td>
  <td width=128 valign=top style='width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=73A763AA-2DAF-421F-8E06-E42D5A7C71C9 class=MsoNormal style='margin:
  0in;line-height:normal'>727</p>
  </td>
  <td width=128 valign=top style='width:95.8pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=DE4A47A0-EDA9-49BD-B16B-B5DE90C04A24 class=MsoNormal style='margin:
  0in;line-height:normal'>59% </p>
  </td>
 </tr>
 <tr>
  <td width=120 valign=top style='width:1.25in;border:solid windowtext 1.0pt;
  border-top:none;padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=DC95E307-068A-4C42-AC12-4A7A60C68663 class=MsoNormal style='margin:
  0in;line-height:normal'>NVDA</p>
  </td>
  <td width=128 valign=top style='width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=8514C1BE-2C6E-4F73-A62D-6435F581F805 class=MsoNormal style='margin:
  0in;line-height:normal'>238</p>
  </td>
  <td width=128 valign=top style='width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=833D3201-FBEF-4010-8415-31471D4AE5EA class=MsoNormal style='margin:
  0in;line-height:normal'>14%</p>
  </td>
  <td width=128 valign=top style='width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=33EEF3BD-5A4D-4382-B037-C856B50AA9E2 class=MsoNormal style='margin:
  0in;line-height:normal'>105</p>
  </td>
  <td width=128 valign=top style='width:95.8pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=7B5C32B7-F35E-496E-A405-4DF6F0C13187 class=MsoNormal style='margin:
  0in;line-height:normal'>9% </p>
  </td>
 </tr>
 <tr>
  <td width=120 valign=top style='width:1.25in;border:solid windowtext 1.0pt;
  border-top:none;padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=863321BB-1E5A-4FE3-9998-78C40A9D45CE class=MsoNormal style='margin:
  0in;line-height:normal'>Window-Eyes</p>
  </td>
  <td width=128 valign=top style='width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=73FF85E1-CA29-4C89-8AF7-E08FD974BE21 class=MsoNormal style='margin:
  0in;line-height:normal'>214</p>
  </td>
  <td width=128 valign=top style='width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=2A5421F7-3357-4012-9B75-2FC19EC2AF59 class=MsoNormal style='margin:
  0in;line-height:normal'>12%</p>
  </td>
  <td width=128 valign=top style='width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=C8208AE6-4D39-4CD5-83E5-21AD152EA50E class=MsoNormal style='margin:
  0in;line-height:normal'>138</p>
  </td>
  <td width=128 valign=top style='width:95.8pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=FE32232C-1F85-48CC-9AC2-14219AA03742 class=MsoNormal style='margin:
  0in;line-height:normal'>11% </p>
  </td>
 </tr>
 <tr>
  <td width=120 valign=top style='width:1.25in;border:solid windowtext 1.0pt;
  border-top:none;padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=AA6A8948-2E37-47BA-BFF2-A7A106180029 class=MsoNormal style='margin:
  0in;line-height:normal'>System Access</p>
  </td>
  <td width=128 valign=top style='width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=1169D08C-7512-4103-87C2-95202C30130D class=MsoNormal style='margin:
  0in;line-height:normal'>181</p>
  </td>
  <td width=128 valign=top style='width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=77BFD201-E181-4EB9-8074-2732D5032E8A class=MsoNormal style='margin:
  0in;line-height:normal'>10%</p>
  </td>
  <td width=128 valign=top style='width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=3B4C7286-A33C-4139-8244-533FD109EBB9 class=MsoNormal style='margin:
  0in;line-height:normal'>58</p>
  </td>
  <td width=128 valign=top style='width:95.8pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=ECC669B5-20B8-4A02-A749-93A441D0E0CE class=MsoNormal style='margin:
  0in;line-height:normal'>5% </p>
  </td>
 </tr>
 <tr>
  <td width=120 valign=top style='width:1.25in;border:solid windowtext 1.0pt;
  border-top:none;padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=88B4C197-B84F-44C6-A52E-9CBA10CB040F class=MsoNormal style='margin:
  0in;line-height:normal'>VoiceOver</p>
  </td>
  <td width=128 valign=top style='width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=DA8202A4-0446-4D5E-AD5E-89B9F50B7BE4 class=MsoNormal style='margin:
  0in;line-height:normal'>159</p>
  </td>
  <td width=128 valign=top style='width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=959BEC99-1DB8-49EE-BC24-5B7214068278 class=MsoNormal style='margin:
  0in;line-height:normal'>9%</p>
  </td>
  <td width=128 valign=top style='width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=0C929078-C37C-499F-A721-54AC012E23B8 class=MsoNormal style='margin:
  0in;line-height:normal'>120</p>
  </td>
  <td width=128 valign=top style='width:95.8pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt'>
  <p id=6B3F9D52-6AE4-4D1C-BC43-23AA574ED874 class=MsoNormal style='margin:
  0in;line-height:normal'>10%</p>
  </td>
 </tr>
</table>

<p id=142BDBEB-2770-4857-9678-348519477806 class=MsoNormal style='margin:0in;
line-height:50%'> </p>

<h2 id=A594533E-0206-4F6F-BDEC-99C665C06F20>Columns</h2>

</div>

<span style='font-size:11.0pt;line-height:115%;font-family:"Calibri",sans-serif'><br
clear=all style='page-break-before:auto'>
</span>

<div class=WordSection2>

<p id=F9272A0B-7DF7-41F2-9774-615F94561317 class=MsoNormal>This is an example
of columns. With columns, the page is split into two or more horizontal
sections. Unlike tables, in which you usually read across a row and then down
to the next, in columns, you read down a column and then across to the next.<br
clear=all>
When columns are not created correctly, screen readers may run lines together,
reading the first line of the first column, then the first line of the second
column, then the second line of the first column, and so on. Obviously, that is
not accessible.</p>

<p id=D28A4404-35E7-4FAC-A898-4B457508029E class=MsoNormal>&nbsp;</p>

</div>

</body>

</html>
 `,
  },
  {
    id: 'post-3',
    title: 'Central Banks Signal Cautious Approach to Rate Cuts',
    author: 'Marta Kowalski',
    publishedAt: '2026-07-21',
    category: 'Rates',
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

<p id=A733BB31-786E-4923-A182-3AAA08CA2587 class=MsoNormal><b>Here is a paragraph
about my findings on this POC.</b></p>

<p id=8F87BDD2-AE51-49B9-A1F5-46E01C770B44 class=MsoListParagraphCxSpFirst
style='margin-left:38.25pt;text-indent:-.25in'>1.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Doesn’t seem like a foolproof way to go</p>

<p id=F4396FE3-FFAA-4668-B22F-7CD9314FC4DD class=MsoListParagraphCxSpMiddle
style='margin-left:38.25pt;text-indent:-.25in'>2.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>We will end up manually handling lot of the word html to json
translations</p>

<p id=40D383A5-6E6A-494A-87B6-A3C0EC1167A9 class=MsoListParagraphCxSpMiddle
style='margin-left:1.0in;text-indent:-.25in'>a.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Unless we use the html as it is</p>

<p id=3C6B6771-988A-4BC0-8338-74A52D770A4F class=MsoListParagraphCxSpMiddle
style='margin-left:1.0in;text-indent:-.25in'>b.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Which will screw up the rendering UIs</p>

<p id=78615748-485D-4A61-9219-78DB53392514 class=MsoListParagraphCxSpMiddle
style='margin-left:38.25pt;text-indent:-.25in'>3.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>I think we should not use word as our editor</p>

<p id=38FD0893-C6F5-42AC-9737-AA748A921ABD class=MsoListParagraphCxSpLast
style='margin-left:38.25pt;text-indent:-.25in'>4.<span style='font:7.0pt "Times New Roman"'>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Rather use the tip tap – enhance it to make it work like a word editor,
so we know what we are dealing with</p>

<p id=F85DD488-CD10-4C8F-8647-AA1762089B3F class=MsoNormal>&nbsp;</p>

</div>

</body>

</html>

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
