import type { JSONContent } from '@tiptap/react';
import { htmlToTiptapJson } from './tiptap-utils/htmlJsonConversion';

export interface FeedPost {
  id: string;
  title: string;
  author: string;
  publishedAt: string;
  category: string;
  html: string;
  json: JSONContent;
}

// Placeholder bodies below approximate what Word's `getHtml()` export tends to
// produce (headings, bold/italic/underline, lists, tables, links). Once
// you've used "Save article as HTML" and copied real captured HTML, replace
// a post's html constant below with the pasted content to see how it renders.
//
// `json` is our single source of truth for rendering a post — Post Feed's News
// format renders straight from it (tiptapJsonToHtml), and "My Articles" loads
// it directly into the editor via DocumentAdapter.setContentJson. `html` stays
// around purely as a backup and for Post Feed's Word format, which
// deliberately renders the raw captured markup as-is.
//
// Each post's `json` is computed from its own `html` constant right below —
// a fixture-authoring convenience (keeps the two in sync without duplicating
// content by hand) rather than hand-written, but every post object here
// genuinely carries both fields, not just html with json bolted on
// afterward.

const post1Html = `<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
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
of a green dot <img width="14" height="14" id="Picture 0" src="data:image/png;base64,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" alt="dot_green.png">. Inline images do not cause breaks in the text and are
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
`;

const post2Html = `

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
span.MsoIntenseEmphasis
	{color:#4F81BD;
	font-style:italic;}
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
p.Style1, li.Style1, div.Style1
	{mso-style-name:Style1;
	mso-style-link:"Style1 Char";
	margin-top:6.0pt;
	margin-right:0in;
	margin-bottom:6.0pt;
	margin-left:0in;
	line-height:115%;
	font-size:18.0pt;
	font-family:"Calibri",sans-serif;
	font-weight:bold;}
span.Style1Char
	{mso-style-name:"Style1 Char";
	mso-style-link:Style1;
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





<div class="WordSection1">

<h1 id="3AE3594D-9776-40BA-9962-4FADC8D99DCF" style="margin-top:0in"><span style="color:#C00000">Sample Document</span></h1>

<p id="638E7FFE-8495-48A6-9D7D-4A296032B605" class="MsoNormal">This document was
created using accessibility techniques for headings, lists, image alternate
text, tables, and columns. It should be completely accessible using assistive
technologies such as screen readers.</p>

<h2 id="6C00B04F-1832-4494-9F82-6F544154BB43">Headings</h2>

<p id="EFECD354-B5BC-4A2E-8DAA-419D67E3C287" class="MsoNormal">There are eight
section headings in this document. At the beginning, "Sample
Document" is a level 1 heading. The main section headings, such as
"Headings" and "Lists" are level 2 headings. The Tables
section contains two sub-headings, "Simple Table" and "Complex
Table," which are both level 3 headings.</p>

<p id="23019E1B-61EE-4E5B-83EA-34745CFE3323" class="Style1">Lists</p>

<p id="93B0A416-87ED-4687-86EE-26C5DF7BF7CA" class="MsoNormal">The following
outline of the sections of this <span class="MsoIntenseEmphasis">document is an
ordered (numbered</span>) list with six items. The fifth item,
"Tables," contains a nested unordered (bulleted) list with two items.</p>

<p id="037DE3B8-6C74-45B8-BF3D-E953CB5E8431" class="MsoListParagraphCxSpFirst" style="text-indent:-.25in">1.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Headings </p>

<p id="3F40C392-A1F8-4661-A3EB-7972AC1C5079" class="MsoListParagraphCxSpMiddle" style="text-indent:-.25in">2.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Lists </p>

<p id="B952EF07-F66F-4E7E-8737-0300A6470E73" class="MsoListParagraphCxSpMiddle" style="text-indent:-.25in">3.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Links </p>

<p id="124D8999-8BE0-4B7E-9896-B6FBFD9A30D3" class="MsoListParagraphCxSpMiddle" style="text-indent:-.25in">4.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Images </p>

<p id="1AEC8CEE-AB17-4C66-A4E0-4193E9C3E4D6" class="MsoListParagraphCxSpMiddle" style="text-indent:-.25in">5.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Tables </p>

<p id="6F60623B-93A9-499F-90D4-12C8FD7BFEAB" class="MsoListParagraphCxSpMiddle" style="margin-left:1.0in;text-indent:-1.0in"><span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>I.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Simple Tables </p>

<p id="F2499026-4FA4-4678-A0B2-C96524D6AEFD" class="MsoListParagraphCxSpMiddle" style="margin-left:1.0in;text-indent:-1.0in"><span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>II.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Complex Tables </p>

<p id="59B3E877-AEE7-441E-9F5F-A24F09AA9B65" class="MsoListParagraphCxSpLast" style="text-indent:-.25in">6.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Columns </p>

<h2 id="46C17EAF-E6D0-4823-8643-2045F458B413">Links</h2>

<p id="5042ED55-EE6F-4471-BB49-0023C0A65129" class="MsoNormal">In web documents,
links can point different locations on the page, different pages, or even
downloadable documents, such as Word documents or PDFs:</p>

<p id="89794476-A320-447D-9314-D1AA053A8074" class="MsoNormal"><a href="#_top">Top
of this Page</a><br>
<a href="http://www.dhs.state.il.us/page.aspx?item=67072">Sample Document</a><br>
<a href="http://www.dhs.state.il.us/OneNetLibrary/27897/documents/Initiatives/IITAA/Sample-Document.docx">Sample
Document (docx)</a></p>

<h2 id="985D9BC4-F0F7-4885-B78E-3390A2D3B96B">Images</h2>

<p id="0F2BFA23-361D-4106-83D4-CAF86B238CE9" class="MsoNormal"><img border="0" width="141" height="131" id="Picture 3" src="data:image/png;base64,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" alt="Web Access Symbol"></p>

<p id="EC6FCAF1-3AC3-4B55-8190-DE32771B681C" class="MsoNormal">Documents may
contain images. For example, there is an image of the web accessibility symbol
to the left of this paragraph. Its alternate text is "Web Access
Symbol".</p>

<p id="2AF31A71-61F8-4062-8691-0BB3E7491C91" class="MsoNormal">Alt text should
communicate what an image means, not how it looks.</p>

<br clear="all" style="page-break-before:always">

<p id="F54F2F4D-1078-48B9-BFC4-AF5FA38CBB00" class="MsoNormal" style="margin-top:
0in;margin-right:0in;margin-bottom:10.0pt;margin-left:0in"><img width="284" height="140" src="~WRS%7bA9E64188-D389-4D6B-9EA4-7B0AEB32B99C%7d_files/image002.png" align="left" hspace="12" alt="Chart of Screen Reader Market Share.
(Unfortunately, there isn't a way in Word or PDF to include rich formatting, such as a table, in alternate text.)">Some
images, such as charts or graphs, require long descriptions, but not all
document types allow that. In web pages, long descriptions may be provided in
several ways: on the page below the image, via a link below the image, or via a
link on the image.</p>

<p id="BB3ACA11-B958-4292-AD16-5B69163AAD46" class="MsoNormal">&nbsp;</p>

<h2 id="D9B6DC88-9C42-42F4-93B0-F99F06CE9B2F">Tables</h2>

<h3 id="65166F07-0F08-457D-B805-B11F58CBDDA8">Simple Tables</h3>

<p id="B9B6832B-0587-4FC4-90BA-0F0C246AF90B" class="MsoNormal">Simple tables have a
uniform number of columns and rows, without any merged cells:</p>

<table class="MsoTableGrid" border="1" cellspacing="0" cellpadding="0" style="margin-left:5.4pt;border-collapse:collapse;border:none">
 <thead>
  <tr>
   <td width="108" valign="top" style="width:80.85pt;border:solid windowtext 1.0pt;
   padding:2.9pt 5.75pt 2.9pt 5.75pt">
   <p id="0E9F4156-7025-4845-A3C2-753D7D803F66" class="MsoNormal" style="margin:
   0in;line-height:normal"><b>Screen Reader</b></p>
   </td>
   <td width="85" valign="top" style="width:63.5pt;border:solid windowtext 1.0pt;
   border-left:none;padding:2.9pt 5.75pt 2.9pt 5.75pt">
   <p id="2A22C849-20E7-490C-AA64-1A2424C3F17E" class="MsoNormal" style="margin:
   0in;line-height:normal"><b>Responses</b></p>
   </td>
   <td width="59" valign="top" style="width:44.2pt;border:solid windowtext 1.0pt;
   border-left:none;padding:2.9pt 5.75pt 2.9pt 5.75pt">
   <p id="778C47FF-40BD-4B7E-9A2E-2D7A0B6B583D" class="MsoNormal" style="margin:
   0in;line-height:normal"><b>Share </b></p>
   </td>
  </tr>
  <tr>
   <td width="108" valign="top" style="width:80.85pt;border:solid windowtext 1.0pt;
   border-top:none;padding:2.9pt 5.75pt 2.9pt 5.75pt">
   <p id="97D456C5-3C37-4254-B001-F6654190EE59" class="MsoNormal" style="margin:
   0in;line-height:normal">JAWS</p>
   </td>
   <td width="85" valign="top" style="width:63.5pt;border-top:none;border-left:
   none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
   padding:2.9pt 5.75pt 2.9pt 5.75pt">
   <p id="B278D178-882E-4443-8E18-34463F4B966B" class="MsoNormal" style="margin:
   0in;line-height:normal">853</p>
   </td>
   <td width="59" valign="top" style="width:44.2pt;border-top:none;border-left:
   none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
   padding:2.9pt 5.75pt 2.9pt 5.75pt">
   <p id="6D502024-A107-417A-93B2-6E85C789FCDB" class="MsoNormal" style="margin:
   0in;line-height:normal">49% </p>
   </td>
  </tr>
  <tr>
   <td width="108" valign="top" style="width:80.85pt;border:solid windowtext 1.0pt;
   border-top:none;padding:2.9pt 5.75pt 2.9pt 5.75pt">
   <p id="C80E1A02-4C37-4C6F-9129-3B2349740177" class="MsoNormal" style="margin:
   0in;line-height:normal">NVDA</p>
   </td>
   <td width="85" valign="top" style="width:63.5pt;border-top:none;border-left:
   none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
   padding:2.9pt 5.75pt 2.9pt 5.75pt">
   <p id="8F7397CA-90DE-49F1-99C8-275E22771976" class="MsoNormal" style="margin:
   0in;line-height:normal">238</p>
   </td>
   <td width="59" valign="top" style="width:44.2pt;border-top:none;border-left:
   none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
   padding:2.9pt 5.75pt 2.9pt 5.75pt">
   <p id="859FB660-E39D-4670-880D-4E9C87EEB424" class="MsoNormal" style="margin:
   0in;line-height:normal">14% </p>
   </td>
  </tr>
  <tr>
   <td width="108" valign="top" style="width:80.85pt;border:solid windowtext 1.0pt;
   border-top:none;padding:2.9pt 5.75pt 2.9pt 5.75pt">
   <p id="F997F24A-92A1-4844-A7FB-CAE982261C53" class="MsoNormal" style="margin:
   0in;line-height:normal">Window-Eyes</p>
   </td>
   <td width="85" valign="top" style="width:63.5pt;border-top:none;border-left:
   none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
   padding:2.9pt 5.75pt 2.9pt 5.75pt">
   <p id="DDE1BEC0-1935-456B-BDB9-B29A194C9A21" class="MsoNormal" style="margin:
   0in;line-height:normal">214</p>
   </td>
   <td width="59" valign="top" style="width:44.2pt;border-top:none;border-left:
   none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
   padding:2.9pt 5.75pt 2.9pt 5.75pt">
   <p id="3CCA5BB7-342D-4709-815A-F4F9E100E8BF" class="MsoNormal" style="margin:
   0in;line-height:normal">12% </p>
   </td>
  </tr>
  <tr>
   <td width="108" valign="top" style="width:80.85pt;border:solid windowtext 1.0pt;
   border-top:none;padding:2.9pt 5.75pt 2.9pt 5.75pt">
   <p id="602AE445-44A4-4397-9466-7F234EE143A8" class="MsoNormal" style="margin:
   0in;line-height:normal">System Access</p>
   </td>
   <td width="85" valign="top" style="width:63.5pt;border-top:none;border-left:
   none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
   padding:2.9pt 5.75pt 2.9pt 5.75pt">
   <p id="3C5AC098-43DC-4924-91E9-2661FB4C766F" class="MsoNormal" style="margin:
   0in;line-height:normal">181</p>
   </td>
   <td width="59" valign="top" style="width:44.2pt;border-top:none;border-left:
   none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
   padding:2.9pt 5.75pt 2.9pt 5.75pt">
   <p id="D6A3AA1D-C54D-489E-90E5-511B09D51C96" class="MsoNormal" style="margin:
   0in;line-height:normal">10% </p>
   </td>
  </tr>
  <tr>
   <td width="108" valign="top" style="width:80.85pt;border:solid windowtext 1.0pt;
   border-top:none;padding:2.9pt 5.75pt 2.9pt 5.75pt">
   <p id="F8E302D2-AAD7-4E55-A19F-56462F552668" class="MsoNormal" style="margin:
   0in;line-height:normal">VoiceOver</p>
   </td>
   <td width="85" valign="top" style="width:63.5pt;border-top:none;border-left:
   none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
   padding:2.9pt 5.75pt 2.9pt 5.75pt">
   <p id="6B4417F9-19ED-4C74-A5EB-64B47CB49326" class="MsoNormal" style="margin:
   0in;line-height:normal">159</p>
   </td>
   <td width="59" valign="top" style="width:44.2pt;border-top:none;border-left:
   none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
   padding:2.9pt 5.75pt 2.9pt 5.75pt">
   <p id="85EF6186-F8F9-4BA4-84DD-D858B780A125" class="MsoNormal" style="margin:
   0in;line-height:normal">9% </p>
   </td>
  </tr>
 </thead>
</table>

<p id="D3FB4488-06E1-4AC2-9CB9-31620049A5F5" class="MsoNormal" style="margin:0in;
line-height:50%">&nbsp;</p>

<h3 id="86C256FD-34AF-408B-BF85-7F91D0474848">Complex Tables</h3>

<p id="A6323DEA-CA19-4B4B-A68B-DE37DD4A1164" class="MsoNormal">The following is a
complex table, using merged cells as headers for sections within the table.
This can't be made accessible in all types of documents:</p>

<table class="MsoTableGrid" border="1" cellspacing="0" cellpadding="0" style="margin-left:5.75pt;border-collapse:collapse;border:none">
 <tbody><tr>
  <td width="120" valign="top" style="width:1.25in;border:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="5B37A472-C62E-4B8F-9C6C-FCF8731710E3" class="MsoNormal" style="margin:
  0in;line-height:normal"><b>&nbsp;</b></p>
  </td>
  <td width="255" colspan="2" valign="top" style="width:191.5pt;border:solid windowtext 1.0pt;
  border-left:none;padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="50285FC9-FFA3-4954-994C-76E10FF202E3" class="MsoNormal" style="margin:
  0in;line-height:normal"><b>May 2012</b></p>
  </td>
  <td width="255" colspan="2" valign="top" style="width:191.55pt;border:solid windowtext 1.0pt;
  border-left:none;padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="64BA4D10-E0EC-40D8-B05B-9E4C4D438E26" class="MsoNormal" style="margin:
  0in;line-height:normal"><b>&nbsp;September 2010</b></p>
  </td>
 </tr>
 <tr>
  <td width="120" valign="top" style="width:1.25in;border:solid windowtext 1.0pt;
  border-top:none;padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="4E6EE737-405D-4C5C-A756-B1F7456E6F3D" class="MsoNormal" style="margin:
  0in;line-height:normal"><b>Screen Reader</b></p>
  </td>
  <td width="128" valign="top" style="width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="D6CEA54F-FCE4-4F04-AACA-FF3761EF5146" class="MsoNormal" style="margin:
  0in;line-height:normal"><b>Responses</b></p>
  </td>
  <td width="128" valign="top" style="width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="C83556D2-D78F-4D3D-BCBF-546BF61A2C74" class="MsoNormal" style="margin:
  0in;line-height:normal"><b>Share</b></p>
  </td>
  <td width="128" valign="top" style="width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="6CE11916-0F86-40CD-B0B4-811576499C6A" class="MsoNormal" style="margin:
  0in;line-height:normal"><b>Responses</b></p>
  </td>
  <td width="128" valign="top" style="width:95.8pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="07A22699-EBB7-491D-B98F-EB41A36B6E1A" class="MsoNormal" style="margin:
  0in;line-height:normal"><b>Share </b></p>
  </td>
 </tr>
 <tr>
  <td width="120" valign="top" style="width:1.25in;border:solid windowtext 1.0pt;
  border-top:none;padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="811BF2FC-8016-42F1-A364-F6DB6548F3A5" class="MsoNormal" style="margin:
  0in;line-height:normal">JAWS</p>
  </td>
  <td width="128" valign="top" style="width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="F7E6EB61-8550-421D-96E7-A60B8F75D4CA" class="MsoNormal" style="margin:
  0in;line-height:normal">853</p>
  </td>
  <td width="128" valign="top" style="width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="C44D77F5-0947-48ED-9790-0C925453953A" class="MsoNormal" style="margin:
  0in;line-height:normal">49%</p>
  </td>
  <td width="128" valign="top" style="width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="17D4412F-9EAD-4718-B19D-8283DE87994B" class="MsoNormal" style="margin:
  0in;line-height:normal">727</p>
  </td>
  <td width="128" valign="top" style="width:95.8pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="3B3374C2-B565-4B02-AA00-161448296DB4" class="MsoNormal" style="margin:
  0in;line-height:normal">59% </p>
  </td>
 </tr>
 <tr>
  <td width="120" valign="top" style="width:1.25in;border:solid windowtext 1.0pt;
  border-top:none;padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="1833E96E-5A25-43A3-9B0D-A6C586B011D4" class="MsoNormal" style="margin:
  0in;line-height:normal">NVDA</p>
  </td>
  <td width="128" valign="top" style="width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="196E13E4-1590-4624-92E3-304A107A203C" class="MsoNormal" style="margin:
  0in;line-height:normal">238</p>
  </td>
  <td width="128" valign="top" style="width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="3452B6BC-7B20-4859-B810-2373914FEBAC" class="MsoNormal" style="margin:
  0in;line-height:normal">14%</p>
  </td>
  <td width="128" valign="top" style="width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="99B9A942-657B-4A89-A24C-EF6AFBFFF763" class="MsoNormal" style="margin:
  0in;line-height:normal">105</p>
  </td>
  <td width="128" valign="top" style="width:95.8pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="437BF62D-7E2F-4592-B66F-A867E5800889" class="MsoNormal" style="margin:
  0in;line-height:normal">9% </p>
  </td>
 </tr>
 <tr>
  <td width="120" valign="top" style="width:1.25in;border:solid windowtext 1.0pt;
  border-top:none;padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="7155F485-8E05-49B1-B960-6AA785F9796C" class="MsoNormal" style="margin:
  0in;line-height:normal">Window-Eyes</p>
  </td>
  <td width="128" valign="top" style="width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="36488B12-BBE9-46A7-B902-64BF3EE654C1" class="MsoNormal" style="margin:
  0in;line-height:normal">214</p>
  </td>
  <td width="128" valign="top" style="width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="3573FAFB-FBB4-49AC-8E70-D8037A64BB79" class="MsoNormal" style="margin:
  0in;line-height:normal">12%</p>
  </td>
  <td width="128" valign="top" style="width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="64FC011B-2A79-40B4-B8C2-1CA6530E034D" class="MsoNormal" style="margin:
  0in;line-height:normal">138</p>
  </td>
  <td width="128" valign="top" style="width:95.8pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="4A70FB94-B57C-4718-8ECE-47DA6A6350E6" class="MsoNormal" style="margin:
  0in;line-height:normal">11% </p>
  </td>
 </tr>
 <tr>
  <td width="120" valign="top" style="width:1.25in;border:solid windowtext 1.0pt;
  border-top:none;padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="8F98DD13-5092-41E0-9498-E238B849089B" class="MsoNormal" style="margin:
  0in;line-height:normal">System Access</p>
  </td>
  <td width="128" valign="top" style="width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="EFE8AB01-DB99-4E54-B4FE-144F34AA9E57" class="MsoNormal" style="margin:
  0in;line-height:normal">181</p>
  </td>
  <td width="128" valign="top" style="width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="E05DDC0E-E829-41BA-B066-45BEFDBA6C83" class="MsoNormal" style="margin:
  0in;line-height:normal">10%</p>
  </td>
  <td width="128" valign="top" style="width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="165EB4E3-0F18-4A5C-AC03-35697663C467" class="MsoNormal" style="margin:
  0in;line-height:normal">58</p>
  </td>
  <td width="128" valign="top" style="width:95.8pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="AB2E8665-09CA-4EF2-864A-4FF2177E38DF" class="MsoNormal" style="margin:
  0in;line-height:normal">5% </p>
  </td>
 </tr>
 <tr>
  <td width="120" valign="top" style="width:1.25in;border:solid windowtext 1.0pt;
  border-top:none;padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="E1E1FA51-144B-4562-831F-03AC0152D1BC" class="MsoNormal" style="margin:
  0in;line-height:normal">VoiceOver</p>
  </td>
  <td width="128" valign="top" style="width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="DBE96C1D-2010-494A-B972-5E020E54AE70" class="MsoNormal" style="margin:
  0in;line-height:normal">159</p>
  </td>
  <td width="128" valign="top" style="width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="FA10009F-CB31-42F1-9D76-0ACF92E156D6" class="MsoNormal" style="margin:
  0in;line-height:normal">9%</p>
  </td>
  <td width="128" valign="top" style="width:95.75pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="F90063D0-76D9-411B-8440-DCA2D327A575" class="MsoNormal" style="margin:
  0in;line-height:normal">120</p>
  </td>
  <td width="128" valign="top" style="width:95.8pt;border-top:none;border-left:
  none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;
  padding:2.9pt 5.75pt 2.9pt 5.75pt">
  <p id="C4FADEFD-A3E4-4338-BD37-F85261C67868" class="MsoNormal" style="margin:
  0in;line-height:normal">10%</p>
  </td>
 </tr>
</tbody></table>

<p id="F493A6C7-F621-4E04-9DC7-5E388A0E55BA" class="MsoNormal" style="margin:0in;
line-height:50%">&nbsp;</p>

<h2 id="50965942-C20D-4A8B-A07A-505EA7812EA5">Columns</h2>

</div>

<span style="font-size:11.0pt;line-height:115%;font-family:&quot;Calibri&quot;,sans-serif"><br clear="all" style="page-break-before:auto">
</span>

<div class="WordSection2">

<p id="C4A6ED94-2360-4E21-97E2-9A9489B116EF" class="MsoNormal">This is an example
of columns. With columns, the page is split into two or more horizontal
sections. Unlike tables, in which you usually read across a row and then down
to the next, in columns, you read down a column and then across to the next.<br clear="all">
When columns are not created correctly, screen readers may run lines together,
reading the first line of the first column, then the first line of the second
column, then the second line of the first column, and so on. Obviously, that is
not accessible.</p>

<p id="FDDA3E82-8C9E-4536-955D-71531BA99653" class="MsoNormal">&nbsp;</p>

</div>




`;

const post3Html = `      <html>

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

`;

const post4Html = `



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
	{font-family:Aptos;}
@font-face
	{font-family:"Aptos Display";}
 /* Style Definitions */
 p.MsoNormal, li.MsoNormal, div.MsoNormal
	{margin-top:0in;
	margin-right:0in;
	margin-bottom:8.0pt;
	margin-left:0in;
	line-height:115%;
	font-size:12.0pt;
	font-family:"Aptos",sans-serif;}
h1
	{mso-style-link:"Heading 1 Char";
	margin-top:.25in;
	margin-right:0in;
	margin-bottom:4.0pt;
	margin-left:0in;
	line-height:115%;
	page-break-after:avoid;
	font-size:20.0pt;
	font-family:"Aptos Display",sans-serif;
	color:#0F4761;
	font-weight:normal;}
h2
	{mso-style-link:"Heading 2 Char";
	margin-top:8.0pt;
	margin-right:0in;
	margin-bottom:4.0pt;
	margin-left:0in;
	line-height:115%;
	page-break-after:avoid;
	font-size:16.0pt;
	font-family:"Aptos Display",sans-serif;
	color:#0F4761;
	font-weight:normal;}
h3
	{mso-style-link:"Heading 3 Char";
	margin-top:8.0pt;
	margin-right:0in;
	margin-bottom:4.0pt;
	margin-left:0in;
	line-height:115%;
	page-break-after:avoid;
	font-size:14.0pt;
	font-family:"Aptos",sans-serif;
	color:#0F4761;
	font-weight:normal;}
h4
	{mso-style-link:"Heading 4 Char";
	margin-top:4.0pt;
	margin-right:0in;
	margin-bottom:2.0pt;
	margin-left:0in;
	line-height:115%;
	page-break-after:avoid;
	font-size:12.0pt;
	font-family:"Aptos",sans-serif;
	color:#0F4761;
	font-weight:normal;
	font-style:italic;}
a:link, span.MsoHyperlink
	{color:blue;
	text-decoration:underline;}
p.MsoListParagraph, li.MsoListParagraph, div.MsoListParagraph
	{margin-top:0in;
	margin-right:0in;
	margin-bottom:8.0pt;
	margin-left:.5in;
	line-height:115%;
	font-size:12.0pt;
	font-family:"Aptos",sans-serif;}
p.MsoListParagraphCxSpFirst, li.MsoListParagraphCxSpFirst, div.MsoListParagraphCxSpFirst
	{margin-top:0in;
	margin-right:0in;
	margin-bottom:0in;
	margin-left:.5in;
	line-height:115%;
	font-size:12.0pt;
	font-family:"Aptos",sans-serif;}
p.MsoListParagraphCxSpMiddle, li.MsoListParagraphCxSpMiddle, div.MsoListParagraphCxSpMiddle
	{margin-top:0in;
	margin-right:0in;
	margin-bottom:0in;
	margin-left:.5in;
	line-height:115%;
	font-size:12.0pt;
	font-family:"Aptos",sans-serif;}
p.MsoListParagraphCxSpLast, li.MsoListParagraphCxSpLast, div.MsoListParagraphCxSpLast
	{margin-top:0in;
	margin-right:0in;
	margin-bottom:8.0pt;
	margin-left:.5in;
	line-height:115%;
	font-size:12.0pt;
	font-family:"Aptos",sans-serif;}
span.Heading1Char
	{mso-style-name:"Heading 1 Char";
	mso-style-link:"Heading 1";
	font-family:"Aptos Display",sans-serif;
	color:#0F4761;}
span.Heading2Char
	{mso-style-name:"Heading 2 Char";
	mso-style-link:"Heading 2";
	font-family:"Aptos Display",sans-serif;
	color:#0F4761;}
span.Heading3Char
	{mso-style-name:"Heading 3 Char";
	mso-style-link:"Heading 3";
	font-family:"Times New Roman",serif;
	color:#0F4761;}
span.Heading4Char
	{mso-style-name:"Heading 4 Char";
	mso-style-link:"Heading 4";
	font-family:"Times New Roman",serif;
	color:#0F4761;
	font-style:italic;}
.MsoPapDefault
	{margin-bottom:8.0pt;
	line-height:115%;}
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

<h1 id="8CF200CE-EC49-4F95-878F-DCA9465069EE">Publisher Heading </h1>

<h1 id="1159F890-1FF0-45D0-A5CE-A3E9B0132834">Publisher Heading </h1>

<h2 id="AFCBCE76-37E9-49EA-83B4-DA1A08517DC2">Publisher Heading </h2>

<h3 id="0FCFAF30-E1C8-420F-B31E-33CB4D0CF3B4">Publisher Heading </h3>

<h4 id="826E029D-818F-47F4-89E3-08B925A0165C">Publisher Heading </h4>

<p id="EDF5A8DD-1660-47F9-BA74-F0CFEDB8627A" class="MsoNormal">&nbsp;</p>

<p id="95E846E7-90FE-440A-9362-4E3873C58D81" class="MsoNormal" align="right" style="text-align:right"><b>A paragraph that is right aligned</b>: This was a
developer-supply-chain breach with unusually efficient distribution. Stolen
maintainer credentials let the attacker publish poisoned Mastra versions tagged
as the latest releases. The malicious easy-day-js dependency then ran during
installation, exposing developer machines and build pipelines to credential
theft and remote code execution. CrowdStrike found that 87% of identified
software-registry threats in the first half of 2026 involved npm packages.</p>

<p id="A60F27E7-9827-4D30-BED8-69B778AEBE8D" class="MsoNormal">&nbsp;</p>

<p id="01C785E7-2984-4F07-BB16-5497EC52AB72" class="MsoNormal"><b>A paragraph that
is Left aligned</b>: This was a developer-supply-chain breach with unusually
efficient distribution. Stolen maintainer credentials let the attacker publish
poisoned Mastra versions tagged as the latest releases. The malicious
easy-day-js dependency then ran during installation, exposing developer
machines and build pipelines to credential theft and remote code execution.
CrowdStrike found that 87% of identified software-registry threats in the first
half of 2026 involved npm packages.</p>

<p id="AEDF2299-D162-4FCE-9974-7D9EA13241E6" class="MsoNormal">&nbsp;</p>

<p id="D224221D-9F97-4D0F-BB2D-4F6DFDDE49C9" class="MsoNormal">List of Items </p>

<p id="AFA2ACC2-7807-4531-A42C-28919636EE4D" class="MsoListParagraphCxSpFirst" style="text-indent:-.25in"><span style="font-family:Symbol">·<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span></span>Item 1</p>

<p id="3CF4FD93-528D-4B5F-930E-BA6155CFC82E" class="MsoListParagraphCxSpMiddle" style="margin-left:1.0in;text-indent:-.25in"><span style="font-family:&quot;Courier New&quot;">o<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp; </span></span>Item 1.1</p>

<p id="A6EBA016-FD88-4768-AC7F-D5A0DA870EE1" class="MsoListParagraphCxSpMiddle" style="margin-left:1.0in;text-indent:-.25in"><span style="font-family:&quot;Courier New&quot;">o<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp; </span></span>Item 1.2</p>

<p id="8655932F-7528-41E5-A277-DADE78F113BE" class="MsoListParagraphCxSpMiddle" style="margin-left:1.5in;text-indent:-.25in">1.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Item 1.2.1</p>

<p id="68F221A3-8874-49A7-AD75-50DA8ACB10CF" class="MsoListParagraphCxSpMiddle" style="margin-left:1.5in;text-indent:-.25in">2.<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span>Item 1.2.2</p>

<p id="90FC0716-8ACD-4A58-A45E-079FDCA215B8" class="MsoListParagraphCxSpMiddle" style="text-indent:-.25in"><span style="font-family:Symbol">·<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span></span>Item 2</p>

<p id="FDB78ED3-0C28-43A8-A5D1-368D713DD38A" class="MsoListParagraphCxSpMiddle" style="text-indent:-.25in"><span style="font-family:Symbol">·<span style="font:7.0pt &quot;Times New Roman&quot;">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
</span></span>Item 3</p>

<p id="888EFB82-C3FE-48DD-9321-42EF4EB245C6" class="MsoListParagraphCxSpLast">&nbsp;</p>

<p id="55B19A07-B387-4DFC-96D8-7459BB4A8206" class="MsoNormal"><b><span style="background:yellow">A paragraph that is written in yellow background</span></b><span style="background:yellow">: This was a developer-supply-chain breach with
unusually efficient distribution. Stolen maintainer credentials let the
attacker publish poisoned Mastra versions tagged as the latest releases. The
malicious easy-day-js dependency then ran during installation, exposing
developer machines and build pipelines to credential theft and remote code
execution. CrowdStrike found that 87% of identified software-registry threats
in the first half of 2026 involved npm packages</span></p>

<p id="13BE76AD-0151-4DDE-9F08-1C027181E987" class="MsoNormal">&nbsp;</p>

<p id="807756E2-BBC2-4E62-BECE-F505F1D84B73" class="MsoNormal"><b>A paragraph that has
Red font color</b>: <span style="color:#EE0000">This was a
developer-supply-chain breach with unusually efficient distribution. Stolen
maintainer credentials let the attacker publish poisoned Mastra versions tagged
as the latest releases. The malicious easy-day-js dependency then ran during
installation, exposing developer machines and build pipelines to credential
theft and remote code execution. CrowdStrike found that 87% of identified
software-registry threats in the first half of 2026 involved npm packages</span></p>

<p id="7C35B71A-C032-487C-8F24-36C449CCC286" class="MsoNormal"><span style="color:#EE0000">&nbsp;</span></p>

<p id="F046B1E7-2E8F-4F54-ABDC-140477ABD346" class="MsoNormal"><b>A paragraph that has
Bold, Italic, Underlined and Strikethorugh</b>: <b>This was a
developer-supply-chain breach with unusually efficient distribution</b>. <i>Stolen
maintainer credentials let the attacker publish poisoned Mastra versions tagged
as the latest releases</i>. <s>The malicious easy-day-js dependency then ran
during installation, exposing developer machines and build pipelines to
credential theft and remote code execution</s>. <u>CrowdStrike found that 87%
of identified software-registry threats in the first half of 2026 involved npm
packages</u></p>

<p id="E3FE8F6B-0AAB-48E0-B101-104C53021384" class="MsoNormal">&nbsp;</p>

<p id="DAD28A60-21B8-4359-8F8C-609BE8697A23" class="MsoNormal"><b><u>Images Testing</u></b></p>

<p id="0A128B83-32A1-4224-9BF1-E37E5E2081F2" class="MsoNormal">Online Picture:</p>

<p id="BAC10EAC-1F48-481C-898C-25B7F65B311F" class="MsoNormal"><img width="370" height="235" id="Picture 5" src="data:image/png;base64,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"></p>

<p id="54A82CCD-6429-48A8-8A48-D212475EB1EA" class="MsoNormal">Icons: </p>

<p id="40C6450D-3B1D-4CD5-8B31-147E0DFC2624" class="MsoNormal"><img width="96" height="96" id="Graphic 6" src="data:image/svg+xml;base64,PHN2ZyBvdmVyZmxvdz0iaGlkZGVuIiB2aWV3Qm94PSIwLCAwLCA5NiwgOTYiIGlkPSJJY29uc19NYXBsZUxlYWYiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyIgeG1sbnM6eGxpbms9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkveGxpbmsiIHhtbG5zOnhtbD0iaHR0cDovL3d3dy53My5vcmcvWE1MLzE5OTgvbmFtZXNwYWNlIiB2ZXJzaW9uPSIxLjEiPjxnIGlkPSJJY29ucyI+PHBhdGggZD0iTTY4LjcxNiA2NS4zNTMgTDkwLjk2NCA1MS4xODkgTDg0Ljg2NCA0OS4xMjQgQzgzLjgxNzYgNDguNzcwNCA4My4yNTU5IDQ3LjYzNTQgODMuNjA5NSA0Ni41ODkgQzgzLjYyNTggNDYuNTQwNyA4My42NDQgNDYuNDkzIDgzLjY2NCA0Ni40NDYgTDg4LjcyNyAzNC41NjEgTDc1Ljk1MSAzNy4zMTEgQzc0Ljg3MTEgMzcuNTQzMyA3My44MDc0IDM2Ljg1NjMgNzMuNTc1MSAzNS43NzY0IEM3My41NjE4IDM1LjcxNDUgNzMuNTUxNCAzNS42NTE5IDczLjU0NCAzNS41ODkgTDcyLjgwMiAyOS4yODEgTDYzLjM1NiAzOS43NjEgQzYyLjgwMTIgNDAuMzc2MiA2MS44NTI3IDQwLjQyNTIgNjEuMjM3NSAzOS44NzA0IEM2MC44NTA1IDM5LjUyMTQgNjAuNjcyMyAzOC45OTY1IDYwLjc2NyAzOC40ODQgTDY1LjE2NyAxNC43MzIgTDU3LjUgMTkuODQ3IEM1Ni41ODExIDIwLjQ1OTkgNTUuMzM5MyAyMC4yMTE4IDU0LjcyNjQgMTkuMjkyOSBDNTQuNjYxMyAxOS4xOTU0IDU0LjYwNSAxOS4wOTI0IDU0LjU1OCAxOC45ODUgTDQ4IDQgTDQxLjQ0NCAxOC45ODUgQzQxLjAwMTIgMTkuOTk2OSAzOS44MjE4IDIwLjQ1ODIgMzguODA5OSAyMC4wMTU0IEMzOC43MDI1IDE5Ljk2ODQgMzguNTk5NSAxOS45MTIgMzguNTAyIDE5Ljg0NyBMMzAuODI5IDE0LjczMiBMMzUuMjI5IDM4LjQ4NCBDMzUuMzc5NSAzOS4yOTg2IDM0Ljg0MTEgNDAuMDgxIDM0LjAyNjQgNDAuMjMxNSBDMzMuNTEzOSA0MC4zMjYyIDMyLjk4OSA0MC4xNDggMzIuNjQgMzkuNzYxIEwyMy4xOTQgMjkuMjgxIEwyMi40NTIgMzUuNTg5IEMyMi4zMjMyIDM2LjY4NiAyMS4zMjk0IDM3LjQ3MDkgMjAuMjMyNCAzNy4zNDIxIEMyMC4xNjk1IDM3LjMzNDcgMjAuMTA2OSAzNy4zMjQzIDIwLjA0NSAzNy4zMTEgTDcuMjY4IDM0LjU2MSBMMTIuMzMxIDQ2LjQ0NiBDMTIuNzYzNiA0Ny40NjIzIDEyLjI5MDMgNDguNjM2OSAxMS4yNzQgNDkuMDY5NSBDMTEuMjI3IDQ5LjA4OTUgMTEuMTc5MyA0OS4xMDc3IDExLjEzMSA0OS4xMjQgTDUuMDMxIDUxLjE4OSBMMjcuMjg0IDY1LjM1MyBDMjguMDk1MyA2NS44Njk0IDI4LjQyNDIgNjYuODkwMSAyOC4wNjcgNjcuNzgzIEwyNC4zOSA3Ni45NzYgTDQ0LjcxMSA3Mi4zIEM0NS4yNDkyIDcyLjE3NjIgNDUuNzg1OSA3Mi41MTIxIDQ1LjkwOTggNzMuMDUwMyBDNDUuOTMwNCA3My4xNCA0NS45Mzg1IDczLjIzMjEgNDUuOTM0IDczLjMyNCBMNDUgOTIgTDUxIDkyIEw1MC4wNjYgNzMuMzIxIEM1MC4wNDA0IDcyLjc2OTMgNTAuNDY2OCA3Mi4zMDEzIDUxLjAxODUgNzIuMjc1NyBDNTEuMTA5NCA3Mi4yNzE0IDUxLjIwMDQgNzIuMjc5NiA1MS4yODkgNzIuMyBMNzEuNjEgNzYuOTc2IEw2Ny45MzMgNjcuNzgzIEM2Ny41NzU4IDY2Ljg5MDEgNjcuOTA0NyA2NS44Njk0IDY4LjcxNiA2NS4zNTMgeiIgLz48L2c+PC9zdmc+" alt="Maple Leaf with solid fill"><img width="96" height="96" id="Graphic 7" src="data:image/svg+xml;base64,PHN2ZyBvdmVyZmxvdz0iaGlkZGVuIiB2aWV3Qm94PSIwLCAwLCA5NiwgOTYiIGlkPSJJY29uc19NYXBsZUxlYWZfTSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIiB4bWxuczp4bGluaz0iaHR0cDovL3d3dy53My5vcmcvMTk5OS94bGluayIgeG1sbnM6eG1sPSJodHRwOi8vd3d3LnczLm9yZy9YTUwvMTk5OC9uYW1lc3BhY2UiIHZlcnNpb249IjEuMSI+PGcgaWQ9Ikljb25zIj48cGF0aCBkPSJNNjguOCA2NS45IEw5Mi4yIDUxIEw4NC4yOSA0OC4zMjIgQzgzLjc4OTcgNDguMTQzIDgzLjUyOTMgNDcuNTkyNCA4My43MDgyIDQ3LjA5MjIgQzgzLjcxMjYgNDcuMDggODMuNzE3MiA0Ny4wNjggODMuNzIyIDQ3LjA1NiBMODkuNCAzMy43MTYgTDc1LjA4OCAzNi44IEM3NC44NDI1IDM2Ljg1MzQgNzQuNTg1OSAzNi44MDY2IDc0LjM3NSAzNi42NyBDNzQuMTQyNCAzNi41MTI3IDczLjk4NzYgMzYuMjY0MSA3My45NDkgMzUuOTg2IEw3Mi45NyAyNy42NTYgTDYyLjMgMzkuNDkgQzYyLjIxODYgMzkuNTgwNCA2Mi4xMDQ1IDM5LjYzNDggNjEuOTgzIDM5LjY0MSBDNjEuODYwOSAzOS42NDgzIDYxLjc0MTEgMzkuNjA1NyA2MS42NTEgMzkuNTIzIEM2MS41MzM1IDM5LjQxNTcgNjEuNDc5NSAzOS4yNTU1IDYxLjUwOCAzOS4wOTkgTDY2LjIyIDEzLjY2MiBMNTYuNzkxIDE5Ljk1MyBDNTYuMzU2NCAyMC4yNDIgNTUuNzY5OSAyMC4xMjUzIDU1LjQ3OSAxOS42OTIgQzU1LjQ0OSAxOS42NDQ5IDU1LjQyMjMgMTkuNTk1OCA1NS4zOTkgMTkuNTQ1IEw0OC4xIDIuODcyIEw0MC44IDE5LjU0NSBDNDAuNTg4NCAyMC4wMjI2IDQwLjAzMDkgMjAuMjM5NyAzOS41NTIgMjAuMDMxIEMzOS41MDI0IDIwLjAwODggMzkuNDU0NSAxOS45ODI3IDM5LjQwOSAxOS45NTMgTDI5Ljk3NSAxMy42NjQgTDM0LjcgMzkuMSBDMzQuNzQ2MSAzOS4zNDkzIDM0LjU4MTQgMzkuNTg4NyAzNC4zMzIxIDM5LjYzNDggQzM0LjMzMTcgMzkuNjM0OSAzNC4zMzE0IDM5LjYzNDkgMzQuMzMxIDM5LjYzNSBDMzQuMTc0MSAzOS42NjU3IDM0LjAxMjcgMzkuNjEwOSAzMy45MDcgMzkuNDkxIEwyMy4yMzQgMjcuNjU0IEwyMi4yNTQgMzUuOTg0IEMyMi4xODAyIDM2LjUxMTggMjEuNjkyNSAzNi44Nzk5IDIxLjE2NDcgMzYuODA2MSBDMjEuMTQ4NCAzNi44MDM4IDIxLjEzMjIgMzYuODAxMSAyMS4xMTYgMzYuNzk4IEw2LjggMzMuNzE2IEwxMi40ODIgNDcuMDU1IEMxMi41ODAyIDQ3LjI4NTkgMTIuNTgyOCA0Ny41NDYzIDEyLjQ4OSA0Ny43NzkgQzEyLjM4MTUgNDguMDMyNyAxMi4xNzI1IDQ4LjIyOTYgMTEuOTEzIDQ4LjMyMiBMNCA1MSBMMjcuNDExIDY1LjkgQzI3Ljc5NTEgNjYuMTQzOCAyNy45NTEgNjYuNjI2NiAyNy43ODIgNjcuMDQ5IEwyMy40OTQgNzcuNzY5IEw0Ny4xIDcyLjMyNyBMNDcuMSA4OC4zNjcgQzQ3LjEgODguOTE5MyA0Ny41NDc3IDg5LjM2NyA0OC4xIDg5LjM2NyBDNDguNjUyMyA4OS4zNjcgNDkuMSA4OC45MTkzIDQ5LjEgODguMzY3IEw0OS4xIDYzLjkgTDY3LjY2NyA1MS4yNTUgQzY4LjEyNCA1MC45NDQxIDY4LjI0MjQgNTAuMzIxNSA2Ny45MzE1IDQ5Ljg2NDUgQzY3LjYyMDYgNDkuNDA3NSA2Ni45OTggNDkuMjg5MSA2Ni41NDEgNDkuNiBMNDkuMSA2MS40NzYgTDQ5LjEgMjkuMzY3IEM0OS4xIDI4LjgxNDcgNDguNjUyMyAyOC4zNjcgNDguMSAyOC4zNjcgQzQ3LjU0NzcgMjguMzY3IDQ3LjEgMjguODE0NyA0Ny4xIDI5LjM2NyBMNDcuMSA2MS40NzYgTDI5LjY2NyA0OS42IEMyOS4yMTA1IDQ5LjI4OTEgMjguNTg4NCA0OS40MDcgMjguMjc3NSA0OS44NjM1IEMyNy45NjY2IDUwLjMyIDI4LjA4NDUgNTAuOTQyMSAyOC41NDEgNTEuMjUzIEw0Ny4xIDYzLjkgTDQ3LjEgNzAuMjc4IEwyNi43NzIgNzQuOTYyIEwyOS42MzkgNjcuNzk0IEMzMC4xNjQxIDY2LjQ3ODkgMjkuNjggNjQuOTc2MiAyOC40ODYgNjQuMjE1IEw4LjYgNTEuNTU2IEwxMi41NjEgNTAuMjE1IEMxMi42MzEgNTAuMTkxIDEyLjY5OSA1MC4xNjUgMTIuNzY3IDUwLjEzNiBDMTQuMjYyNyA0OS40OTcyIDE0Ljk1OTIgNDcuNzY4MyAxNC4zMjQgNDYuMjcxIEwxMC4xNTQgMzYuNDg0IEwyMC43MDQgMzguNzU0IEMyMC43OTMgMzguNzczIDIwLjg4MiAzOC43ODggMjAuOTc1IDM4Ljc5OSBDMjIuNTkwMyAzOC45ODkgMjQuMDUzOSAzNy44MzM1IDI0LjI0MzkgMzYuMjE4MiBDMjQuMjQzOSAzNi4yMTc4IDI0LjI0NCAzNi4yMTc0IDI0LjI0NCAzNi4yMTcgTDI0LjcwOCAzMi4yNzIgTDMyLjQyIDQwLjgyOSBDMzIuOTk0OSA0MS40NTk1IDMzLjg1MyA0MS43NTEgMzQuNjkzIDQxLjYwMSBDMzYuMDI4MyA0MS4zNTMxIDM2LjkxMDcgNDAuMDcwNyAzNi42NjUgMzguNzM1IEwzMi44MTUgMTcuOTUzIEwzOC4zMTUgMjEuNjE3IEMzOC40NTgzIDIxLjcxMjcgMzguNjA5OSAyMS43OTU0IDM4Ljc2OCAyMS44NjQgQzQwLjI1ODEgMjIuNTE0IDQxLjk5MzIgMjEuODM1NCA0Mi42NDcgMjAuMzQ3IEw0OC4xIDcuODYxIEw1My41NjUgMjAuMzQ4IEM1NC4yMTM5IDIxLjgzNzYgNTUuOTQ3NSAyMi41MTkxIDU3LjQzNzEgMjEuODcwMiBDNTcuNTk4NyAyMS43OTk5IDU3Ljc1MzYgMjEuNzE1MSA1Ny45IDIxLjYxNyBMNjMuMzg5IDE3Ljk1NCBMNTkuNTQyIDM4LjczNSBDNTkuMzg5OCAzOS41NzQ5IDU5LjY4MTIgNDAuNDM0IDYwLjMxMyA0MS4wMDggQzYxLjMyMDcgNDEuOTE3NyA2Mi44NzUgNDEuODM4MiA2My43ODQ2IDQwLjgzMDUgQzYzLjc4NTEgNDAuODMgNjMuNzg1NSA0MC44Mjk1IDYzLjc4NiA0MC44MjkgTDcxLjUgMzIuMjczIEw3MS45NjQgMzYuMjIgQzcxLjk3NSAzNi4zMSA3MS45OSAzNi4zOTkgNzIuMDEgMzYuNDkzIEM3Mi4zNTI5IDM4LjA4MzIgNzMuOTE5NSAzOS4wOTQ4IDc1LjUxIDM4Ljc1MyBMODYuMDUzIDM2LjQ4NCBMODEuODg1IDQ2LjI2OSBDODEuODU0MSA0Ni4zMzg5IDgxLjgyNjcgNDYuNDEwMyA4MS44MDMgNDYuNDgzIEM4MS4yODQ1IDQ4LjAyMzcgODIuMTEwNyA0OS42OTM0IDgzLjY1IDUwLjIxNiBMODcuNjA3IDUxLjU1NiBMNjcuNzIxIDY0LjIxNSBDNjYuNTI4IDY0Ljk3NyA2Ni4wNDQxIDY2LjQ3ODkgNjYuNTY4IDY3Ljc5NCBMNjkuNDM2IDc0Ljk2MiBMNTIuMzg2IDcxLjAyNyBDNTEuODQ3OCA3MC45MDI3IDUxLjMxMDggNzEuMjM4MyA1MS4xODY1IDcxLjc3NjUgQzUxLjA2MjIgNzIuMzE0NyA1MS4zOTc4IDcyLjg1MTcgNTEuOTM2IDcyLjk3NiBMNzIuNzEzIDc3Ljc3NiBMNjguNDI2IDY3LjA1MiBDNjguMjU2MyA2Ni42MjgxIDY4LjQxMzYgNjYuMTQzNCA2OC44IDY1LjkgeiIgLz48L2c+PC9zdmc+" alt="Maple Leaf outline"><img width="96" height="96" id="Graphic 8" src="data:image/svg+xml;base64,PHN2ZyB2ZXJzaW9uPSIxLjEiIHZpZXdCb3g9IjAgMCA5NiA5NiIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIiB4bWxuczp4bGluaz0iaHR0cDovL3d3dy53My5vcmcvMTk5OS94bGluayIgaWQ9Ikljb25zX1N1Y2N1bGVudCIgb3ZlcmZsb3c9ImhpZGRlbiI+PHBhdGggZD0iIE0gNTIgMzkuOTQgQyA1NS4yMTc5IDM3LjUwMTQgNTguODU0MyAzNS42NzEzIDYyLjczIDM0LjU0IEMgNjQuNzMwMSAzMC45MzIgNjYuMTY5OSAyNy4wNDA5IDY3IDIzIEMgNjEuMDY5IDIzLjU0OTUgNTUuNTAxIDI2LjA5OSA1MS4yMSAzMC4yMyBDIDUwLjA5NyAzMS4zNjAxIDQ5LjA5MjYgMzIuNTkyMSA0OC4yMSAzMy45MSBDIDQ5LjQ3NjYgMzUuOTg4MyA1MC41NDQ3IDM4LjE4MTQgNTEuNCA0MC40NiBaIi8+PHBhdGggZD0iIE0gMjguOSA4Ny43IEMgMjkuMjM2NyA4OS41OTUgMzAuODc1NCA5MC45ODE2IDMyLjggOTEgTCA2Mi4xIDkxIEMgNjQuMDI0NiA5MC45ODE2IDY1LjY2MzMgODkuNTk1IDY2IDg3LjcgTCA3MC41IDYzIEwgMjQuNSA2MyBaIi8+PHBhdGggZD0iIE0gNzAuNSA1MSBMIDYxLjcyIDUxIEwgNjEuODkgNTAuODcgQyA2OS44OSA0NC41OSA3Mi4xIDM0Ljg3IDcyLjEgMzQuODcgQyA3Mi4xIDM0Ljg3IDYxLjI3IDM1LjI4IDUzLjI1IDQxLjU2IEMgNTEuNjAxMSA0Mi44NTc1IDUwLjExMzggNDQuMzQ4MSA0OC44MiA0NiBDIDQ5Ljk2MDIgNDcuNTc4NCA1MC45NjM5IDQ5LjI1MTEgNTEuODIgNTEgTCA0OS41NSA1MSBDIDQ3LjcwNzcgNDcuMzI2MiA0NS4wOTUgNDQuMDkyNyA0MS44OSA0MS41MiBDIDMzLjg5IDM1LjI0IDIzLjA1IDM0LjgzIDIzLjA1IDM0LjgzIEMgMjMuMDUgMzQuODMgMjUuMjMgNDQuNTkgMzMuMjUgNTAuODMgTCAzMy40MiA1MC45NiBMIDI0LjYgNTAuOTYgQyAyMi4zOTA5IDUwLjk2IDIwLjYgNTIuNzUwOSAyMC42IDU0Ljk2IEwgMjAuNiA1OC45NiBMIDc0LjYgNTguOTYgTCA3NC42IDU1IEMgNzQuNTEwMSA1Mi43OTA1IDcyLjcxMTEgNTEuMDM1MyA3MC41IDUxIFoiLz48cGF0aCBkPSIgTSA0Ny41NyA0NC4zMyBDIDQ4LjI3MzUgNDMuNDgxMyA0OS4wMjEyIDQyLjY3MDEgNDkuODEgNDEuOSBDIDQ4LjMzNzUgMzcuNTU0NCA0NS45NTA2IDMzLjU3NSA0Mi44MSAzMC4yMyBDIDM4LjUyNjggMjYuMTA1MSAzMi45NzA0IDIzLjU1NjEgMjcuMDUgMjMgQyAyNy44NDE5IDI2Ljg5OTUgMjkuMjA2NyAzMC42NjAyIDMxLjEgMzQuMTYgQyAzNS40NDI4IDM1LjI3NDQgMzkuNTIxNCAzNy4yMzkgNDMuMSAzOS45NCBDIDQ0Ljc0OTEgNDEuMjMxOSA0Ni4yNDg1IDQyLjcwNDQgNDcuNTcgNDQuMzMgWiIvPjxwYXRoIGQ9IiBNIDQ3IDMyLjEyIEMgNDcuNzQ2OCAzMS4wNjc3IDQ4LjU2MTkgMzAuMDY1NSA0OS40NCAyOS4xMiBDIDQ4LjU3MDggMjUuOTY5OSA0Ny4yNTA1IDIyLjk2MiA0NS41MiAyMC4xOSBDIDQxLjA1IDEzLjQ1IDM1IDEzIDM1IDEzIEMgMzUuNTA1NiAxNi43NDc1IDM2LjM5MzYgMjAuNDMzNCAzNy42NSAyNCBDIDQwLjA5NTMgMjUuMjI4OCA0Mi4zMjcxIDI2Ljg0MjcgNDQuMjYgMjguNzggQyA0NS4yNjIyIDI5LjgxNzQgNDYuMTc4NSAzMC45MzQ0IDQ3IDMyLjEyIFoiLz48cGF0aCBkPSIgTSA0OS40NiAyMC4yMyBDIDQ5LjE3IDIwLjY3IDQ4LjkgMjEuMTIgNDguNjMgMjEuNTggQyA0OS42MTMzIDIzLjUwNzggNTAuNDMyOCAyNS41MTQ5IDUxLjA4IDI3LjU4IEMgNTMuMDIxOCAyNS44OTc2IDU1LjE5NzkgMjQuNTA2MyA1Ny41NCAyMy40NSBDIDU4LjY5NTcgMjAuMDU0MyA1OS41MTk2IDE2LjU1NDYgNjAgMTMgQyA2MCAxMyA1NCAxMy40NSA0OS40NiAyMC4yMyBaIi8+PHBhdGggZD0iIE0gNDcuMjEgMTkuMTMgQyA0Ny4zMSAxOS4yOCA0Ny4zOSAxOS40NCA0Ny40OSAxOS42IEwgNDcuNzkgMTkuMTMgQyA0OC43OTg2IDE3LjU5NDYgNTAuMDEzMyAxNi4yMDUgNTEuNCAxNSBDIDUwLjcyMzEgMTEuODM3OSA0OS41NTU3IDguODAxMTcgNDcuOTQgNiBDIDQ2LjA3ODMgOC44MTM5OSA0NC42OTMxIDExLjkxNTUgNDMuODQgMTUuMTggQyA0NS4xMjczIDE2LjM0NzQgNDYuMjU5OSAxNy42NzQ5IDQ3LjIxIDE5LjEzIFoiLz48L3N2Zz4=" alt="Succulent with solid fill"></p>

<p id="00B07C48-2467-47EF-B49C-950A9106369F" class="MsoNormal">&nbsp;</p>

<p id="7B42938F-17C6-470A-B4F7-5E2D97FCD426" class="MsoNormal"><img width="96" height="96" id="Graphic 9" src="data:image/svg+xml;base64,PHN2ZyB2aWV3Qm94PSIwIDAgOTYgOTYiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyIgeG1sbnM6eGxpbms9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkveGxpbmsiIGlkPSJJY29uc18zZEdsYXNzZXMiIG92ZXJmbG93PSJoaWRkZW4iPjxwYXRoIGQ9Ik0xMS42MjIgNTUuMzMzIDIwLjE3OCA1OC45QzIxLjM3OCA1OS40IDIyLjM1NiA1OC4yNTggMjIuMzYgNTYuMzQ3TDIyLjM3MSA1MC4xNTdDMjIuMzcxIDQ4LjI0NyAyMS40MDMgNDYuMjkyIDIwLjIwMSA0NS43OTJMMTEuNjQ1IDQyLjIzQzEwLjQ0NSA0MS43MyA5LjQ2NiA0Mi44NzIgOS40NjMgNDQuNzgzTDkuNDUyIDUwLjk3M0M5LjQ0OSA1Mi44NzggMTAuNDIgNTQuODMzIDExLjYyMiA1NS4zMzNaIi8+PHBhdGggZD0iTTM2Ljc5NCA2NS44MTQgNDUuMzUgNjkuMzc2QzQ1LjUyNDYgNjkuNDUxOCA0NS43MTI2IDY5LjQ5MTkgNDUuOTAzIDY5LjQ5NCA0Ni44MzggNjkuNDk0IDQ3LjUyOSA2OC40MyA0Ny41MzIgNjYuODIzTDQ3LjU0MyA2MC42MzNDNDcuNTQzIDU4LjcyMyA0Ni41NzUgNTYuNzY4IDQ1LjM3MyA1Ni4yNjhMMzYuODE4IDUyLjcwNkMzNi42NDM0IDUyLjYzMDMgMzYuNDU1MyA1Mi41OTAyIDM2LjI2NSA1Mi41ODggMzUuMzMgNTIuNTg4IDM0LjYzOCA1My42NTIgMzQuNjM1IDU1LjI1OUwzNC42MjQgNjEuNDQ5QzM0LjYyMSA2My4zNTkgMzUuNTkyIDY1LjMxNCAzNi43OTQgNjUuODE0Wk0zOC42MzEgNTcuNzk0IDQzLjM1NyA1OS43NjFDNDMuNDc1NSA2MC4wMzQyIDQzLjUzODcgNjAuMzI4MiA0My41NDMgNjAuNjI2TDQzLjUzNyA2NC4yODkgMzguODExIDYyLjMyMUMzOC42OTE5IDYyLjA0NzkgMzguNjI4MyA2MS43NTM5IDM4LjYyNCA2MS40NTZaIi8+PHBhdGggZD0iTTMuMjYgMzEuMTcxQzMuMTQ1ODEgMzEuMjE1MyAzLjAzNiAzMS4yNzAzIDIuOTMyIDMxLjMzNSAyLjM3NDggMzEuNzA1OCAyLjA0MDAxIDMyLjMzMDcgMi4wNCAzM0wyIDU1LjQ4MUMxLjk5ODQ4IDU2LjI4OTkgMi40ODQzMiA1Ny4wMiAzLjIzMSA1Ny4zMzFMNS4yNjIgNTguMTc2IDYuNTU2IDYwLjc3NkM2Ljc3MDY3IDYxLjIwNzUgNy4xMzM5NSA2MS41NDcgNy41NzkgNjEuNzMyTDIwLjM2NSA2Ny4wNTZDMjIuODcgNjguMTAxIDI1LjM3MiA2Ny4xMTYgMjYuNzQxIDY0LjU1NkwyOC4xMjcgNjEuOTYyQzI4LjIyNSA2Mi4xMzYxIDI4LjMwODMgNjIuMzE4MSAyOC4zNzYgNjIuNTA2TDMwLjEwNyA2Ny4yMjFDMzEuMTggNzAuMTQyIDMyLjk4MSA3Mi4zMTEgMzUuMDQ4IDczLjE3Mkw0Ny44MzUgNzguNUM0OC42MzU5IDc4LjgzMzQgNDkuNTYwNCA3OC42MTI1IDUwLjEyNCA3Ny45NTNMNTEuMDg0IDc2LjgzIDUyLjc2OCA3Ni45OTIgNTMgNzcgNTMuMjM2IDc2Ljk3N0M1My40MTMxIDc2Ljk1OSA1My41ODcgNzYuOTE3MyA1My43NTMgNzYuODUzTDY2Ljc1MyA3MS41NjdDNjcuMzYwMyA3MS4zMjEgNjcuODA4NiA3MC43OTMxIDY3Ljk1MyA3MC4xNTRMNzAuOTc1IDU2LjczNiA4MS45MjkgNTIuMkM4Mi4wMzc1IDUyLjE1MjIgODIuMTYzMSA1Mi4xNjQ3IDgyLjI2IDUyLjIzMyA4Mi4zNjA3IDUyLjI5NTggODIuNDIwNCA1Mi40MDc0IDgyLjQxNyA1Mi41MjZMODIuNDE3IDUzLjIzOEM4Mi40MTcyIDUzLjM5MTIgODIuMzE4IDUzLjUyNjcgODIuMTcyIDUzLjU3MyA4MC4zNzk4IDU0LjE1NjEgNzkuMTY4IDU1LjgyODMgNzkuMTcyIDU3LjcxM0w3OS4xNzIgNTguNzg5Qzc5LjE3NTcgNjEuMTkyNSA4MS4xMjcxIDYzLjEzOCA4My41MzA2IDYzLjEzNDQgODQuMDY3NSA2My4xMzM2IDg0LjU5OTYgNjMuMDMzNCA4NS4xIDYyLjgzOUw4Ni4wNyA2Mi40NjFDOTAuODY1NSA2MC42MTI2IDk0LjAyMSA1NS45OTQ0IDk0IDUwLjg1NUw5NCA0OC41NzhDOTMuOTk5NCA0My42MTQxIDg5Ljk3NDggMzkuNTkwNSA4NS4wMTA4IDM5LjU5MTIgODMuODczMiAzOS41OTEzIDgyLjc0NiAzOS44MDc0IDgxLjY4OSA0MC4yMjhMNTMuNTI1IDUxLjQzOCAyMS4zMjUgMzguMDMxIDIxLjkyNSAzNS43NTQgMzIuOTI1IDMxLjE5NkMzMy4xMDUxIDMxLjEyMTQgMzMuMzExNiAzMS4yMDcgMzMuMzg2MiAzMS4zODcyIDMzLjQwMzkgMzEuNDI5OSAzMy40MTMgMzEuNDc1NyAzMy40MTMgMzEuNTIyTDMzLjQxMyAzMi4yMzRDMzMuNDEzMiAzMi4zODcyIDMzLjMxNCAzMi41MjI3IDMzLjE2OCAzMi41NjkgMzEuMzc1OCAzMy4xNTIxIDMwLjE2NCAzNC44MjQzIDMwLjE2OCAzNi43MDlMMzAuMTY4IDM3Ljc4NUMzMC4xNjcgNDAuMTg3NCAzMi4xMTM3IDQyLjEzNTggMzQuNTE2MSA0Mi4xMzY5IDM1LjA1OCA0Mi4xMzcxIDM1LjU5NTIgNDIuMDM2MSAzNi4xIDQxLjgzOUwzNy4wNyA0MS40NjFDNDEuODY1NSAzOS42MTI2IDQ1LjAyMSAzNC45OTQ0IDQ1IDI5Ljg1NUw0NSAyNy41NzhDNDQuOTk5NiAyMi42MTQxIDQwLjk3NTIgMTguNTkwMyAzNi4wMTEyIDE4LjU5MDggMzQuODY5NSAxOC41OTA5IDMzLjczODMgMTguODA4NSAzMi42NzggMTkuMjMyWk0zNC4xNjggMjIuOTQ1QzM2LjcyNjggMjEuOTI2IDM5LjYyNzIgMjMuMTc0MyA0MC42NDYyIDI1LjczMzIgNDAuODc5OSAyNi4zMjAxIDQxIDI2Ljk0NjIgNDEgMjcuNTc4TDQxIDI5Ljg1NUM0MS4wMTQ0IDMzLjM0MzggMzguODczIDM2LjQ3OTEgMzUuNjE4IDM3LjczNUwzNC42NDcgMzguMTEzQzM0LjQ2NjcgMzguMTg0MiAzNC4yNjI4IDM4LjA5NTkgMzQuMTkxNiAzNy45MTU2IDM0LjE3NTEgMzcuODc0IDM0LjE2NjggMzcuODI5NyAzNC4xNjcgMzcuNzg1TDM0LjE2NyAzNi43MDlDMzQuMTY2NiAzNi41NTY0IDM0LjI2NDggMzYuNDIxIDM0LjQxIDM2LjM3NCAzNi4yMDUgMzUuNzkzNCAzNy40MiAzNC4xMjA1IDM3LjQxNyAzMi4yMzRMMzcuNDE3IDMxLjUyMkMzNy40MTc2IDI5LjExODUgMzUuNDY5NiAyNy4xNjk1IDMzLjA2NjEgMjcuMTY4OSAzMi40OTQ0IDI3LjE2ODggMzEuOTI4MiAyNy4yODEzIDMxLjQgMjcuNUwxOS40ODQgMzIuNDM4QzE4LjkwODIgMzIuNjc2OCAxOC40NzY3IDMzLjE3MDUgMTguMzE3IDMzLjc3M0wxNy42IDM2LjQ4IDkuMjg2IDMzLjAxOFpNNTAuOTY0IDcyLjggNTAuNDMyIDcyLjc0OUM0OS43ODI1IDcyLjY4NjEgNDkuMTQzMSA3Mi45NDQgNDguNzE5IDczLjQ0TDQ4LjAyOSA3NC4yNDkgMzYuNTg1IDY5LjQ4MkMzNS44MyA2OS4xNjcgMzQuNjYzIDY4LjAyNiAzMy44NjIgNjUuODQ1TDMyLjEzMSA2MS4xM0MzMS4zNTMgNTkuMDEgMjkuODMxIDU3LjYyMSAyOC4xNjYgNTcuNTA1IDI4LjA4OSA1Ny41IDI4LjAxNCA1Ny40OTcgMjcuOTM4IDU3LjQ5NyAyNi42NjM5IDU3LjUyMDQgMjUuNTExNCA1OC4yNTkxIDI0Ljk1OCA1OS40MDdMMjMuMjEzIDYyLjY3NEMyMi42NzEgNjMuNjg4IDIyLjE5IDYzLjQ4OCAyMS45MDMgNjMuMzc0TDkuODAzIDU4LjMzNyA4LjUxMSA1NS43MjdDOC4yOTcwMiA1NS4yOTUzIDcuOTMzOTggNTQuOTU1NyA3LjQ4OSA1NC43NzFMNiA1NC4xNTIgNi4wMzUgMzYgNTEgNTQuNzE4Wk05MCA1MC44NTVDOTAuMDE0NCA1NC4zNDM4IDg3Ljg3MyA1Ny40NzkxIDg0LjYxOCA1OC43MzVMODMuNjQ3IDU5LjExM0M4My40NjY3IDU5LjE4NDIgODMuMjYyOCA1OS4wOTU5IDgzLjE5MTYgNTguOTE1NiA4My4xNzUxIDU4Ljg3NCA4My4xNjY4IDU4LjgyOTcgODMuMTY3IDU4Ljc4NUw4My4xNjcgNTcuNzA5QzgzLjE2NjUgNTcuNTU2NCA4My4yNjQ4IDU3LjQyMSA4My40MSA1Ny4zNzQgODUuMjA1IDU2Ljc5MzQgODYuNDIgNTUuMTIwNSA4Ni40MTcgNTMuMjM0TDg2LjQxNyA1Mi41MjJDODYuNDE3NiA1MC4xMTg1IDg0LjQ2OTYgNDguMTY5NSA4Mi4wNjYxIDQ4LjE2ODkgODEuNDk0NCA0OC4xNjg4IDgwLjkyODIgNDguMjgxMyA4MC40IDQ4LjVMNjguNDg0IDUzLjQzOEM2Ny44ODQxIDUzLjY4NzIgNjcuNDQyNiA1NC4yMTIyIDY3LjMgNTQuODQ2TDY0LjI3OCA2OC4yNTUgNTUgNzIuMDI4IDU1IDU1LjE1NiA4My4xNjggNDMuOTQ1Qzg1LjcyNjggNDIuOTI2IDg4LjYyNzIgNDQuMTc0MyA4OS42NDYyIDQ2LjczMzIgODkuODc5OSA0Ny4zMjAxIDkwIDQ3Ljk0NjIgOTAgNDguNTc4WiIvPjwvc3ZnPg==" alt="3d Glasses with solid fill"><img width="96" height="96" id="Graphic 10" src="data:image/svg+xml;base64,PHN2ZyBvdmVyZmxvdz0iaGlkZGVuIiB2aWV3Qm94PSIwLCAwLCA5NiwgOTYiIGlkPSJJY29uc18zZEdsYXNzZXNfTSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIiB4bWxuczp4bGluaz0iaHR0cDovL3d3dy53My5vcmcvMTk5OS94bGluayIgeG1sbnM6eG1sPSJodHRwOi8vd3d3LnczLm9yZy9YTUwvMTk5OC9uYW1lc3BhY2UiIHZlcnNpb249IjEuMSI+PGcgaWQ9Ikljb25zIj48cGF0aCBkPSJNMTEuNjIxOSA1NS4zMzY5IEwyMC4xNzc2IDU4Ljg5OTIgQzIwLjM1MjMgNTguOTc0NyAyMC41NDA0IDU5LjAxNDcgMjAuNzMwNyA1OS4wMTY4IEMyMS42NjUzIDU5LjAxNjggMjIuMzU2OSA1Ny45NTI5IDIyLjM1OTcgNTYuMzQ2MiBMMjIuMzcwNyA1MC4xNTYyIEMyMi4zNzQxIDQ4LjI0NjIgMjEuNDAyNiA0Ni4yOTE1IDIwLjIwMDcgNDUuNzkxMiBMMTEuNjQ1MSA0Mi4yMjg3IEMxMS40NzA0IDQyLjE1MzMgMTEuMjgyNCA0Mi4xMTM0IDExLjA5MjEgNDIuMTExMyBDMTAuMTU3NSA0Mi4xMTEzIDkuNDY1OCA0My4xNzUyIDkuNDYyOSA0NC43ODE5IEw5LjQ1MTkgNTAuOTcxOSBDOS40NDg3IDUyLjg4MjEgMTAuNDIgNTQuODM2NCAxMS42MjE5IDU1LjMzNjkgeiBNMTEuNDUxOSA1MC44Nzc0IEwxNi4wODU0IDQ2LjI0MzcgTDE3Ljk3MTMgNDcuMDI4OCBMMTEuOTkzMSA1My4wMDc4IEMxMS42MzggNTIuMzg5MyAxMS40NTE0IDUxLjY4ODUgMTEuNDUxNyA1MC45NzUzIHogTTIwLjM3MDMgNTAuMTUyMyBMMjAuMzcwMyA1MC4yMTUzIEwxNS43MTE5IDU0Ljg3MzggTDEzLjc2NTcgNTQuMDYzOCBMMTkuNzgxNyA0OC4wNDY5IEMyMC4xNjgyIDQ4LjY4MTEgMjAuMzcyIDQ5LjQwOTYgMjAuMzcwNyA1MC4xNTIzIHogTTIwLjMyNjYgNTYuNzk0OSBMMTcuNzA4NiA1NS43MDQ5IEwyMC4zNjUxIDUzLjA0ODQgTDIwLjM2IDU2LjM0MjUgQzIwLjM2MDYgNTYuNDkzOSAyMC4zNDk2IDU2LjY0NTIgMjAuMzI3IDU2Ljc5NDkgeiBNMTEuNDk1OCA0NC4zMzI4IEwxNC4wODkxIDQ1LjQxMjggTDExLjQ1NzMgNDguMDQ0NiBMMTEuNDYzMSA0NC43ODU4IEMxMS40NjI0IDQ0LjYzNDIgMTEuNDczMyA0NC40ODI3IDExLjQ5NTggNDQuMzMyOCB6IiAvPjxwYXRoIGQ9Ik0zNi43OTQzIDY1LjgxODEgTDQ1LjM1IDY5LjM4IEM0NS41MjQ3IDY5LjQ1NTUgNDUuNzEyOCA2OS40OTU1IDQ1LjkwMzEgNjkuNDk3NiBDNDYuODM3NyA2OS40OTc2IDQ3LjUyOTMgNjguNDMzNCA0Ny41MzIxIDY2LjgyNyBMNDcuNTQzMiA2MC42MzcgQzQ3LjU0NjUgNTguNzI3IDQ2LjU3NTEgNTYuNzcyMyA0NS4zNzMyIDU2LjI3MTggTDM2LjgxNzYgNTIuNzEgQzM2LjY0MjggNTIuNjM0NSAzNi40NTQ4IDUyLjU5NDUgMzYuMjY0NCA1Mi41OTIzIEMzNS4zMjk4IDUyLjU5MjMgMzQuNjM4MyA1My42NTYzIDM0LjYzNTQgNTUuMjYzIEwzNC42MjQ0IDYxLjQ1MyBDMzQuNjIxIDYzLjM2MzMgMzUuNTkyNSA2NS4zMTc2IDM2Ljc5NDMgNjUuODE4MSB6IE0zNi42MzU0IDU1LjI2NjQgQzM2LjYzNDcgNTUuMTE1IDM2LjY0NTcgNTQuOTYzNyAzNi42NjgxIDU0LjgxNCBMNDQuNTkzNyA1OC4xMTQgQzQ1LjI0MTIgNTguNzg5IDQ1LjU4NDIgNTkuNjk5NiA0NS41NDMxIDYwLjYzNCBMNDUuNTMyIDY2LjgyNCBDNDUuNTMyNyA2Ni45NzU0IDQ1LjUyMTggNjcuMTI2NyA0NS40OTkzIDY3LjI3NjQgTDM3LjU3MzcgNjMuOTc2NCBDMzYuOTI2MyA2My4zMDEzIDM2LjU4MzMgNjIuMzkwOCAzNi42MjQ0IDYxLjQ1NjQgeiIgLz48cGF0aCBkPSJNOTAuODE4NiA0My4wOTI1IEM4OC41NzA0IDQwLjcxNDIgODUuMDk4NSAzOS45NDg3IDgyLjA1ODYgNDEuMTYxMSBMNTMuNDkyMiA1Mi41MzIgQzUzLjQ1NzcgNTIuNTA3OSA1My40MjE3IDUyLjQ4NjEgNTMuMzg0MyA1Mi40NjY4IEwyMC4xMzU3IDM4LjYyMjggTDIxLjA4OTggMzUuMDI0NCBMMzIuNTQ2NCAzMC4yNzY0IEMzMy4yMzU3IDI5Ljk5IDM0LjAyNjYgMzAuMzE2NyAzNC4zMTMgMzEuMDA2IEMzNC4zODE0IDMxLjE3MDYgMzQuNDE2NSAzMS4zNDcxIDM0LjQxNjQgMzEuNTI1NCBMMzQuNDE2NCAzMi4yMDg5IEMzNC40MzE0IDMyLjc5MTEgMzQuMDcwMiAzMy4zMTY5IDMzLjUyMTQgMzMuNTExOCBDMzIuMTE5MiAzMy45NDY3IDMxLjE2NDEgMzUuMjQ0OCAzMS4xNjY0IDM2LjcxMjkgTDMxLjE2NjQgMzcuNjgzNCBDMzEuMTUwNyAzOC40MTIxIDMxLjM2NDcgMzkuMTI3MiAzMS43NzggMzkuNzI3NSBDMzIuNjcwOSA0MC45ODM1IDM0LjMwMzQgNDEuNDcxNyAzNS43MzkyIDQwLjkxMjEgTDM2LjcwNzkgNDAuNTMzMiBDNDEuMTA1OCAzOC44MTMxIDQ0IDM0LjU3MzUgNDQgMjkuODUxMiBMNDQgMjcuODQxOCBDNDQuMDM2MyAyNS43MTIzIDQzLjI1MjIgMjMuNjUwNSA0MS44MSAyMi4wODMzIEMzOS41NTkgMTkuNzExNSAzNi4wODk5IDE4Ljk1MDkgMzMuMDUzMSAyMC4xNjMzIEwzLjYzMyAzMi4wOTgxIEMzLjU4MTk2IDMyLjExODYgMy41MzI3MyAzMi4xNDMyIDMuNDg1OCAzMi4xNzE5IEMzLjIwNzggMzIuMzU2NiAzLjA0MDUxIDMyLjY2ODEgMy4wNCAzMy4wMDE5IEwzIDU1LjQ4NjMgQzIuOTk5MzMgNTUuODkwNiAzLjI0MjE0IDU2LjI1NTUgMy42MTUzIDU2LjQxMSBMNS45OTA3IDU3LjQgTDcuNDUxNyA2MC4zNCBDNy41NTg2NiA2MC41NTU4IDcuNzQwMyA2MC43MjU1IDcuOTYyOSA2MC44MTc2IEwyMC43NSA2Ni4xNDE2IEMyMi43NzQ5IDY2Ljk4NzMgMjQuNzMyOSA2Ni4xOTkyIDI1Ljg1OTQgNjQuMDkwOCBMMjcuNjA1IDYwLjgyMzIgQzI3LjY4NiA2MC42Njk5IDI3LjgyNSA2MC41MTU2IDI3Ljk1OCA2MC41MDEgQzI4LjI1MzkgNjAuNTIxNSAyOC44ODM4IDYwLjk5MjIgMjkuMzE1NCA2Mi4xNjggTDMxLjA0NjQgNjYuODgyOCBDMzIuMDIzIDY5LjU0MjggMzMuNjIyMSA3MS41MDEgMzUuNDMzMSA3Mi4yNTQ5IEw0OC4yMTk3IDc3LjU3OTEgQzQ4LjYyMDEgNzcuNzQ2NCA0OS4wODI3IDc3LjYzNTkgNDkuMzY0MyA3Ny4zMDU3IEw1MC42NjIxIDc1Ljc4NzEgTDUyLjg2NDMgNzUuOTk5IEM1Mi44ODQzIDc2LjAwMSA1Mi45MDMxIDc1Ljk5MTcgNTIuOTIzMiA3NS45OTI0IEM1Mi45NDkzIDc1Ljk5NDYgNTIuOTc0MyA3Ni4wMDM5IDUzLjAwMDUgNzYuMDAzOSBDNTMuMDgxNyA3NS45OTgxIDUzLjE2MiA3NS45ODI1IDUzLjIzOTUgNzUuOTU3NSBDNTMuMjU0MSA3NS45NTMxIDUzLjI2OTUgNzUuOTUzNiA1My4yODM3IDc1Ljk0ODcgQzUzLjMxNTEgNzUuOTQ0MyA1My4zNDYyIDc1LjkzODMgNTMuMzc3IDc1LjkzMDcgTDY2LjM3NyA3MC42NDQ2IEM2Ni42ODA0IDcwLjUyMTIgNjYuOTA0MyA3MC4yNTcxIDY2Ljk3NjMgNjkuOTM3NSBMNzAuMTEyNSA1Ni4wMTU2IEw4MS41NDcxIDUxLjI3NjQgQzgyLjIzNjQgNTAuOTkgODMuMDI3MyA1MS4zMTY3IDgzLjMxMzcgNTIuMDA2IEM4My4zODIxIDUyLjE3MDYgODMuNDE3MiA1Mi4zNDcxIDgzLjQxNzEgNTIuNTI1NCBMODMuNDE3MSA1My4yMDg5IEM4My40MzIxIDUzLjc5MTEgODMuMDcwOSA1NC4zMTY5IDgyLjUyMjEgNTQuNTExOCBDODEuMTE5OSA1NC45NDY3IDgwLjE2NDggNTYuMjQ0OCA4MC4xNjcxIDU3LjcxMjkgTDgwLjE2NzEgNTguNzM4MyBDODAuMTMwMiA2MC41ODM3IDgxLjU5NjMgNjIuMTA5NiA4My40NDE3IDYyLjE0NjUgQzgzLjg4NTcgNjIuMTU1NCA4NC4zMjcgNjIuMDc1NyA4NC43Mzk5IDYxLjkxMjEgTDg1LjcwODYgNjEuNTMzMiBDOTAuMTA2MiA1OS44MTI5IDkzLjAwMDEgNTUuNTczMyA5MyA1MC44NTEyIEw5MyA0OC44NDYxIEM5My4wMzY2IDQ2LjcxOTYgOTIuMjU1OCA0NC42NjAyIDkwLjgxODYgNDMuMDkyNSB6IE0zMy43OTg2IDIyLjAxODYgQzM1LjcxNzIgMjEuMjQwOSAzNy45MDI2IDIxLjUyMTggMzkuNTYyMSAyMi43NTk2IEM0MS4xMjg1IDIzLjk1ODIgNDIuMDMzMSAyNS44Mjg5IDQyIDI3LjgwMSBMNDIgMjkuODUgQzQxLjk5OTkgMzMuNzQ4OCAzOS42MTA3IDM3LjI0OTIgMzUuOTggMzguNjcgTDM1LjAxMTcgMzkuMDQ4OSBDMzQuMzE1MyAzOS4zMjA1IDMzLjUzMDYgMzguOTc2MiAzMy4yNTkgMzguMjc5OCBDMzMuMTk3OSAzOC4xMjMxIDMzLjE2NjUgMzcuOTU2NCAzMy4xNjY1IDM3Ljc4ODIgTDMzLjE2NjUgMzYuNzQyMiBDMzMuMTUxNSAzNi4xNiAzMy41MTI3IDM1LjYzNDIgMzQuMDYxNSAzNS40Mzk0IEMzNS40NjM3IDM1LjAwNDUgMzYuNDE4OCAzMy43MDY0IDM2LjQxNjUgMzIuMjM4MyBMMzYuNDE2NSAzMS42NzA1IEMzNi40MzY2IDMwLjc5MDYgMzYuMTI3MSAyOS45MzUgMzUuNTQ4NyAyOS4yNzE2IEMzNC41OTYzIDI4LjIyMzUgMzMuMDg4NCAyNy44ODYzIDMxLjc4MDMgMjguNDI4NyBMMTkuODY3MiAzMy4zNjYyIEMxOS41NzkxIDMzLjQ4NTUgMTkuMzYzMiAzMy43MzI1IDE5LjI4MzMgMzQuMDMzOSBMMTguMjcyMiAzNy44NDY5IEw2LjY3NDkgMzMuMDE4MiBDNi42NjI2IDMzLjAxMyA2LjY2MjYgMzMuMDA0NyA2LjY3NDkgMzIuOTk5NiB6IE01MS45NjE5IDczLjkwMzMgTDUwLjMzNTkgNzMuNzQ2MSBDNTAuMDExMiA3My43MTQ3IDQ5LjY5MTYgNzMuODQzNyA0OS40Nzk2IDc0LjA5MTcgTDQ4LjMxNTkgNzUuNDUzMSBMMzYuMjAxMiA3MC40MDkyIEMzNC45NDEyIDY5Ljg4NDggMzMuNjg1NSA2OC4yNjkyIDMyLjkyMzggNjYuMTkzNCBMMzEuMTkyNCA2MS40Nzg1IEMzMC41NTE4IDU5LjczMzQgMjkuMzY1MiA1OC41OTM4IDI4LjA5NjIgNTguNTA1OSBDMjguMDQyNSA1OC41MDIgMjcuOTg5MiA1OC41MDAxIDI3LjkzNjIgNTguNSBDMjcuMDMxMSA1OC41MjMyIDI2LjIxODEgNTkuMDU5MiAyNS44NCA1OS44ODE4IEwyNC4wOTQ5IDYzLjE0ODQgQzIzLjc0NDEgNjQuMTU3MyAyMi42NDE4IDY0LjY5MDkgMjEuNjMyOSA2NC4zNDAxIEMyMS41OTQxIDY0LjMyNjYgMjEuNTU1NyA2NC4zMTE4IDIxLjUxNzggNjQuMjk1OSBMOS4wNzYyIDU5LjExNTIgTDcuNjE2MiA1Ni4xNzcyIEM3LjUwODg2IDU1Ljk2MTEgNy4zMjcwNyA1NS43OTEgNy4xMDQzIDU1LjY5ODIgTDUuMDAxIDU0LjgyMjMgTDUuMDM3NiAzNC41MTY5IEM1LjAzNjk5IDM0LjUxMTggNS4wNDA2NCAzNC41MDcyIDUuMDQ1NzQgMzQuNTA2NiBDNS4wNDc3IDM0LjUwNjMgNS4wNDk2OCAzNC41MDY3IDUuMDUxNCAzNC41MDc3IEw1MS45OTg1IDU0LjA1NTcgeiBNOTEgNTAuODUgQzkwLjk5OTkgNTQuNzQ4OCA4OC42MTA3IDU4LjI0OTIgODQuOTggNTkuNjcgTDg0LjAxMTcgNjAuMDQ4OSBDODMuMzE1MyA2MC4zMjA1IDgyLjUzMDYgNTkuOTc2MiA4Mi4yNTkgNTkuMjc5OCBDODIuMTk3OSA1OS4xMjMxIDgyLjE2NjUgNTguOTU2NCA4Mi4xNjY1IDU4Ljc4ODIgTDgyLjE2NjUgNTcuNzQyMiBDODIuMTUxNSA1Ny4xNiA4Mi41MTI3IDU2LjYzNDIgODMuMDYxNSA1Ni40Mzk0IEM4NC40NjM3IDU2LjAwNDUgODUuNDE4OCA1NC43MDY0IDg1LjQxNjUgNTMuMjM4MyBMODUuNDE2NSA1Mi42NzIyIEM4NS40MzY3IDUxLjc5MzIgODUuMTI4MyA1MC45MzgzIDg0LjU1MTYgNTAuMjc0NyBDODMuNTk5NiA0OS4yMjQxIDgyLjA4OTkgNDguODg1NCA4MC43ODAzIDQ5LjQyODcgTDY4Ljg2NzEgNTQuMzY2MiBDNjguNTY2OSA1NC40OTA2IDY4LjM0NTggNTQuNzUzMiA2OC4yNzQ0IDU1LjA3MDMgTDY1LjEzOTIgNjguOTg4MyBMNTQgNzMuNTE3NiBMNTQgNTQuNDgyNCBMODIuNzk4MyA0My4wMTg2IEM4NC43MTcyIDQyLjI0MDcgODYuOTAzMSA0Mi41MjE5IDg4LjU2MjYgNDMuNzYwMiBDOTAuMTI4NyA0NC45NTg5IDkxLjAzMzEgNDYuODI5NSA5MSA0OC44MDE0IHoiIC8+PC9nPjwvc3ZnPg==" alt="3d Glasses outline"><img width="96" height="96" id="Graphic 11" src="data:image/svg+xml;base64,PHN2ZyB2aWV3Qm94PSIwIDAgOTYgOTYiIHhtbG5zPSJodHRwOi8vd3d3LnczLm9yZy8yMDAwL3N2ZyIgeG1sbnM6eGxpbms9Imh0dHA6Ly93d3cudzMub3JnLzE5OTkveGxpbmsiIGlkPSJJY29uc19BYmFjdXMiIG92ZXJmbG93PSJoaWRkZW4iPjxwYXRoIGQ9Ik03NiA3IDIwIDdDMTcuNzkwOSA3IDE2IDguNzkwODYgMTYgMTFMMTYgODVDMTYgODcuMjA5MSAxNy43OTA5IDg5IDIwIDg5TDc2IDg5Qzc4LjIwOTEgODkgODAgODcuMjA5MSA4MCA4NUw4MCAxMUM4MCA4Ljc5MDg2IDc4LjIwOTEgNyA3NiA3Wk03NCAxMyA3NCAyMiA2OCAyMkM2OCAxOS43OTA5IDY2LjIwOTEgMTggNjQgMTggNjEuNzkwOSAxOCA2MCAxOS43OTA5IDYwIDIyTDU4IDIyQzU4IDE5Ljc5MDkgNTYuMjA5MSAxOCA1NCAxOCA1MS43OTA5IDE4IDUwIDE5Ljc5MDkgNTAgMjJMMzYgMjJDMzYgMTkuNzkwOSAzNC4yMDkxIDE4IDMyIDE4IDI5Ljc5MDkgMTggMjggMTkuNzkwOSAyOCAyMkwyMiAyMiAyMiAxM1pNNjQgNjZDNjEuNzkwOSA2NiA2MCA2Ny43OTA5IDYwIDcwTDU4IDcwQzU4IDY3Ljc5MDkgNTYuMjA5MSA2NiA1NCA2NiA1MS43OTA5IDY2IDUwIDY3Ljc5MDkgNTAgNzBMMzYgNzBDMzYgNjcuNzkwOSAzNC4yMDkxIDY2IDMyIDY2IDI5Ljc5MDkgNjYgMjggNjcuNzkwOSAyOCA3MEwyMiA3MCAyMiA1OCAyOCA1OEMyOCA2MC4yMDkxIDI5Ljc5MDkgNjIgMzIgNjIgMzQuMjA5MSA2MiAzNiA2MC4yMDkxIDM2IDU4TDM4IDU4QzM4IDYwLjIwOTEgMzkuNzkwOSA2MiA0MiA2MiA0NC4yMDkxIDYyIDQ2IDYwLjIwOTEgNDYgNThMNjAgNThDNjAgNjAuMjA5MSA2MS43OTA5IDYyIDY0IDYyIDY2LjIwOTEgNjIgNjggNjAuMjA5MSA2OCA1OEw3NCA1OCA3NCA3MCA2OCA3MEM2OCA2Ny43OTA5IDY2LjIwOTEgNjYgNjQgNjZaTTY4IDU0QzY4IDUxLjc5MDkgNjYuMjA5MSA1MCA2NCA1MCA2MS43OTA5IDUwIDYwIDUxLjc5MDkgNjAgNTRMNDYgNTRDNDYgNTEuNzkwOSA0NC4yMDkxIDUwIDQyIDUwIDM5Ljc5MDkgNTAgMzggNTEuNzkwOSAzOCA1NEwzNiA1NEMzNiA1MS43OTA5IDM0LjIwOTEgNTAgMzIgNTAgMjkuNzkwOSA1MCAyOCA1MS43OTA5IDI4IDU0TDIyIDU0IDIyIDQyIDQwIDQyQzQwIDQ0LjIwOTEgNDEuNzkwOSA0NiA0NCA0NiA0Ni4yMDkxIDQ2IDQ4IDQ0LjIwOTEgNDggNDJMNTAgNDJDNTAgNDQuMjA5MSA1MS43OTA5IDQ2IDU0IDQ2IDU2LjIwOTEgNDYgNTggNDQuMjA5MSA1OCA0Mkw2MCA0MkM2MCA0NC4yMDkxIDYxLjc5MDkgNDYgNjQgNDYgNjYuMjA5MSA0NiA2OCA0NC4yMDkxIDY4IDQyTDc0IDQyIDc0IDU0Wk02OCAzOEM2OCAzNS43OTA5IDY2LjIwOTEgMzQgNjQgMzQgNjEuNzkwOSAzNCA2MCAzNS43OTA5IDYwIDM4TDU4IDM4QzU4IDM1Ljc5MDkgNTYuMjA5MSAzNCA1NCAzNCA1MS43OTA5IDM0IDUwIDM1Ljc5MDkgNTAgMzhMNDggMzhDNDggMzUuNzkwOSA0Ni4yMDkxIDM0IDQ0IDM0IDQxLjc5MDkgMzQgNDAgMzUuNzkwOSA0MCAzOEwyMiAzOCAyMiAyNiAyOCAyNkMyOCAyOC4yMDkxIDI5Ljc5MDkgMzAgMzIgMzAgMzQuMjA5MSAzMCAzNiAyOC4yMDkxIDM2IDI2TDUwIDI2QzUwIDI4LjIwOTEgNTEuNzkwOSAzMCA1NCAzMCA1Ni4yMDkxIDMwIDU4IDI4LjIwOTEgNTggMjZMNjAgMjZDNjAgMjguMjA5MSA2MS43OTA5IDMwIDY0IDMwIDY2LjIwOTEgMzAgNjggMjguMjA5MSA2OCAyNkw3NCAyNiA3NCAzOFpNMjIgODMgMjIgNzQgMjggNzRDMjggNzYuMjA5MSAyOS43OTA5IDc4IDMyIDc4IDM0LjIwOTEgNzggMzYgNzYuMjA5MSAzNiA3NEw1MCA3NEM1MCA3Ni4yMDkxIDUxLjc5MDkgNzggNTQgNzggNTYuMjA5MSA3OCA1OCA3Ni4yMDkxIDU4IDc0TDYwIDc0QzYwIDc2LjIwOTEgNjEuNzkwOSA3OCA2NCA3OCA2Ni4yMDkxIDc4IDY4IDc2LjIwOTEgNjggNzRMNzQgNzQgNzQgODNaIi8+PC9zdmc+" alt="Abacus with solid fill"><img width="96" height="96" id="Graphic 12" src="data:image/svg+xml;base64,PHN2ZyBvdmVyZmxvdz0iaGlkZGVuIiB2aWV3Qm94PSIwLCAwLCA5NiwgOTYiIGlkPSJJY29uc19BYmFjdXNfTSIgeG1sbnM9Imh0dHA6Ly93d3cudzMub3JnLzIwMDAvc3ZnIiB4bWxuczp4bGluaz0iaHR0cDovL3d3dy53My5vcmcvMTk5OS94bGluayIgeG1sbnM6eG1sPSJodHRwOi8vd3d3LnczLm9yZy9YTUwvMTk5OC9uYW1lc3BhY2UiIHZlcnNpb249IjEuMSI+PGcgaWQ9Ikljb25zIj48cGF0aCBkPSJNNzQgOCBMMjIgOCBDMTkuMjM5OSA4LjAwMzMxIDE3LjAwMzMgMTAuMjM5OSAxNyAxMyBMMTcgODMgQzE3LjAwMzMgODUuNzYwMSAxOS4yMzk5IDg3Ljk5NjcgMjIgODggTDc0IDg4IEM3Ni43NjAxIDg3Ljk5NjcgNzguOTk2NyA4NS43NjAxIDc5IDgzIEw3OSAxMyBDNzguOTk2NyAxMC4yMzk5IDc2Ljc2MDEgOC4wMDMzMSA3NCA4IHogTTIyIDEwIEw3NCAxMCBDNzUuNjU2OSAxMCA3NyAxMS4zNDMxIDc3IDEzIEw3NyAyMyBMNjggMjMgTDY4IDIyIEM2OCAxOS43OTA5IDY2LjIwOTEgMTggNjQgMTggQzYxLjc5MDkgMTggNjAgMTkuNzkwOSA2MCAyMiBMNjAgMjMgTDU4IDIzIEw1OCAyMiBDNTggMTkuNzkwOSA1Ni4yMDkxIDE4IDU0IDE4IEM1MS43OTA5IDE4IDUwIDE5Ljc5MDkgNTAgMjIgTDUwIDIzIEwzNiAyMyBMMzYgMjIgQzM2IDE5Ljc5MDkgMzQuMjA5MSAxOCAzMiAxOCBDMjkuNzkwOSAxOCAyOCAxOS43OTA5IDI4IDIyIEwyOCAyMyBMMTkgMjMgTDE5IDEzIEMxOSAxMS4zNDMxIDIwLjM0MzEgMTAgMjIgMTAgeiBNNjQgNjYgQzYxLjc5MDkgNjYgNjAgNjcuNzkwOSA2MCA3MCBMNjAgNzEgTDU4IDcxIEw1OCA3MCBDNTggNjcuNzkwOSA1Ni4yMDkxIDY2IDU0IDY2IEM1MS43OTA5IDY2IDUwIDY3Ljc5MDkgNTAgNzAgTDUwIDcxIEwzNiA3MSBMMzYgNzAgQzM2IDY3Ljc5MDkgMzQuMjA5MSA2NiAzMiA2NiBDMjkuNzkwOSA2NiAyOCA2Ny43OTA5IDI4IDcwIEwyOCA3MSBMMTkgNzEgTDE5IDU3IEwyOCA1NyBMMjggNTggQzI4IDYwLjIwOTEgMjkuNzkwOSA2MiAzMiA2MiBDMzQuMjA5MSA2MiAzNiA2MC4yMDkxIDM2IDU4IEwzNiA1NyBMMzggNTcgTDM4IDU4IEMzOCA2MC4yMDkxIDM5Ljc5MDkgNjIgNDIgNjIgQzQ0LjIwOTEgNjIgNDYgNjAuMjA5MSA0NiA1OCBMNDYgNTcgTDYwIDU3IEw2MCA1OCBDNjAgNjAuMjA5MSA2MS43OTA5IDYyIDY0IDYyIEM2Ni4yMDkxIDYyIDY4IDYwLjIwOTEgNjggNTggTDY4IDU3IEw3NyA1NyBMNzcgNzEgTDY4IDcxIEw2OCA3MCBDNjggNjcuNzkwOSA2Ni4yMDkxIDY2IDY0IDY2IHogTTY2IDcwIEw2NiA3NCBDNjYgNzUuMTA0NiA2NS4xMDQ2IDc2IDY0IDc2IEM2Mi44OTU0IDc2IDYyIDc1LjEwNDYgNjIgNzQgTDYyIDcwIEM2MiA2OC44OTU0IDYyLjg5NTQgNjggNjQgNjggQzY1LjEwNDYgNjggNjYgNjguODk1NCA2NiA3MCB6IE01NiA3MCBMNTYgNzQgQzU2IDc1LjEwNDYgNTUuMTA0NiA3NiA1NCA3NiBDNTIuODk1NCA3NiA1MiA3NS4xMDQ2IDUyIDc0IEw1MiA3MCBDNTIgNjguODk1NCA1Mi44OTU0IDY4IDU0IDY4IEM1NS4xMDQ2IDY4IDU2IDY4Ljg5NTQgNTYgNzAgeiBNMzQgNzAgTDM0IDc0IEMzNCA3NS4xMDQ2IDMzLjEwNDYgNzYgMzIgNzYgQzMwLjg5NTQgNzYgMzAgNzUuMTA0NiAzMCA3NCBMMzAgNzAgQzMwIDY4Ljg5NTQgMzAuODk1NCA2OCAzMiA2OCBDMzMuMTA0NiA2OCAzNCA2OC44OTU0IDM0IDcwIHogTTMwIDU4IEwzMCA1NCBDMzAgNTIuODk1NCAzMC44OTU0IDUyIDMyIDUyIEMzMy4xMDQ2IDUyIDM0IDUyLjg5NTQgMzQgNTQgTDM0IDU4IEMzNCA1OS4xMDQ2IDMzLjEwNDYgNjAgMzIgNjAgQzMwLjg5NTQgNjAgMzAgNTkuMTA0NiAzMCA1OCB6IE00MCA1OCBMNDAgNTQgQzQwIDUyLjg5NTQgNDAuODk1NCA1MiA0MiA1MiBDNDMuMTA0NiA1MiA0NCA1Mi44OTU0IDQ0IDU0IEw0NCA1OCBDNDQgNTkuMTA0NiA0My4xMDQ2IDYwIDQyIDYwIEM0MC44OTU0IDYwIDQwIDU5LjEwNDYgNDAgNTggeiBNNjIgNTggTDYyIDU0IEM2MiA1Mi44OTU0IDYyLjg5NTQgNTIgNjQgNTIgQzY1LjEwNDYgNTIgNjYgNTIuODk1NCA2NiA1NCBMNjYgNTggQzY2IDU5LjEwNDYgNjUuMTA0NiA2MCA2NCA2MCBDNjIuODk1NCA2MCA2MiA1OS4xMDQ2IDYyIDU4IHogTTY4IDU1IEw2OCA1NCBDNjggNTEuNzkwOSA2Ni4yMDkxIDUwIDY0IDUwIEM2MS43OTA5IDUwIDYwIDUxLjc5MDkgNjAgNTQgTDYwIDU1IEw0NiA1NSBMNDYgNTQgQzQ2IDUxLjc5MDkgNDQuMjA5MSA1MCA0MiA1MCBDMzkuNzkwOSA1MCAzOCA1MS43OTA5IDM4IDU0IEwzOCA1NSBMMzYgNTUgTDM2IDU0IEMzNiA1MS43OTA5IDM0LjIwOTEgNTAgMzIgNTAgQzI5Ljc5MDkgNTAgMjggNTEuNzkwOSAyOCA1NCBMMjggNTUgTDE5IDU1IEwxOSA0MSBMNDAgNDEgTDQwIDQyIEM0MCA0NC4yMDkxIDQxLjc5MDkgNDYgNDQgNDYgQzQ2LjIwOTEgNDYgNDggNDQuMjA5MSA0OCA0MiBMNDggNDEgTDUwIDQxIEw1MCA0MiBDNTAgNDQuMjA5MSA1MS43OTA5IDQ2IDU0IDQ2IEM1Ni4yMDkxIDQ2IDU4IDQ0LjIwOTEgNTggNDIgTDU4IDQxIEw2MCA0MSBMNjAgNDIgQzYwIDQ0LjIwOTEgNjEuNzkwOSA0NiA2NCA0NiBDNjYuMjA5MSA0NiA2OCA0NC4yMDkxIDY4IDQyIEw2OCA0MSBMNzcgNDEgTDc3IDU1IHogTTQyIDQyIEw0MiAzOCBDNDIgMzYuODk1NCA0Mi44OTU0IDM2IDQ0IDM2IEM0NS4xMDQ2IDM2IDQ2IDM2Ljg5NTQgNDYgMzggTDQ2IDQyIEM0NiA0My4xMDQ2IDQ1LjEwNDYgNDQgNDQgNDQgQzQyLjg5NTQgNDQgNDIgNDMuMTA0NiA0MiA0MiB6IE01MiA0MiBMNTIgMzggQzUyIDM2Ljg5NTQgNTIuODk1NCAzNiA1NCAzNiBDNTUuMTA0NiAzNiA1NiAzNi44OTU0IDU2IDM4IEw1NiA0MiBDNTYgNDMuMTA0NiA1NS4xMDQ2IDQ0IDU0IDQ0IEM1Mi44OTU0IDQ0IDUyIDQzLjEwNDYgNTIgNDIgeiBNNjIgNDIgTDYyIDM4IEM2MiAzNi44OTU0IDYyLjg5NTQgMzYgNjQgMzYgQzY1LjEwNDYgMzYgNjYgMzYuODk1NCA2NiAzOCBMNjYgNDIgQzY2IDQzLjEwNDYgNjUuMTA0NiA0NCA2NCA0NCBDNjIuODk1NCA0NCA2MiA0My4xMDQ2IDYyIDQyIHogTTY4IDM5IEw2OCAzOCBDNjggMzUuNzkwOSA2Ni4yMDkxIDM0IDY0IDM0IEM2MS43OTA5IDM0IDYwIDM1Ljc5MDkgNjAgMzggTDYwIDM5IEw1OCAzOSBMNTggMzggQzU4IDM1Ljc5MDkgNTYuMjA5MSAzNCA1NCAzNCBDNTEuNzkwOSAzNCA1MCAzNS43OTA5IDUwIDM4IEw1MCAzOSBMNDggMzkgTDQ4IDM4IEM0OCAzNS43OTA5IDQ2LjIwOTEgMzQgNDQgMzQgQzQxLjc5MDkgMzQgNDAgMzUuNzkwOSA0MCAzOCBMNDAgMzkgTDE5IDM5IEwxOSAyNSBMMjggMjUgTDI4IDI2IEMyOCAyOC4yMDkxIDI5Ljc5MDkgMzAgMzIgMzAgQzM0LjIwOTEgMzAgMzYgMjguMjA5MSAzNiAyNiBMMzYgMjUgTDUwIDI1IEw1MCAyNiBDNTAgMjguMjA5MSA1MS43OTA5IDMwIDU0IDMwIEM1Ni4yMDkxIDMwIDU4IDI4LjIwOTEgNTggMjYgTDU4IDI1IEw2MCAyNSBMNjAgMjYgQzYwIDI4LjIwOTEgNjEuNzkwOSAzMCA2NCAzMCBDNjYuMjA5MSAzMCA2OCAyOC4yMDkxIDY4IDI2IEw2OCAyNSBMNzcgMjUgTDc3IDM5IHogTTMwIDI2IEwzMCAyMiBDMzAgMjAuODk1NCAzMC44OTU0IDIwIDMyIDIwIEMzMy4xMDQ2IDIwIDM0IDIwLjg5NTQgMzQgMjIgTDM0IDI2IEMzNCAyNy4xMDQ2IDMzLjEwNDYgMjggMzIgMjggQzMwLjg5NTQgMjggMzAgMjcuMTA0NiAzMCAyNiB6IE01MiAyNiBMNTIgMjIgQzUyIDIwLjg5NTQgNTIuODk1NCAyMCA1NCAyMCBDNTUuMTA0NiAyMCA1NiAyMC44OTU0IDU2IDIyIEw1NiAyNiBDNTYgMjcuMTA0NiA1NS4xMDQ2IDI4IDU0IDI4IEM1Mi44OTU0IDI4IDUyIDI3LjEwNDYgNTIgMjYgeiBNNjIgMjYgTDYyIDIyIEM2MiAyMC44OTU0IDYyLjg5NTQgMjAgNjQgMjAgQzY1LjEwNDYgMjAgNjYgMjAuODk1NCA2NiAyMiBMNjYgMjYgQzY2IDI3LjEwNDYgNjUuMTA0NiAyOCA2NCAyOCBDNjIuODk1NCAyOCA2MiAyNy4xMDQ2IDYyIDI2IHogTTc0IDg2IEwyMiA4NiBDMjAuMzQzMSA4NiAxOSA4NC42NTY5IDE5IDgzIEwxOSA3MyBMMjggNzMgTDI4IDc0IEMyOCA3Ni4yMDkxIDI5Ljc5MDkgNzggMzIgNzggQzM0LjIwOTEgNzggMzYgNzYuMjA5MSAzNiA3NCBMMzYgNzMgTDUwIDczIEw1MCA3NCBDNTAgNzYuMjA5MSA1MS43OTA5IDc4IDU0IDc4IEM1Ni4yMDkxIDc4IDU4IDc2LjIwOTEgNTggNzQgTDU4IDczIEw2MCA3MyBMNjAgNzQgQzYwIDc2LjIwOTEgNjEuNzkwOSA3OCA2NCA3OCBDNjYuMjA5MSA3OCA2OCA3Ni4yMDkxIDY4IDc0IEw2OCA3MyBMNzcgNzMgTDc3IDgzIEM3NyA4NC42NTY5IDc1LjY1NjkgODYgNzQgODYgeiIgLz48L2c+PC9zdmc+" alt="Abacus outline"></p>

<p id="63A27E55-7285-44E4-9151-F5360626A251" class="MsoNormal">&nbsp;</p>

<p id="D05ABC2A-37C7-442F-862B-43CAB3D422C3" class="MsoNormal"><b><u>Shapes Testing</u></b></p>

<p id="6BF130B2-3B72-408E-8176-4E5CECFC0735" class="MsoNormal"><b><u><span style="text-decoration:none">&nbsp;</span></u></b></p>

<p id="83297E36-F86F-4E73-A99C-D708A050B7A5" class="MsoNormal">

</p><table cellpadding="0" cellspacing="0" align="left">
 <tbody><tr>
  <td width="32" height="10"></td>
  <td width="82"></td>
  <td width="106"></td>
  <td width="174"></td>
 </tr>
 <tr>
  <td height="19"></td>
  <td rowspan="3" align="left" valign="top"><img width="82" height="120" src="data:image/png;base64,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"></td>
 </tr>
 <tr>
  <td height="94"></td>
  <td></td>
  <td align="left" valign="top"><img width="174" height="94" src="~WRS%7b5111BE3F-C27E-48D6-AC5F-95677CFFC3E0%7d_files/image010.png"></td>
 </tr>
 <tr>
  <td height="7"></td>
 </tr>
</tbody></table>

<b><u><span style="text-decoration:none">&nbsp;</span></u></b><p></p>

<p id="F449ACEF-360D-4C83-8319-DEFF0B96504E" class="MsoNormal">&nbsp;</p>

<h2 id="9367CF17-8552-42E3-8F0A-6955BECCA880">&nbsp;</h2>

<p id="9D4A6829-883E-45D0-9C7D-E29E473D5074" class="MsoNormal">&nbsp;</p>

<p id="C3FDAE03-63B5-4842-A6F2-3441F9228756" class="MsoNormal">&nbsp;</p>

<br clear="ALL">

<p id="04EEEE8C-00E0-4A45-8A13-C3AF96F8BD16" class="MsoNormal"><b><u>Links testing</u></b></p>

<p id="B2F9A0B4-84D4-46C8-92CF-CD71D0220E8D" class="MsoNormal"><b><u><a href="https://edition.cnn.com/">https://edition.cnn.com/</a><br>
<br>
</u></b><a href="https://www.youtube.com/embed/p-rxzvgUaTE?feature=oembed"><b><span style="color:windowtext"><img border="0" width="480" height="360" id="Video 15" src="~WRS%7b5111BE3F-C27E-48D6-AC5F-95677CFFC3E0%7d_files/image011.jpg" alt="Top 7 Stocks to BUY NOW (High Growth Stocks)"></span></b></a></p>

<p id="FEF64B1F-1F1B-4A4E-AA6C-C95AF94CAEE8" class="MsoNormal">&nbsp;</p>

</div>





`;

const post5Html = `      <h2>Semiconductor Demand Rebounds on AI Infrastructure Buildout</h2>
      <p>Chipmakers report <strong>improving order books</strong> as hyperscalers ramp capital expenditure on data center capacity.</p>
      <p>Foundry utilization rates are back above 90% for advanced nodes, with lead times <em>stretching into 2027</em> for the most in-demand packaging technologies.</p>
      <blockquote>"We're seeing the strongest bookings visibility in three years," one industry executive noted on a recent earnings call.</blockquote>
      <p>Capital spending guidance was raised across the sector for the third consecutive quarter.</p>
`;

export const MOCK_POSTS: FeedPost[] = [
  {
    id: 'post-1',
    title: 'Global Supply Chains Show Signs of Stabilization',
    author: 'Priya Nair',
    publishedAt: '2026-07-28',
    category: 'Trade',
    html: post1Html,
    json: {
  "type": "doc",
  "content": [
    {
      "type": "heading",
      "attrs": {
        "level": 1
      },
      "content": [
        {
          "type": "text",
          "text": "Sample Document",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#C00000"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "This document was created using accessibility techniques for headings, lists, image alternate text, tables, and columns. It should be completely accessible using assistive technologies such as screen readers."
        }
      ]
    },
    {
      "type": "heading",
      "attrs": {
        "level": 2
      },
      "content": [
        {
          "type": "text",
          "text": "Headings"
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "There are eight section headings in this document. At the beginning, \"Sample Document\" is a level 1 heading. The main section headings, such as \"Headings\" and \"Lists\" are level 2 headings. The Tables section contains two sub-headings, \"Simple Table\" and \"Complex Table,\" which are both level 3 headings."
        }
      ]
    },
    {
      "type": "heading",
      "attrs": {
        "level": 2
      },
      "content": [
        {
          "type": "text",
          "text": "Lists"
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "The following outline of the sections of this "
        },
        {
          "type": "text",
          "text": "document is an ordered (numbered"
        },
        {
          "type": "text",
          "text": ") list with six items. The fifth item, \"Tables,\" contains a nested unordered (bulleted) list with two items."
        }
      ]
    },
    {
      "type": "orderedList",
      "content": [
        {
          "type": "listItem",
          "content": [
            {
              "type": "paragraph",
              "content": [
                {
                  "type": "text",
                  "text": "Headings "
                }
              ]
            }
          ]
        },
        {
          "type": "listItem",
          "content": [
            {
              "type": "paragraph",
              "content": [
                {
                  "type": "text",
                  "text": "Lists "
                }
              ]
            }
          ]
        },
        {
          "type": "listItem",
          "content": [
            {
              "type": "paragraph",
              "content": [
                {
                  "type": "text",
                  "text": "Links "
                }
              ]
            }
          ]
        },
        {
          "type": "listItem",
          "content": [
            {
              "type": "paragraph",
              "content": [
                {
                  "type": "text",
                  "text": "Images "
                }
              ]
            }
          ]
        },
        {
          "type": "listItem",
          "content": [
            {
              "type": "paragraph",
              "content": [
                {
                  "type": "text",
                  "text": "Tables "
                }
              ]
            },
            {
              "type": "orderedList",
              "content": [
                {
                  "type": "listItem",
                  "content": [
                    {
                      "type": "paragraph",
                      "content": [
                        {
                          "type": "text",
                          "text": "Simple Tables "
                        }
                      ]
                    }
                  ]
                },
                {
                  "type": "listItem",
                  "content": [
                    {
                      "type": "paragraph",
                      "content": [
                        {
                          "type": "text",
                          "text": "Complex Tables "
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "listItem",
          "content": [
            {
              "type": "paragraph",
              "content": [
                {
                  "type": "text",
                  "text": "Columns "
                }
              ]
            }
          ]
        }
      ]
    },
    {
      "type": "heading",
      "attrs": {
        "level": 2
      },
      "content": [
        {
          "type": "text",
          "text": "Links"
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "In web documents, links can point different locations on the page, different pages, or even downloadable documents, such as Word documents or PDFs:"
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Top of this Page"
        },
        {
          "type": "hardBreak"
        },
        {
          "type": "text",
          "text": "Sample Document",
          "marks": [
            {
              "type": "link",
              "attrs": {
                "href": "http://www.dhs.state.il.us/page.aspx?item=67072"
              }
            }
          ]
        },
        {
          "type": "hardBreak"
        },
        {
          "type": "text",
          "text": "Sample Document (docx)",
          "marks": [
            {
              "type": "link",
              "attrs": {
                "href": "http://www.dhs.state.il.us/OneNetLibrary/27897/documents/Initiatives/IITAA/Sample-Document.docx"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "heading",
      "attrs": {
        "level": 2
      },
      "content": [
        {
          "type": "text",
          "text": "Images"
        }
      ]
    },
    {
      "type": "image",
      "attrs": {
        "src": "data:image/gif;base64,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",
        "alt": "Web Access Symbol",
        "width": 141,
        "height": 131
      }
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Documents may contain images. For example, there is an image of the web accessibility symbol to the left of this paragraph. Its alternate text is \"Web Access Symbol\"."
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Alt text should communicate what an image means, not how it looks."
        }
      ]
    },
    {
      "type": "image",
      "attrs": {
        "src": "data:image/png;base64,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",
        "alt": "Chart of Screen Reader Market Share.\n(Unfortunately, there isn't a way in Word or PDF to include rich formatting, such as a table, in alternate text.)",
        "width": 284,
        "height": 140,
        "align": "float-left"
      }
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "hardBreak"
        },
        {
          "type": "text",
          "text": "Some images, such as charts or graphs, require long descriptions, but not all document types allow that. In web pages, long descriptions may be provided in several ways: on the page below the image, via a link below the image, or via a link on the image."
        }
      ]
    },
    {
      "type": "paragraph"
    },
    {
      "type": "heading",
      "attrs": {
        "level": 2
      },
      "content": [
        {
          "type": "text",
          "text": "Tables"
        }
      ]
    },
    {
      "type": "heading",
      "attrs": {
        "level": 3
      },
      "content": [
        {
          "type": "text",
          "text": "Simple Tables"
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Simple tables have a uniform number of columns and rows, without any merged cells:"
        }
      ]
    },
    {
      "type": "table",
      "content": [
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Screen Reader",
                      "marks": [
                        {
                          "type": "bold"
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Responses",
                      "marks": [
                        {
                          "type": "bold"
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Share ",
                      "marks": [
                        {
                          "type": "bold"
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "JAWS"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "853"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "49% "
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "NVDA"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "238"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "14% "
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Window-Eyes"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "214"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "12% "
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "System Access"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "181"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "10% "
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "VoiceOver"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "159"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "9% "
                    }
                  ]
                }
              ]
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph"
    },
    {
      "type": "heading",
      "attrs": {
        "level": 3
      },
      "content": [
        {
          "type": "text",
          "text": "Complex Tables"
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "The following is a complex table, using merged cells as headers for sections within the table. This can't be made accessible in all types of documents:"
        }
      ]
    },
    {
      "type": "table",
      "content": [
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": []
                }
              ]
            },
            {
              "type": "tableCell",
              "attrs": {
                "colspan": 2
              },
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "May 2012",
                      "marks": [
                        {
                          "type": "bold"
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "attrs": {
                "colspan": 2
              },
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": " September 2010",
                      "marks": [
                        {
                          "type": "bold"
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Screen Reader",
                      "marks": [
                        {
                          "type": "bold"
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Responses",
                      "marks": [
                        {
                          "type": "bold"
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Share",
                      "marks": [
                        {
                          "type": "bold"
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Responses",
                      "marks": [
                        {
                          "type": "bold"
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Share ",
                      "marks": [
                        {
                          "type": "bold"
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "JAWS"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "853"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "49%"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "727"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "59% "
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "NVDA"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "238"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "14%"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "105"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "9% "
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Window-Eyes"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "214"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "12%"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "138"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "11% "
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "System Access"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "181"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "10%"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "58"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "5% "
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "VoiceOver"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "159"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "9%"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "120"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "10%"
                    }
                  ]
                }
              ]
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "heading",
      "attrs": {
        "level": 2
      },
      "content": [
        {
          "type": "text",
          "text": "Columns"
        }
      ]
    },
    {
      "type": "columnLayout",
      "attrs": {
        "columns": 2
      },
      "content": [
        {
          "type": "paragraph",
          "content": [
            {
              "type": "text",
              "text": "This is an example of columns. With columns, the page is split into two or more horizontal sections. Unlike tables, in which you usually read across a row and then down to the next, in columns, you read down a column and then across to the next."
            }
          ]
        },
        {
          "type": "paragraph",
          "content": [
            {
              "type": "text",
              "text": "When columns are not created correctly, screen readers may run lines together, reading the first line of the first column, then the first line of the second column, then the second line of the first column, and so on. Obviously, that is not accessible."
            }
          ]
        },
        {
          "type": "paragraph"
        }
      ]
    }
  ]
},
  },
  {
    id: 'post-2',
    title: 'Renewable Energy Investment Hits Record High',
    author: 'Daniel Osei',
    publishedAt: '2026-07-25',
    category: 'Energy',
    html: post2Html,
    json: {
  "type": "doc",
  "content": [
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Lorem ipsum "
        }
      ]
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      }
    },
    {
      "type": "bulletList",
      "content": [
        {
          "type": "listItem",
          "content": [
            {
              "type": "heading",
              "attrs": {
                "level": 1
              },
              "content": [
                {
                  "type": "text",
                  "text": "Lorem ipsum dolor sit amet, consectetur adipiscing elit. Nunc ac faucibus odio. "
                }
              ]
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      }
    },
    {
      "type": "image",
      "attrs": {
        "src": "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAsIAAAEQCAYAAAC6DBJgAAAAAXNSR0IArs4c6QAAAARnQU1BAACxjwv8YQUAAAAJcEhZcwAADsMAAA7DAcdvqGQAAABhaVRYdFNuaXBNZXRhZGF0YQAAAAAAeyJjbGlwUG9pbnRzIjpbeyJ4IjowLCJ5IjowfSx7IngiOjcwNiwieSI6MH0seyJ4Ijo3MDYsInkiOjI3Mn0seyJ4IjowLCJ5IjoyNzJ9XX3DINdGAAAZJUlEQVR4Xu3dXYhd1f038DXPvcS8XPVC2jQWApYR0kykjQUtmkGmFyUVUwwloBgc+xesMOk4kV4YzJMpIk1bY6RCLvLoiM1NO8ZJpAbaaamTRAwtBPJW8SJXjjF4/Wee/Jb7zJyMGaPJPjN7Zn0+cDhr7X3eJlms+c46v71319QVCQAACvN/qnsAACiKIAwAQJEEYQAAiiQIAwBQJEEYAIAiCcIAABRJEAYAoEiCMAAARRKEAQAokiAMAECRBGEAAIokCAMAUCRBGACAIgnCAAAUSRAGAKBIgjAAAEUShAEAKNJNBeH+/v7U29tb9T43OTmZt3V1daUVK1akkZGRag8AADTHDQXh8fHxtGbNmnT8+PFqy4wnnngirVu3Lk1NTaVjx47lsHzhwoVqLwAANMMNrwgfOHAgvfjii1VvxhtvvJEeeeSR3O7u7k49PT3p4sWLuQ8AAE1xQ0F448aN+XYtsRK8evXqqpfSxMREuuWWW6oeAAA0Q9eV4DpVtb+2KJHYtWtXGhsbq7ZcrVUWMdf+dmfOnMm3lvhYUWcMANAJfX19VYtSdSwIDw0NpaNHj+Z9K1eurLZ+daOjowYoANARcgahI6dPi5XgmwnBAADQabUH4QjBcTYJIRgAgCarNQhHqcS+ffvSiRMn0qpVq3KNb9yGh4erRwAAQDPcVI1wJ6ndAQA6Rc4gdKRGGAAAmk4QBgCgSIIwAABFEoQBACiSIAwAQJEEYQAAiiQIAwBQJEEYAIAiCcIAABRJEAYAoEiCMAAARRKEAQAokiAMAECRBGEAAIokCAMAUCRBGACAIgnCAAAUSRAGAKBIgjAAAEUShAEAKJIgDABAkQRhAACKJAgDAFAkQRgAgCIJwgAAFOmmgnB/f3/q7e2tep+bnJzM27q6utKaNWvS4cOHqz0AANAcNxSEx8fHc8g9fvx4tWXGs88+m2699dY0NTWVDh06lLZu3ZrDMQAANMkNrwgfOHAgvfjii1VvxsjISBocHMzt7u7u1NPTk955553cBwCAprihILxx48Z8u5ZLly7lANyybt269NFHH1U9AABohq6pqGG4QVEisWvXrjQ2Njbdv/vuu3NZRMvw8HC+HxgYyPdzOXPmTL4BAMyHvr6+qkWpag3CIQ6Sa3/JoaGhtGzZsusG4dlGR0cNUACgI+QMQu2nT1u+fHk6depU1Uvp5MmT6bbbbqt6AADQDLUH4S1btqTdu3fndgTiiYmJdN999+U+AAA0Re1B+LnnnkuffvppLpHYvHlzOnjwYFq5cmW1FwAAmuGmaoQ7Se0OANApcgah9hVhAABYDKwIQ+n+862qscDu+G/VAOg8OYNgRRgAgCIJwgAAFEkQBgCgSIIwAABFEoQBACiSIAwAQJEEYQAAiiQIAwBQpCV/QY2u+3dXrYUzdXSwakEDuaAGUCAX1CBYEQYAoEhWhOeBFWEazYowLB4/WlY1FtBfL1eNxc2KMMGKMAAARRKEAQAokiAMAECR1AjPAzXCzKkJ9X6/XVE1FpgaYbg+NcK1USNMsCIMAECRBGEAAIokCAMAUCRBGACAIgnCAAAUSRAGAKBIgjAAAEUShAEAKFLtQXhycjL19vamrq6utGbNmjQyMlLtAQBgPkUOizwWuWzFihWpv78/Z7WvKp630OLzrl+/Pg0PD1db6lN7EH722WfTvffem+KCdYcOHUo/+9nPvtY/OAAA1xZXzJ3rNtv+/fvTzp0704EDB3IuO3v2bN4eC5aLRfwMt99+e9WrX+1B+MKFC+n73/9+bnd3d6dNmzal06dP5z4AAPNjcHAwh+CNGzfm/sqVK9NLL72ULl26lMbHx/O2plu2bFk6ceJEevDBB6st9ao9CMdq8GuvvZbbp06dShMTE2nt2rW5DwBA50XQjcDbCsHtzp07N7398OHDuXQiSg+uV9I6u0wiVpbbA3Xsj+dHCUa0o5Rhdr9l9mOHhoaqPVfbsmVLWr16ddWrX9dUrJXXKFaE77///nT+/Pncf/nll9P27dtz+8ucOXMm3+r2473/rloL5y9PfrdqwdX6Xny4ai2g366oGgtr9MPfVS1gLk2YM0af+n9Va/Hr6+urWovHtUogWqaODlatz4Pw3XffnUsi5hKlq1F2cOzYsfwt/ux+iJDaeo32doggHKUXrVAd+1u5L15r1apVX+i3v9bsfZEd5wq9rRA9MDCQ7+tSexCOvygeffTR/IO1QnH7svxXNTo6WssA/bIBM1/aByZc5UfLqsYCakgQTnf8t2oAc2rCnPHXy1VjcasrZ8y3OoNwPGbXrl1pbGys2pLywXTf/OY3pwPn1w3C7fuj/8EHH1wVqlv9673WbJ0KwrWXRkQdR2sFOFL9Y489lv75z3/mPgAAnReBcvny5VeVLrRECcS1tnfCZ599VrU+N7u/0GoPwvGP3qoviaXuN998M9122225DwDA/Ni9e3d66qmnpkNv5LJY8Y2sFkE5juGKY7nimK7W/shwcaKDdrG9pfXYeM147mJXexCOupIXXnghL3lHnUkc5ReFzgAAzJ/4hv7pp59O27Ztm85loVUKEWeROHjwYNq8eXMubd2wYUM+q0SrlCHs2bMnPy8C8Ouvv57uvPPO/Fpvv/126unpqR61eNVeI1wXNcIUQY3wDDXCcH1qhGuzWGuEqVftK8IAALAYCMIAABRJEAYAoEiCMAAARRKEAQAokiAMAECRBGEAAIokCAMAUCRBGACAIgnCAAAUSRAGAKBIgjAAAEUShAEAKJIgDABAkQRhAIAlamRkJK1ZsyZ1dXWlFStWpP7+/jQ5OVntvb543kI6derU9OeP++jXSRAGAFgsfrRs7tss+/fvTzt37kwHDhxIU1NT6ezZs3l7b29vvl8M7rnnnrR37978+Xft2pU2b95c7amHIAwAsAQNDg7mELxx48bcX7lyZXrppZfSpUuX0vj4eN7WZPEZYxX7gQceyP0tW7ak8+fP53ZdBGEAgCUmQmQE3lYIbnfu3Lnp7YcPH84lB+vXr8/3UUoxl9llErGy3B6oY388P8JrtIeHh7/Qb5n92KGhoWrPjPiM8Vlboixi+fLlVa8egjAAQIGiVnjr1q3p0KFD6fjx4+m9997LNcQ3U4d7+fLl9Mknn6SPP/447dix4wv9du37nn/++XThwoVqzxfFZ3300UfTr371q2pLPQRhAIACnT59OvX09KTu7u7cj9KJKD84cuRI7t+I7du35/t4rXDXXXfl+1a/PWS3P3bTpk3p4sWLuT9bhOCHH344r1oPDAxUW+shCAMALDFRVhBlBNeqBY4SiPmqEf7ss8+q1udm968nQnCUYKxevTrXN9dNEAYAWIJ2796dnnrqqenQG6EySh8iIEdQXrt2bZqYmJhepY39Ubcbq7PtYntL67HxmvHcTmqF4FgJ7kQIDoIwAMASFKUHTz/9dNq2bVs+IO3222/P28fGxvJ9lCQcPHgwn5IswuaGDRty4GyVSoQ9e/bk50UAfv3119Odd96ZX+vtt9/OZRWd9Oqrr6YTJ06kffv25fds3epcze6aihOzNdDo6Gjq6+urejeu6/7dVWvhTB0drFowyzXO+zjvfruiaiywO/5bNYA5NWHO+OvlqrG41ZUzWNw6siIcp8CIxB6nxLjW6TAAAGCh1R6E4xxxcbLj1hVMTp48+aXnpAMAgIVQe2lErAIfO3bsqvqSG6E0giIojZihNAKuz5wx4ybnDKURhNpXhOMqJv/617++9EohAACw0GoNwq2j+NqvFHL06FGlEQAANE6tpRFxvrdVq1bl+uCWqBl+//33rxuGz5w5k291+/Hef1ethfOXJ79bteBqfS8+XLUWUEO+5hz98HdVC5iLOWNGHXOG0gg6UiMc53yLK4CECMIffvjh1z4RshphiqDeb4YaYbg+c8YMNcLUoPYa4ccffzw988wzuR0rxK+88kr64Q9/mPsAANAUtQfhX/7yl9MHysWVSB577LG0ZcuWai8AADRD7UE4LtcXZRBRcREHzA0MDFR7AACgOWoPwgAAsBgIwgAAFEkQBgCgSIIwAABFEoQBAJaouKDZmjVr8tm84qxe/f39+fS2X1U8byGdOnUqrV+/Pn+O+DmiXydBGABgCdq/f3/auXNnOnDgQD6b19mzZ/P23t7efL8YbN68OT399NP588fPcc8991R76iEIAwAsFv/51ty3WQYHB3N43LhxY+63TnF76dKlND4+nrc12YULF9L58+enr0cRP0d89joJwgAAS0wE3QiNrRDc7ty5c9PbDx8+nEsOovwg7qOUYi6zyyRiZbk9UMf+eH7rwmrDw8Nf6LfMfuzQ0FC1Z8bq1avzSnBLPH7Tpk1Vrx6CMABAgaJWeOvWrenQoUPp+PHj6b333ss1xDdTh3v58uV8QbWPP/447dix4wv9du37nn/++bwCfC0RtiMsx2fbs2dPtbUegjAAQIFOnz6denp6Und3d+5H6USUIRw5ciT3b8T27dvzfbxWuOuuu/J9q98estsfGyu9Fy9ezP3ZYvU6Vob//Oc/5xrhuQLzjRCEAQCWmAiPy5cvv2YtcJRAzFeN8GeffVa1Pje7/2XiM7Yf2Bc/UwT3uQLzjRCEAQCWoN27d6ennnpqOvRGKUSUF0RAjlC5du3aNDExMb1KG/uvVYfbfrq11mPjNeO5nfSNb3wjr0633jPu4z1je10EYQCAJShKD+LUY9u2bcs1trfffnvePjY2lu+jJOHgwYP5FGVxsNyGDRvyWSVapRIhanLjeRFCX3/99XTnnXfm13r77bfz6mwnxcFyb731Vv588Z5xH58vttela6r9cLwGGR0dTX19fVXvxnXdv7tqLZypo4NVC2b50bKqsYB+u6JqLLA7/ls1gDmZM2bc5JxRV85gcbMiDABAkQRhAACKJAgDAFAkQRgAgCIJwgAAFEkQBgCgSIIwAABFch7hedCY8wj/51tVYwE5V+zVnBN0hrEB12fOmOE8wtTAijAAAEUShAEAKJIgDABAkTpaI9zb25vvx8bG8v3XoUa4A9QIN496vxnGBg3WhN8lYep//2/VWkBqhFlCOrYivH///nTkyJGqBwAAzdKRIHzhwoX0m9/8Jj3zzDPVFgAAaJaOBOH+/v60d+/etGxZA772BQCAa6i9RjhKIo4dO5ZGRkbS8PBwevfdd79SjfCZM2fyrW4/3vvvqrVw/vLkd6vWwur75v9UrYUz+uHvqhah78WHq9YCaki9n7FBkzXhd0lQIzyjjjlDjTC1BuHJycm0YcOG9N5776WVK1d+rSA8m4PlOsDBcs3jYLkZxgYN5mC5Ng6WYwmptTTi1VdfTefPn0+rVq1KXV1daceOHfmAudbZIwAAoClqDcIDAwMpFphbtz179qRNmzbd0IowAAB0UsdOnwYAAE3W0SAcK8RWgwEAaKLazxpRFwfLdYCD5ZrHwXIzjI3mMWdMc7BcGwfLsYQojQAAoEiCMAAARRKEAQAokiAMAECRBGEAAIokCAMAUCRBGACAIgnCAAAUyQU15kEjToAemnASdBdNuJoLaswwNprHBTWmuaBGGxfUYAmxIgwAQJGsCM8DK8JtrPpdzYrwjAaMjcas+rks+wwrwlexItzGijA1sCIMAECRBGEAAIokCAMAUCRBGACAIgnCAAAUSRAGAKBIgjAAAEUShAEAKJIgDABAkQRhAACKJAgDAFCkrqkrqnaj1HUN8CZcH74R14YPTbg+/E1eG74uTRgXoRFjownjIjRgbDRmXBwdrFoL7D/fqhoLyJxxFXNGm5scG3XlDBY3K8IAABSp9iA8OTmZent7U1dXV1qxYkUaGRmp9gAAQHPUHoSfeOKJtG7duhQVF8eOHUv9/f3pwoUL1V4AAGiG2oPwG2+8kR555JHc7u7uTj09PenixYu5DwAATVF7EI6V4NWrV1e9lCYmJtItt9xS9QAAoBk6etaIVlnE2NhYtWVuZ86cybe6/Xjvv6vWwnHWiBmjH/6uai2sJoyL4AjwGU0YG00ZF3958rtVa2H1ffN/qtbCMWdczZwxo46x4awRdCwIDw0NpaNHj+YQvHLlymrrV+f0aR3QhMnLqZCu4pdaG6dPm+b0aW3MGVcxZ7Rx+jRq0JHTp8VK8M2EYAAA6LTag3CE4OPHjwvBAAA0Wq1BeHx8PO3bty+dOHEirVq1Kp9LOG7Dw8PVIwAAoBlqDcIbN27MZ42YfRsYGKgeAQAAzdCRGmEAAGg6QRgAgCIJwgAAFEkQBmiaHy1rxg1giROEAQAokiAMAECRBGEAAIokCAMAUCRBGACAIgnCAAAUSRAGAKBIgjAAAEUShAEAKJIgDABAkQRhAACKJAgDAFAkQRgAgCIJwgAAFEkQBgCgSIIwAABFEoQBACiSIAwAQJEEYQAAiiQIAwBQJEEYAIAi1R6EJycnU29vb+rq6kpr1qxJhw8frvYAAEBz1B6En3322XTrrbemqampdOjQobR169YcjgEAoElqD8IjIyNpcHAwt7u7u1NPT0965513ch8AAJqi9iB86dKlHIBb1q1blz766KOqBwAAzdA1FTUMNRkfH0933313LotoGR4ezvcDAwP5fi5nzpzJt5Z4jagzvln1/XQ3ris14EOEm//nrEEjPkQjxkVoxNhoxn/JFQv/QYyLWcwZ04yNNktozujr66talKrWIBwivLa/5NDQUFq2bNl1g/Bso6OjBijXZGxwLcYFczE2uBbjglB7acTy5cvTqVOnql5KJ0+eTLfddlvVAwCAZqg9CG/ZsiXt3r07tyMQT0xMpPvuuy/3AQCgKWoPws8991z69NNPc4nE5s2b08GDB9PKlSurvV/dd77znaoFVzM2uBbjgrkYG1yLcUGovUYYAAAWg9pXhAEAYDEQhAEAKJIgDABAkQRhAACKJAgDAFAkQRgAgCIJwgAAFEkQngfDw8P5AiOt24oVK9LQ0FC1t3NGRkby+9Fc7eMibuvXr7/qEuV1i9des2ZNfq+47+R7ceMWYs6I14/3ifebj/mJG9M+LuLW6TmjZXJyMo+PGJuwlAjC82TTpk0prl0St7Nnz6Y33ngjB9VOuHDhQurt7U0vvPBCtYUm+/vf/z49Nh599NF8RcZOueeee9LevXvze+3atauj78XNmc85I1736NGj+X3i/U6ePJn2799f7aVp5nPOaHniiSfSpUuXqh4sHYLwAohLTt9///3po48+qrbUb9u2ben48eNVj8XivvvuS+fPn6969RofH88rOg888EDub9mypWPvRb06PWf87W9/Sw8++OD05fDvvffedPny5dym2To5Z7QcPnw4ffrpp/mPM1hqBOEFECu2sfry/e9/P/fjK6f+/v4cUuLr6ggosS0mn/jaq6X11WVLrOLE82ZbvXp1fg0Wnz/96U9X/bKJMRBjIsZB3LdWBKMf+0KMlfiKtH0FLx4b46zdxo0b07lz56re52USy5cvr3o0WafnjJdeeikNDAzkdrzOm2++Of1eNFsn54wQj33yySfzGIElaYqO27Nnz1T8U7ffXn755Wrv5/sfeuihqnd1/0pQmbry135uf+9735u6MuFNvfXWW7kfj2m15+K/uNlmj4tvf/vbUx9//HHeF/fx///BBx98oR/j5/HHH8/bYwzEuIjxEWJ/qz2XeK14TIw1mmeh5ox4bLxXa2zRPLPHRafnjBgzrbEXzzFnsNRYEZ4nVyaQ6ZquK5NOGhwczH9ph3fffTf94he/yO3w05/+NNcDhljpeeedd/Jf6vG1ZZQ8/OMf/8j7YoWo9TU3i1d7vd+VX0bp1VdfzdtPnz6denp6Und3d+7H/3+MhyNHjuSvQ1srPaOjo2nnzp25fi/GVOyPusG5xGMefvjhvELUWgWkeRZizhgbG0tXwlP65JNPHDDXYPM1Z0Q5VZREbN++vdoCS48gvABikrryV3z+ZXU9fX196dixY/mxP/nJT/JkFr/w4iuumOBYWn7+85+nV155perNLcpfYgxFeUOEmyh7eOihh/I4ia+1Y5xcS/zSiwMp4/m+6lw8Oj1nxJiI0BMiPEXIjgPmaL5OzhlxQG2E5NYZKqK9Y8cOZ45gSRGEF0gcmBIHqIQ4MOX3v/99boeo+YoJKsTqTUxaf/zjH/NEFb+kli9fnn7961/nX3gsLfH/HQe+xGre2rVr08TExPSpkSLExopOrBSGWMGJWxxEFX7wgx9MnykkfunN1grBsRIsBC8+nZwzYry0v95rr712zTFE83RyzohvCForz3GL19mzZ49vklhSBOEFcscdd+RfVuGRRx7JB7S0Dnx5//330x/+8Ie8L7RWcVoTVUxkMfF92VecLF7xVWf8MosAc/DgwXxqpAivGzZsyAG29bVnhJwTJ05Mh5vWL8QYH9cSX5/G4/ft2ze9whO31kogzdbJOeO5557Lr9UaE1EaEdtYHDo1Z0AJuq78lRcF9wAAUBQrwgAAFEkQBgCgSIIwAABFEoQBACiSIAwAQJEEYQAAiiQIAwBQJEEYAIAiCcIAABRJEAYAoEiCMAAARRKEAQAokiAMAECRBGEAAIokCAMAUCRBGACAIgnCAAAUSRAGAKBAKf1/pgJqIWyaDYEAAAAASUVORK5CYII=",
        "alt": "Picture 1",
        "width": 706,
        "height": 272,
        "align": "float-right"
      }
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      },
      "content": [
        {
          "type": "text",
          "text": "Vestibulum neque massa, scelerisque sit amet ligula eu, congue molestie mi. Praesent ut varius sem. Nullam at porttitor arcu, nec lacinia nisi. Ut ac dolor vitae odio interdum condimentum. ",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#000000"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": "Vivamus dapibus sodales ex, vitae malesuada ipsum cursus convallis. Maecenas sed egestas nulla, ac condimentum orci. ",
          "marks": [
            {
              "type": "bold"
            },
            {
              "type": "textStyle",
              "attrs": {
                "color": "#000000"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": "Mauris diam felis, vulputate ac suscipit et, iaculis non est. Curabitur semper arcu ac ligula semper, nec luctus nisl blandit. Integer lacinia ante ac libero lobortis imperdiet. ",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#000000"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": "Nullam mollis convallis ipsum, ac accumsan nunc vehicula vitae. ",
          "marks": [
            {
              "type": "italic"
            },
            {
              "type": "textStyle",
              "attrs": {
                "color": "#000000"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": "Nulla eget justo in felis tristique fringilla. Morbi sit amet tortor quis risus auctor condimentum. Morbi in ullamcorper elit. Nulla iaculis tellus sit amet mauris tempus fringilla.",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#000000"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      },
      "content": [
        {
          "type": "text",
          "text": "Maecenas mauris lectus, lobortis et purus mattis, blandit dictum tellus.",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#000000"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "bulletList",
      "content": [
        {
          "type": "listItem",
          "content": [
            {
              "type": "paragraph",
              "attrs": {
                "textAlign": "justify"
              },
              "content": [
                {
                  "type": "text",
                  "text": "Maecenas non lorem quis tellus placerat varius. ",
                  "marks": [
                    {
                      "type": "bold"
                    },
                    {
                      "type": "textStyle",
                      "attrs": {
                        "color": "#000000"
                      }
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "listItem",
          "content": [
            {
              "type": "paragraph",
              "attrs": {
                "textAlign": "justify"
              },
              "content": [
                {
                  "type": "text",
                  "text": "Aenean congue fringilla justo ut aliquam.   ",
                  "marks": [
                    {
                      "type": "underline"
                    },
                    {
                      "type": "textStyle",
                      "attrs": {
                        "color": "#000000"
                      }
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "listItem",
          "content": [
            {
              "type": "paragraph",
              "attrs": {
                "textAlign": "justify"
              },
              "content": [
                {
                  "type": "text",
                  "text": "Mauris id ex erat. ",
                  "marks": [
                    {
                      "type": "textStyle",
                      "attrs": {
                        "color": "#000000"
                      }
                    },
                    {
                      "type": "link",
                      "attrs": {
                        "href": "https://products.office.com/en-us/word"
                      }
                    }
                  ]
                },
                {
                  "type": "text",
                  "text": "Nunc vulputate neque vitae justo facilisis, non condimentum ante sagittis. ",
                  "marks": [
                    {
                      "type": "textStyle",
                      "attrs": {
                        "color": "#000000"
                      }
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "listItem",
          "content": [
            {
              "type": "paragraph",
              "attrs": {
                "textAlign": "justify"
              },
              "content": [
                {
                  "type": "text",
                  "text": "Morbi viverra semper lorem nec molestie. ",
                  "marks": [
                    {
                      "type": "textStyle",
                      "attrs": {
                        "color": "#000000"
                      }
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "listItem",
          "content": [
            {
              "type": "paragraph",
              "attrs": {
                "textAlign": "justify"
              },
              "content": [
                {
                  "type": "text",
                  "text": "Maecenas tincidunt est efficitur ligula euismod, sit amet ornare est vulputate.",
                  "marks": [
                    {
                      "type": "textStyle",
                      "attrs": {
                        "color": "#000000"
                      }
                    }
                  ]
                }
              ]
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      }
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      }
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      }
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      }
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      }
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      }
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      }
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      }
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      }
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      },
      "content": [
        {
          "type": "text",
          "text": "In non mauris justo. Duis vehicula mi vel mi pretium, a viverra erat efficitur. Cras aliquam est ac eros varius, id iaculis dui auctor. Duis pretium neque ligula, et pulvinar mi placerat et. Nulla nec nunc sit amet nunc posuere vestibulum. Ut id neque eget tortor mattis tristique. Donec ante est, blandit sit amet tristique vel, lacinia pulvinar arcu. Pellentesque scelerisque fermentum erat, id posuere justo pulvinar ut. Cras id eros sed enim aliquam lobortis. Sed lobortis nisl ut eros efficitur tincidunt. Cras justo mi, porttitor quis mattis vel, ultricies ut purus. Ut facilisis et lacus eu cursus.",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#000000"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      },
      "content": [
        {
          "type": "text",
          "text": "In eleifend velit vitae libero sollicitudin euismod. Fusce vitae vestibulum velit. Pellentesque vulputate lectus quis pellentesque commodo. Aliquam erat volutpat. Vestibulum in egestas velit. Pellentesque fermentum nisl vitae fringilla venenatis. Etiam id mauris vitae orci maximus ultricies. ",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#000000"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      }
    },
    {
      "type": "bulletList",
      "content": [
        {
          "type": "listItem",
          "content": [
            {
              "type": "heading",
              "attrs": {
                "level": 1
              },
              "content": [
                {
                  "type": "text",
                  "text": "Cras fringilla ipsum magna, in fringilla dui commodo a."
                }
              ]
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      }
    },
    {
      "type": "table",
      "content": [
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": []
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Lorem ipsum",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Lorem ipsum"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Lorem ipsum"
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "1"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "In eleifend velit vitae libero sollicitudin euismod.",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Lorem"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": []
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "2"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Cras fringilla ipsum magna, in fringilla dui commodo a.",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Ipsum"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": []
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "3"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Aliquam erat volutpat. ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Lorem"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": []
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "4"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Fusce vitae vestibulum velit. ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Lorem"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": []
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "5"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Etiam vehicula luctus fermentum.",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Ipsum"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": []
                }
              ]
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      }
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      },
      "content": [
        {
          "type": "text",
          "text": "Etiam vehicula luctus fermentum. In vel metus congue, pulvinar lectus vel, fermentum dui. Maecenas ante orci, egestas ut aliquet sit amet, sagittis a magna. Aliquam ante quam, pellentesque ut dignissim quis, laoreet eget est. Aliquam erat volutpat. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos. Ut ullamcorper justo sapien, in cursus libero viverra eget. Vivamus auctor imperdiet urna, at pulvinar leo posuere laoreet. Suspendisse neque nisl, fringilla at iaculis scelerisque, ornare vel dolor. Ut et pulvinar nunc. Pellentesque fringilla mollis efficitur. Nullam venenatis commodo imperdiet. Morbi velit neque, semper quis lorem quis, efficitur dignissim ipsum. Ut ac lorem sed turpis imperdiet eleifend sit amet id sapien.",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#000000"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      }
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      }
    },
    {
      "type": "bulletList",
      "content": [
        {
          "type": "listItem",
          "content": [
            {
              "type": "heading",
              "attrs": {
                "level": 1
              },
              "content": [
                {
                  "type": "text",
                  "text": "Lorem ipsum dolor sit amet, consectetur adipiscing elit. "
                }
              ]
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      }
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      },
      "content": [
        {
          "type": "text",
          "text": "Nunc ac faucibus odio. Vestibulum neque massa, scelerisque sit amet ligula eu, congue molestie mi. Praesent ut varius sem. Nullam at porttitor arcu, nec lacinia nisi. Ut ac dolor vitae odio interdum condimentum. Vivamus dapibus sodales ex, vitae malesuada ipsum cursus convallis. Maecenas sed egestas nulla, ac condimentum orci. Mauris diam felis, vulputate ac suscipit et, iaculis non est. Curabitur semper arcu ac ligula semper, nec luctus nisl blandit. Integer lacinia ante ac libero lobortis imperdiet. Nullam mollis convallis ipsum, ac accumsan nunc vehicula vitae. Nulla eget justo in felis tristique fringilla. Morbi sit amet tortor quis risus auctor condimentum. Morbi in ullamcorper elit. Nulla iaculis tellus sit amet mauris tempus fringilla.",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#000000"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "bulletList",
      "content": [
        {
          "type": "listItem",
          "content": [
            {
              "type": "heading",
              "attrs": {
                "level": 2
              },
              "content": [
                {
                  "type": "text",
                  "text": "Maecenas mauris lectus, lobortis et purus mattis, blandit dictum tellus. "
                }
              ]
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      },
      "content": [
        {
          "type": "text",
          "text": "Maecenas non lorem quis tellus placerat varius. Nulla facilisi. Aenean congue fringilla justo ut aliquam. Mauris id ex erat. Nunc vulputate neque vitae justo facilisis, non condimentum ante sagittis. Morbi viverra semper lorem nec molestie. Maecenas tincidunt est efficitur ligula euismod, sit amet ornare est vulputate.",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#000000"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      },
      "content": [
        {
          "type": "text",
          "text": "In non mauris justo. Duis vehicula mi vel mi pretium, a viverra erat efficitur. Cras aliquam est ac eros varius, id iaculis dui auctor. Duis pretium neque ligula, et pulvinar mi placerat et. Nulla nec nunc sit amet nunc posuere vestibulum. Ut id neque eget tortor mattis tristique. Donec ante est, blandit sit amet tristique vel, lacinia pulvinar arcu. Pellentesque scelerisque fermentum erat, id posuere justo pulvinar ut. Cras id eros sed enim aliquam lobortis. Sed lobortis nisl ut eros efficitur tincidunt. Cras justo mi, porttitor quis mattis vel, ultricies ut purus. Ut facilisis et lacus eu cursus.",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#000000"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "bulletList",
      "content": [
        {
          "type": "listItem",
          "content": [
            {
              "type": "heading",
              "attrs": {
                "level": 2
              },
              "content": [
                {
                  "type": "text",
                  "text": "In eleifend velit vitae libero sollicitudin euismod. "
                }
              ]
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      },
      "content": [
        {
          "type": "text",
          "text": "Fusce vitae vestibulum velit. Pellentesque vulputate lectus quis pellentesque commodo. Aliquam erat volutpat. Vestibulum in egestas velit. Pellentesque fermentum nisl vitae fringilla venenatis. Etiam id mauris vitae orci maximus ultricies. Cras fringilla ipsum magna, in fringilla dui commodo a.",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#000000"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      },
      "content": [
        {
          "type": "text",
          "text": "Etiam vehicula luctus fermentum. In vel metus congue, pulvinar lectus vel, fermentum dui. Maecenas ante orci, egestas ut aliquet sit amet, sagittis a magna. Aliquam ante quam, pellentesque ut dignissim quis, laoreet eget est. Aliquam erat volutpat. Class aptent taciti sociosqu ad litora torquent per conubia nostra, per inceptos himenaeos. Ut ullamcorper justo sapien, in cursus libero viverra eget. Vivamus auctor imperdiet urna, at pulvinar leo posuere laoreet. Suspendisse neque nisl, fringilla at iaculis scelerisque, ornare vel dolor. Ut et pulvinar nunc. Pellentesque fringilla mollis efficitur. Nullam venenatis commodo imperdiet. Morbi velit neque, semper quis lorem quis, efficitur dignissim ipsum. Ut ac lorem sed turpis imperdiet eleifend sit amet id sapien.",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#000000"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "justify"
      }
    },
    {
      "type": "image",
      "attrs": {
        "src": "data:image/jpeg;base64,/9j/4AAQSkZJRgABAQEASABIAAD/4gxYSUNDX1BST0ZJTEUAAQEAAAxITGlubwIQAABtbnRyUkdCIFhZWiAHzgACAAkABgAxAABhY3NwTVNGVAAAAABJRUMgc1JHQgAAAAAAAAAAAAAAAAAA9tYAAQAAAADTLUhQICAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABFjcHJ0AAABUAAAADNkZXNjAAABhAAAAGx3dHB0AAAB8AAAABRia3B0AAACBAAAABRyWFlaAAACGAAAABRnWFlaAAACLAAAABRiWFlaAAACQAAAABRkbW5kAAACVAAAAHBkbWRkAAACxAAAAIh2dWVkAAADTAAAAIZ2aWV3AAAD1AAAACRsdW1pAAAD+AAAABRtZWFzAAAEDAAAACR0ZWNoAAAEMAAAAAxyVFJDAAAEPAAACAxnVFJDAAAEPAAACAxiVFJDAAAEPAAACAx0ZXh0AAAAAENvcHlyaWdodCAoYykgMTk5OCBIZXdsZXR0LVBhY2thcmQgQ29tcGFueQAAZGVzYwAAAAAAAAASc1JHQiBJRUM2MTk2Ni0yLjEAAAAAAAAAAAAAABJzUkdCIElFQzYxOTY2LTIuMQAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAWFlaIAAAAAAAAPNRAAEAAAABFsxYWVogAAAAAAAAAAAAAAAAAAAAAFhZWiAAAAAAAABvogAAOPUAAAOQWFlaIAAAAAAAAGKZAAC3hQAAGNpYWVogAAAAAAAAJKAAAA+EAAC2z2Rlc2MAAAAAAAAAFklFQyBodHRwOi8vd3d3LmllYy5jaAAAAAAAAAAAAAAAFklFQyBodHRwOi8vd3d3LmllYy5jaAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAABkZXNjAAAAAAAAAC5JRUMgNjE5NjYtMi4xIERlZmF1bHQgUkdCIGNvbG91ciBzcGFjZSAtIHNSR0IAAAAAAAAAAAAAAC5JRUMgNjE5NjYtMi4xIERlZmF1bHQgUkdCIGNvbG91ciBzcGFjZSAtIHNSR0IAAAAAAAAAAAAAAAAAAAAAAAAAAAAAZGVzYwAAAAAAAAAsUmVmZXJlbmNlIFZpZXdpbmcgQ29uZGl0aW9uIGluIElFQzYxOTY2LTIuMQAAAAAAAAAAAAAALFJlZmVyZW5jZSBWaWV3aW5nIENvbmRpdGlvbiBpbiBJRUM2MTk2Ni0yLjEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAHZpZXcAAAAAABOk/gAUXy4AEM8UAAPtzAAEEwsAA1yeAAAAAVhZWiAAAAAAAEwJVgBQAAAAVx/nbWVhcwAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAAo8AAAACc2lnIAAAAABDUlQgY3VydgAAAAAAAAQAAAAABQAKAA8AFAAZAB4AIwAoAC0AMgA3ADsAQABFAEoATwBUAFkAXgBjAGgAbQByAHcAfACBAIYAiwCQAJUAmgCfAKQAqQCuALIAtwC8AMEAxgDLANAA1QDbAOAA5QDrAPAA9gD7AQEBBwENARMBGQEfASUBKwEyATgBPgFFAUwBUgFZAWABZwFuAXUBfAGDAYsBkgGaAaEBqQGxAbkBwQHJAdEB2QHhAekB8gH6AgMCDAIUAh0CJgIvAjgCQQJLAlQCXQJnAnECegKEAo4CmAKiAqwCtgLBAssC1QLgAusC9QMAAwsDFgMhAy0DOANDA08DWgNmA3IDfgOKA5YDogOuA7oDxwPTA+AD7AP5BAYEEwQgBC0EOwRIBFUEYwRxBH4EjASaBKgEtgTEBNME4QTwBP4FDQUcBSsFOgVJBVgFZwV3BYYFlgWmBbUFxQXVBeUF9gYGBhYGJwY3BkgGWQZqBnsGjAadBq8GwAbRBuMG9QcHBxkHKwc9B08HYQd0B4YHmQesB78H0gflB/gICwgfCDIIRghaCG4IggiWCKoIvgjSCOcI+wkQCSUJOglPCWQJeQmPCaQJugnPCeUJ+woRCicKPQpUCmoKgQqYCq4KxQrcCvMLCwsiCzkLUQtpC4ALmAuwC8gL4Qv5DBIMKgxDDFwMdQyODKcMwAzZDPMNDQ0mDUANWg10DY4NqQ3DDd4N+A4TDi4OSQ5kDn8Omw62DtIO7g8JDyUPQQ9eD3oPlg+zD88P7BAJECYQQxBhEH4QmxC5ENcQ9RETETERTxFtEYwRqhHJEegSBxImEkUSZBKEEqMSwxLjEwMTIxNDE2MTgxOkE8UT5RQGFCcUSRRqFIsUrRTOFPAVEhU0FVYVeBWbFb0V4BYDFiYWSRZsFo8WshbWFvoXHRdBF2UXiReuF9IX9xgbGEAYZRiKGK8Y1Rj6GSAZRRlrGZEZtxndGgQaKhpRGncanhrFGuwbFBs7G2MbihuyG9ocAhwqHFIcexyjHMwc9R0eHUcdcB2ZHcMd7B4WHkAeah6UHr4e6R8THz4faR+UH78f6iAVIEEgbCCYIMQg8CEcIUghdSGhIc4h+yInIlUigiKvIt0jCiM4I2YjlCPCI/AkHyRNJHwkqyTaJQklOCVoJZclxyX3JicmVyaHJrcm6CcYJ0kneierJ9woDSg/KHEooijUKQYpOClrKZ0p0CoCKjUqaCqbKs8rAis2K2krnSvRLAUsOSxuLKIs1y0MLUEtdi2rLeEuFi5MLoIuty7uLyQvWi+RL8cv/jA1MGwwpDDbMRIxSjGCMbox8jIqMmMymzLUMw0zRjN/M7gz8TQrNGU0njTYNRM1TTWHNcI1/TY3NnI2rjbpNyQ3YDecN9c4FDhQOIw4yDkFOUI5fzm8Ofk6Njp0OrI67zstO2s7qjvoPCc8ZTykPOM9Ij1hPaE94D4gPmA+oD7gPyE/YT+iP+JAI0BkQKZA50EpQWpBrEHuQjBCckK1QvdDOkN9Q8BEA0RHRIpEzkUSRVVFmkXeRiJGZ0arRvBHNUd7R8BIBUhLSJFI10kdSWNJqUnwSjdKfUrESwxLU0uaS+JMKkxyTLpNAk1KTZNN3E4lTm5Ot08AT0lPk0/dUCdQcVC7UQZRUFGbUeZSMVJ8UsdTE1NfU6pT9lRCVI9U21UoVXVVwlYPVlxWqVb3V0RXklfgWC9YfVjLWRpZaVm4WgdaVlqmWvVbRVuVW+VcNVyGXNZdJ114XcleGl5sXr1fD19hX7NgBWBXYKpg/GFPYaJh9WJJYpxi8GNDY5dj62RAZJRk6WU9ZZJl52Y9ZpJm6Gc9Z5Nn6Wg/aJZo7GlDaZpp8WpIap9q92tPa6dr/2xXbK9tCG1gbbluEm5rbsRvHm94b9FwK3CGcOBxOnGVcfByS3KmcwFzXXO4dBR0cHTMdSh1hXXhdj52m3b4d1Z3s3gReG54zHkqeYl553pGeqV7BHtje8J8IXyBfOF9QX2hfgF+Yn7CfyN/hH/lgEeAqIEKgWuBzYIwgpKC9INXg7qEHYSAhOOFR4Wrhg6GcobXhzuHn4gEiGmIzokziZmJ/opkisqLMIuWi/yMY4zKjTGNmI3/jmaOzo82j56QBpBukNaRP5GokhGSepLjk02TtpQglIqU9JVflcmWNJaflwqXdZfgmEyYuJkkmZCZ/JpomtWbQpuvnByciZz3nWSd0p5Anq6fHZ+Ln/qgaaDYoUehtqImopajBqN2o+akVqTHpTilqaYapoum/adup+CoUqjEqTepqaocqo+rAqt1q+msXKzQrUStuK4trqGvFq+LsACwdbDqsWCx1rJLssKzOLOutCW0nLUTtYq2AbZ5tvC3aLfguFm40blKucK6O7q1uy67p7whvJu9Fb2Pvgq+hL7/v3q/9cBwwOzBZ8Hjwl/C28NYw9TEUcTOxUvFyMZGxsPHQce/yD3IvMk6ybnKOMq3yzbLtsw1zLXNNc21zjbOts83z7jQOdC60TzRvtI/0sHTRNPG1EnUy9VO1dHWVdbY11zX4Nhk2OjZbNnx2nba+9uA3AXcit0Q3ZbeHN6i3ynfr+A24L3hROHM4lPi2+Nj4+vkc+T85YTmDeaW5x/nqegy6LzpRunQ6lvq5etw6/vshu0R7ZzuKO6070DvzPBY8OXxcvH/8ozzGfOn9DT0wvVQ9d72bfb794r4Gfio+Tj5x/pX+uf7d/wH/Jj9Kf26/kv+3P9t////2wBDAAMCAgMCAgMDAwMEAwMEBQgFBQQEBQoHBwYIDAoMDAsKCwsNDhIQDQ4RDgsLEBYQERMUFRUVDA8XGBYUGBIUFRT/2wBDAQMEBAUEBQkFBQkUDQsNFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBQUFBT/wgARCAK8BBoDAREAAhEBAxEB/8QAHAAAAgMBAQEBAAAAAAAAAAAAAAECAwQFBgcI/8QAGwEBAQEBAQEBAQAAAAAAAAAAAAECAwQFBgf/2gAMAwEAAhADEAAAAfN/zX+r6uPT1Px/bu8/LJ28/O9Xgsy08ul3Pcoo6WnfWTrZjerlNHLMocPch21lvfD2cz1cKd87sauxbMW/DfyvSzxfOYWeN7fJ437vy+H9Hy836XDD6uerj6PSfJ+n634n1PQeX19LOsKeW9Hj7/1OH0zt831Ho+Jq78L+nMRSxVSqVLHNRFpQlUqWMqVZ0iKglJUABI4cgFENAAGjAcjkY5AYDRo0kzOxpXjWPh6ef5/Zg4+qvG7pztuLHObFWOnJ8nt8r5/X5/qy9cU7zVvLEoiWraGrZnr6byvp3X5lHfPvP0f5rD8v10c+tOelOelM6wzr8639X0/J6PR/L9Gjlxw+jzc/0+HRz3p5dLMblJn6qtdZOl3PWrjm/nJTSo6K+28t7Yes5Xs81euenlu3GrcXRhu53o8+UOMwduPD+h4/Ifa+bx/pebm+/hi9XPXx9Ppfj/T9d8X6foPL7Olm4K8r6fD2/p8fpnX5/qvR8TV24aOnJCVSpVLGVRGaQLGVKpVEbVKpUqASkqhhQjkcAIwHANAcAI0cNCGBJlo0qx0w8PVh4+rDx9VeNzmZszYcmbn243l9fneXTlbzRvMNZr1mvUhvLgmpTUNSncWt38u/vvP5vaejwW/R8nrfpfFy+H0U46UY61Z6U56Vzf5j9H6bu/P9vc+f3nnjyvZ87L28unl21celmN2ZzR0V762Y6aOetfC381mdpp9LX01l6dMHXHL9fmjc6eW9HPV2daMtvKb+fOnGeZ283B+j4vLfY8HJ+h58Hu4YvVjZw9XpfkfS9Z8b6XofN7OhnWG58p6fF2Pp+f6bvx+r7/F19vPf15ApUsZVmpYqoSksRSqVCVNRlKUAKgCAaAwgCRgOAaAwgRo0cOAaNIZ1zPN7eP5/o5+XaczJlledY+fXjcO/AxeZvDQZhqQuY2Q1CxykrWNV6tmeujn1+oc/m9v0+bq/e+T2seOrl0px0p59ac9Kc7hOn46+59X0Hzfd0fL0q35+Z6/DDWNXLrp49LuepJTs3W/l0v572cNX87ZnpPPa3ParVx9+eL0ced6eE41cd6eW9HPenneh59bOWaLx5vp83C+h4vN/V8XM+hwxezjk9GN/n9fo/lfR9Z8b6Xf83r6U3gZ8p6vF1/pef6bryer6/H19vPo68WKWKqVSxVKshUJVLEUqaQlQlIAlLAaADCAcCEAwRwDCBJI0IcA0y8u/D8v0+d5/ZNFGeb53PrwuXTgJj6ZnlLOJzKZjqRsjYqlDmgFhbDW7sdel5en1rt8WPeex/SfnNHkzXz3TjpVz6UZ6VZ6V53+Lf0f0et4fVo52npwz9OMpb+fSzFnmx1mGg3POrcW7nZRKblOluOlW2Ttzx9+OTtytzrRy3o5bvxvTy10PLu/nmnpx53p83I93l4/0fLh9nHL35UdM9Dz+zvfM9/p/j/R7fj9fU59edvl5b6Pg9D68/V88vWb+Tt7ebR14sQpYqpUqzUIFUJVLFUqUEqCAFEByMACnIAEA4BoQDRyNGAQBLz+Hq4Hk+tk59sk1ysb8/wA9cDoxdVuZZmW4zbjMs5ViuY0qZLIajUNWvW7M9NPLt6TyZ+m+z4eTpr6F+q/K1fP61c91Y3Vz6U46VZ6V53+XPr/pNnPrbikzItzbpqebKGiotkrxHCkhUNUtrsz9M5d8cvXDzZSuFIRfwpOefpywd+HP9PDF6uOP0cs3bnRvOjl23+T1dfwevp+P07vP153p4cH6Pi9Zev2Py69k+bv7eXT14OiIqpUqlUIFSqFCaQlSghKQAAUwRhBBTQgGEjlILCHI0Y0AAhjXM8/t4PH6flOXfzfHfn+sxd5LNnLbjN3PNuJbjMoLmNKo1KHCqvSrWpTro5dt/n6+05eL3Ht+Uezl7X7356nw9YY1Vz3VnpVz6U43Bv5Z7v2NzNkljM0mzNJSTuZI0aMBSxWKq2CohULYwCI1EiEKSuyjWcu81azRvNG85t5gpnU+epYpNYt8+R6PP6nnr6b5sd7Xi29fLp68ZWRlBKoUqUFKhKoSpUCggEMAAIYBRACMAhgjggAcjkYIwKOfXm8Pbwse75hz7+T1a9FU823C7nLsZuxLMGjRLGlYqTJZVqxu7cdtnHv0fL1+kPkd/wBni3fb+f6Pr8yPPVeNQ57qxurHSnHSM1xff9G687Liy4nczSdy7mSSsdgNEqEqlSoSpQSqC1EZUsZYtRWM1BqE1BqC13UJqtYNVTVVtNZ9Z62efYefbvzaunC7pzaAlUqEqgWKoJVCVAoIAEojAAGABIDCiAYZjCCCgcjhoAZOXfm8PbwJ7Pj+OvF6WvUIuxbcNHPN3OW5zKHKyNIViQsjZXoTpp5ejf5+/S8vT6v6fgW+jn6T9H8To8fOsoY3XnVfPdWOlWN1Tb+pLLmVk9ZdzKnY7HY9R00dAIBaSFAwoR0UQClUsclNRljLGajLGWE1GahNQmoTVc3Carm3V15aN8r98p7wKClSoUqEqhKSoQKCAQAA6IAGAACOCgIBhmOGiUAaPIRS4uXp5fH2+Xer4xNZt2rcszbcXTyl/OXc5OZcMBKCsEEq1mra3PbZw9W7z9+15X1X3/ns3XXu/wBP+YPHVLDGoY1XjdWd18+kM6j9SzuZXMrJWS1HY7HqS1JWOwp2FA6EY7CnZLUeoUs1ZoAs2OLGWObGajLGajmxli1CajLGahNKVE7mzWJ6jsUCpUISqVKoBKCBUAUoABHQpAMACBAcOgIACRw5AIACIzXP4+vk8/X4aej5J0Q0hVuF/O6eV0cpbnMpHIqFYhU5Fc07zG6ux32+f19Pyd/WcvL733/Gj6efuf0n5yvw7jLDGo51XjVeN146VzUPpanczuZakrmWo7HZLUdkrDUejCwGFjpj1AURza+e6uXRRZ0zf25S3FixljLHNUsZYzSlUsZYrGVKpQYUAIFQhKpQFQhAqEACAAAYAADGEAQAOGAgkBw5AIAKsb5vH3cXPo+X3r896yvVC/ndHK6eTVyWZy5BmNorECSkhrNW8k66eXp6Pm9PQ8vb39+Z6f3fP2fW8PqPb8iPDSI41DNhjdeNV43Cbq+judzPWZ6zKx6jslY7mWj1HY6djsdFgCxxYYsOe68aox1xcvTy8e6y47Pb5270ebT6OIkc6UsZY5qliqlSxlWbBRVKhqgASoQSoSkJVSVCABDRAAAADABgrAcIIBhDCQAIcggqlz468vh7eBe3xzW/L9UKlF/PWnldPG6OeZ5gyMx1UrhjkbNesxup57a+Pq6Hm9O3hv6l0+N0vX5e59/5HZz4lKojjUc2GNV43XncM6p+j0s1mVzPWZXMqesuyWjseo9R2SsKURxqOLDGq8box1xc/TzMe3iZ93ndd/D+rhpT675eHsfT8Lb6fNZ0wpYyxzVKpYqpURmlKljKCUhKgVAKEqUEqEIQCAAAELRABgAK4KYQDAJSBCVgEjCBCFLj5+jk8vX5rXb4f0c7VhU5bud1cbp5LsRzMplVDRyhLJsDMNQnW/Ho18PXv8/fpebX1n2/nq+z2/6T83b5sgojjUZYY1DGq87hjVP0es9YnrM9ZlY7HZLUeo0ektRRDnqvnuGdQmoTWXHfl8vdw8+/zW/R5Pvjyvo8nF78OL6vKMfe/i/a+v5+L1vR4NPfi6jKs2IpqKkJYyqVClSpVKCVCCVCUVCEIQgBFTQsAAAHIAADUBRXKAMIJRBXIKSA4JItYOPp42PT4zp2+IdcVa0LZm6OV1cbp5LcZHNkaVrlZLMlMx1mNTz208vTr4erbw7+k48fpPt+LV2z779L+Yr8WwMo51GWOLDO68ajNZ/f2nrFmsz1h1LWZajsdOx6gVcd5ePfLy75c98GfXyp7ODr0+R6zx/p8fnfV5cvTlGycubfPk+jyfQPn/T/AEV8rv7Lv8TX289vXmpVNRhNKEKaSqEsZUEqVKCFCUE0hCIgKxIWADAEEFAsIYUQSlhKwAFc01ADIAcAQBEM65nL1cOd/mfXfynvIWylv561cbq427GZ5w5kthTlbTZnmNmFy5b8ejRy9ezh6dnHt7fHi9p7PmX/AEPH7D6Xw48NNFKs2Msc2GdQxuMub3d7NYs3iVzLUlqS1l07HYaLNzebth4+rzPD7HitfR872x5/tw5XTlzOvHH040dMRQsjZVrPL7+WyP0R8L7/ANf5fK6nfxau3B2RmoyoFjKKoiqlBKpUJQUJUoqERAjYIWIACwgsYIBQMAAAGpAEEoNSVtA8gIACApx05XL0+fvX413viOtrtsmtHG7OOtHKW4y5kZLqNrlaSksxHMx1iU6Xc/Vo5+vZw76uPX6fPl970+Prfd+T3c+GSNCI5qljNRxYY1Fcvu9Fmudmsz1mWsy3GktQp2Olm5PL35vD2fO+f6P417/Lyu3nr1AilWsVbzTvAlVzRvGTrx53XzfQvn/W/Rfx/X7bp8fX24XdOcZUqVQpUJVKCVSoSgoFioqEIiiuSkCAgCxoU0AAAQURqArAACUglYBNNSV5AAEZOfXkY9Hm99fhHecDoS3Y3p5a18bfzks5GWJpXTJZzPOZSK4CzHW/Hr08fXq5d+h5t/Xu3xn25+z+/wDndPDnJGgLJSxmo5sMajLm93eesW7xO5lvMtR2Oix09CTP5+2Dz+ryvn+58h9Hr8H7/n0XnZK5a9Zq1nP05w1K7mneM2+eDp57F++/E+79r83h6fXxaevF2JVKhSpQSkJUqAUoJUJUIjcpCkCILALloAMAAVMAAAGEA1FAUyJQcDTleaARXDy68fHfyPXr8B9GMug1fz3q5a08V2CZYLG0VpPOJ5xLOY3Nk3fz9F/P1aOXp1ce/oPNPqfr+NR0n0D9D+Xh4uj1JIIs0iM0soY3Caze7tZrFm8T1me8ysLHTsKdO5r5dMXl9HH4/R+d5+/8q9/i4/Xg0lFOs0azRvFesValOudG8Zt8sHXh7Xx/Q/SPxfoe218vV14Xb5oUqVApCVAqEpCEoqERRWJCkiChC5EAGgFAAjVIKAMAAFBwK1AglJSGrmjNZWvN5duLOvzzv0+Kds1bOa08ta+O9PJPKNkpqFK1yNizObMScxFLsdL8erRy9d/Pvt4dva8fN731/Nfq8/tPq/Ah5dyuXTQhSqWObHOo5ub3drN4s1mWsz3mWo7AdhTp2Rxcnm74OHr8bx/QfM/T08j6fHk1zjVes59Yo3mu5q3inec+uVNxh7cRPvvxfvfaPH5en08mnrwlYglSgghWkJUAKhKhEblIUkQWCK5YAjoQGAAAgABgAACg1IagSgSkrgmnNUy8rn24F6/IvRr5t0lWk5rXxuvjq/kBU5uGqrXMSmLcZnnIzKW/n10cvXp5+u7n23efr9Lz4PTenxdH6/zO9nwuSbLsYBCljlHOo5uf39p6zZrM9ZlvMtR2FjCnREc2jj0y8PR5nz/a8Nr6fgvX5OR14494ruatZp1mnWad86NZy7507xm3zw9fP7TxfQ/R3xvo+2x4dXXz6OnICVCBUCoSgAqEIVzFktiiCwuRAaOhGMLAEAEAgAAVoTQAAoCuAbRBKSkEubPTk8+3menT4P63l9SedXc9aeWtHK25JktTUbtU5zszi3GZSWZjmtHL0aOfqv5+nTy7beG/r9+bs9HD0v2vh7PLznrMkky7GLNQs2Msc2j39bNZlqTuJ7zLUKLGFOlEeeque8/Ltx+H0PKc/t+I9DyPp83M6cc9xVrNdlOs0bxn3jJ05Z986Nc8XTkrn7x8b7v2r586+/Jr6+edyhKKCEqggotBERXIisQhWFwACSsBo7AYWAgEIAAAlBgCoFcAKSsFJXKRGXHjrx8dfH+jf5t+lw591q49dnDe3hu/kszI3CaV1DWmSnOzGbcZcTl0c+t3P0X8/Tdjv0PL26vG/XvR8ujrn3n2vzdfj6SuZWSmZWMIUKWMqlp93SW8zslcz3mWoU7CgdBDlaefSrn0ycu/neH2PKb+j4L1ebgduGDpyqsjZC5z7zk3yx9eWTXOjeM++WDrw9j4/f8Aov4/1foHDy7Onnv3xFUJSkqgAKVAhIrkpJELkuQEB2MLGjsB0ICEKiRAAAOUUHKhiUVwlY5QJoiM1zufTizr4D2X89e7zc7tJ53u83foebpt4WzBzJUbuGrKJ4xbnM8hLca0cu9/P1X8/Rdntu8/X13Dn9J9Hgj34e5+t+ez+PvJmVjslI0AFmxlUV/Q3O5ek7mWsvUdOwphTsji1ct1c+lPPpzOPt8vz+14f0b8P6/HzOvOq5jIrKNZx9OWPpzxdeOffOjWMm+Ubn7l8j7f3H5vTsvNq68JsJUoqEoFgKhBI2FgiFYXKRoWA0B2Fy6BhQiEgqABAOCiABzQEoCikCuVilrzrlc+vEvT5l9CfLvTjmTPOqm60ctbvPvTyujmlnTWNsonnNmY5HF+Omnl3ux6bufqsz12+fp9K4+f2nfx6fo/O9Jr5656ky0lY0aMBZLNUsfpWVjslcy1HqOnqOgYWEQ56q57p59M3Pt5/wA/1vH6+l869vm856PNRrJCsp1nLvnh6csXXll1yp3im5x9OPrPN7f0P8b6n0Tzctm+GjfFKlBAIEKLEFyrASArgsEYWCA7FcuwGFgACABAAAAwUgAFcpKApKK4FzZ3yefTib38v+i8p2nMObmYMzJVVzdi6+O9PHd2NuJSTzLMglGvj30c/Rfy9Fme9uO23hv7Hy8nY7eft/Y+L0vLhsySVjkkjsaKUhZRmn9NK5dkrHY9R09R00LCmRxa+e6ufSnn05vD2eU5/b+f+m+A9vhx7wSFVazk3zxdeePfLLvlTrELMfTlG5+0fL+z9z+X27ueWvfCy5QlVgCFCFJC5KSCK5LAEdggFhclgAIUACAIKAAAoAAAKBDUlJRXKClx43ys9eJ038297j9Jz9OeYYx4mKZyJTqKrMb08t6+N0crbijNud7OHpv5+i3He3HbVx6dHhv7ffFV15+4+t+dj4+jslZJmUCSQRwpVCzZ/UzLUdjsdOwqWo7Cx0ABDnqrlurHTPy6+c832fEdPofM/f4vPd/NELK7M28YunLH055t8qdYrZo3jJvn6Th6f0N8f630bx56N56t8UIVgCFOxBclgiC5LkBGhYAi1ksAQpgCCCFAUAIcCgAoEoAKBBK5oCaRz8dOXnpweuvBe3WTUx1irHbiTJmYcuczg6YydoSdDzb2cN6uOrsa08euzj6bsd5Z7SnXXx6es8r65ryZ+3L6B9T87l8XolY0lcuSQ0aEEKFnVn1uc9ZdCOgKlqPUdgNErSMtfLdfPdOOnK8/v8hy+1869mfnnu8GbeIs1blNzj6cse+VG+dNzXqU3GbfNWfYfm/W+5fK9PpOfPfrnZcRCxI6EKGSwsAQuChBHQhYXKsAQoGgiQooCiCgBih0QAEopKACqHA05a5rmY6cyb8x6L5L1bpKaoXOuSsSY6wZYs552s83txq1Luer8XXw69jxejZ5/XfjvKdbMddXLf0jyT6Fvy3e3w+t6/Mq4dJMupMtJIwkaEKVS3/X5S1HYAA7HqOnY0KIKCHPdXPdfPeTj3835/s+G6+35f7/AAeX9XkjVVzRvOTfLNvnTcUazVvFTNepRrn6Hj6fv/yPq/UPC6Lnr1zSFDIPQRXBTALzdCCOhANYVgNFQggjpUBYAAAAAAArCAUDRKAoEOaozrlY6cu78f6tcLva5qBGWtqis5ksyxhtxSc+YwM83vyy9cW531/D6et4/bu8/psx2v59dnHX2Ty49L083W+n8jr+bDSUkkdjSSORghklUavscXrL0IAHRUrl0WAAAs2rnuvnujl15XD3+P5/Z+ce3l80+h87D05VazTc5enOjXOm5p3iq5rsruadYGfrXz/rfdvler1HHG3XO65SCOx2q4aFCFjuRGhTQDWBChCkjQsKQWAAIYkABWAAAAoECmaKglDLnpycdORrXivZrDu1rCaJYrGWEVLnrJWMwJhTnzPL3yw98q76Pl9nT8ns1cevR8vfp+Xf3vz8X25e0+h8OzzbkjmZSMdjRpKQghCl2fa871HQAAjp2OwoAIAivG6ue6eXXJx7+e4fW8H39fyv3+Dx3t8VOsVXOfeKNYpuatYquYXMLK7mnfPvcfT98+V9X6p8/XSnPZrEkEdhYXLpghcOhHYIwuVclNCgGCilQgAkBI6UFgoCNRSGqAAUCDNGgUuHG+RN8TrvxXr3VUJqKkpmpVNKWuKazW46xpz05zHH3z4nr5Rvfd5/Zr4+jr+L0+3+X0+7ebjk68/ovv8AgZvJ6XY2ZSOR00aOCQggNv3PM7HQjAKB2FA6BQQRGaq5bp59M/LtzuHs8ZPr/Nvbx+XfS+Zy+vCnWKdZpuKdSrWK7muyLMbKtZTP1LxfU+7/ACfX6vhnoXnoQZdFy7HTQuHQjR0I0WsFCMLBBDRIUCQFYIAAAAAADUUgUEErhKSxXm43yZvznfflPTqFsJYSyaeapXNKWMsCpc9tBirk3Pmevn8/6eM511cfTXrro5+j7X+e9f17ycl6PN7r0fJz+bu7GjkcNGjHI5EEB0fveQsYDQoAB0WMSkKCFLXjVPPrTy6ZeXo87x+r8+9Hb5X9Dw+G93zqN4rSjWKtZqua7mNzCyNldzXrHa5ej7v836v1T5uuznnvZlY7BmVjsEdjsLGjGiuTUEKEaJHorAEQWJAAEgCAACtQAENoghQNALVm8vPTkteT9W+D21BqC1lma5olcpNLOkUFcsbrn6zwdZ85044NY2cu2nj1xbtet/o3879D2vPn0Pf83s+WmTRySHIDRyNAMlKR2v0fzXYDsB0UQAPUAhSqFKojnVXPpTz6Z+XXm8fb4t9L5r7uPyr6Xy+L38tVldzVqU3ENSu4VQuY2V3m0+l+P6n3T5Xs9n5c9Nz02FyyVjsLl2OxkkELksKEKGANCwECKwQBEgCAAIKcAKKKACkpBCUXNm8rPTk614v2b5m9RtjLBZSvOnNPImiVS0W0FFcrU4Oscveb+euty6zzrCLlr9HfH93QnP0Xu+ZD5/vcjSUCMIcjkIIBSkes/U/AdCFOmhQKAdAQpVClUsZYY1Ty6Uc+uTl6PP8AL6fz309PlX0fD8/93y8+s16zXZVrNVxHWY1FmNkLiNx1+fo+4fO+p9c+Z17eOe9iyx2OwZlT1GjSSK5LCmhQyIUagAkSFgiQQEFAAAgGAQ1QKKBKQQmsWdcib43TXiPZvPbG2MqmmrzpyuaIM6UtFuHU5mpwtY4usWZ33+Pbr8+hLlTVx6fevkenH1x7Xr4eB8/7TkcjRwQ5AlIQQQoUv0H9f+Rdj0B2OgAFAAQpVCzUsJVnVXPdHPrm5dudy9njb7/mXv8AP8p+j8zz/o8cKrua9Su5hrMbIsxuY3MNYZ9G8f0vuvy/Z7fyTpsa2XY6dy9RpJJIIayUIwsGQNCxAiQRWCCIKBIAAAAKAABK1SgKZRl503x5rz3fXj/XuKqWKksppzRKSqaJaDDpydTi6zwdcs8noePp9Jy7dHG65YR6Ty9vqHz95O3Lu53575v3HDkY4cgkpHI4IJFmrL6v+5/BPR2OixhYBKoFUEKVQpY5sZpRDG6MdKOXXHy9HDx9D536t/K/f4vnX0PlZtYrshrNdzCyGsqyFwkhrEbnqcvR9t+f9P6/8zp3ueejcW3IkrHqSGjZNZdCMVjZQ9FYAiQsQMiIVAIAAAIAAABQcCjSCWuOXnXFuvIerfm/RolFJZTUpqU0SmdRaz2c/Tk6nFuOHrnhZ7nH0el59+3z6ac6WbCPd+Dv7Hyp+vx8b5f6DB4/oPJyyJZspZSNJzLkcjyWaj7N/Qv5tLQsaOgdgIM1CUlWUc1ZsZqMsc1S1Y6Uc+mbn3wcvV43ft+Y+7h8q+h8zzXp8cLIazVcw1I3MbmNzG5izHWSPf8Al+l91+b6/d+J2JjZcOpayySNHcljsAoZEKLCkggKwZEQqARAAWAIAJWEqAFBhKKlol5GdcDevA+3fJ60Vyymp5tubZNOaiuSznavJs4tzxdcsR2ePb0PPv2sdOpjc81SrOvqHy+3VxNnq8PiPifro8ujycs5Z4s5uwnJJJSOZeRkpftX9E/mb1GgFMKEFMkKVSrJSxzY51GajKRXnVOOtHLrl59+Fn2/OvXr5b7/ABfN/d8zLvnVrMLK9ZhrMWVcxuVZG5E6HPv9l8X0vsPzOnp+GOlc3ajskjGyXLosKGRALCwFQhYmRAVCIAoRWAIAogEoE0AAlcAlx53xpfNd9fNfd0yarVy2RfNXZtualx28zV5NnGZ5DnRZ2OXfuc+vXz06fPWzOyVS2c9/XPldsnTHd1w+f/G/UrOiVyzzbM23OrSyJjzHEpDIy+1f0L+ZmhAOx0AIIBSrJSqWvG4Y1GajKBUMapx0z8+vP5evxfX0/MPfw+V+753l/V4q9SFzXrMdYiykVzGwsSEvs/P7vunzvZ9I8DsTGzWZU0bLZdFhQiuQGShCihCwZApICALC5ACgQBBRKAINQzTO3EV5+dcNfH+vfzX2dIKyyW+XRF2dRXNXMrkVx2eXcTzrsY69jn16ed78a3Y1POoqL2fN1+nfO3zO2Orz14f5f6JZpEpZ51ZLbm3RYtmUsnDkII+y/uf5uojnTR6OwEEqhSoJXqKVRCahjUJt2SsdlWOlOOmTl6PPz1/Ovbn5X7/H849vzcPXnC4hrMLmNymSxBYkDVjr9g8f0fsnzenqvPOgxbY7BGjQsKaFyDsEdgOi5aCFFJAKELBAKSAwAFAhVHpyhvDzqfHu87hHMzrga18/92/C99izW6NC2ywMtvOTlVx2MEmzPTsY7dWb6ONbc61Y1dmqVSymvYeHt7Xy2nv5+X4/r8LxfSUpLPNnm2Z1bLdLbLOJ5SkM1wSfYf2n86hjVWNwmoZ3FZXJCUWIlZK5lZLWZpVndGOrLOnO3WIy08+tGevNx6fC+nfy73+f5n7fB5b0+KvWI3KsSRuVYIrECNfW+f2fZfH6vo/gva5Z0pJJ02VYho6bLQsZKyVjR3LGjoQoCyFkUQIkQqCRYW1JRY6xR6PNn68ZZ66vJ67OXWmORNec6a+Z+/fnd7mty2qFJkrBLy7nnTOU2569vHXpTW/N2Z3fnV+dEqyJZZ19F+d37vHN3p83hvmfpauPZTTynnVmdWYt0t0tss5Z5SyeSgj7F+y/nMZqrG6s9MPH1YsehNSuQhNVNpZJo1y1a4ad8LNZox0w8/VVbfrnr6cZ3MM6pz04jv4T1X5/7OHgvT5fI+jx5N869YhrKSNyrlAK5AL87+icPZ9L8vX1vldjlOhia2ZJTWexE0aRsjYFqWyXpYTuZJKmkqaGsxIEbkCsyYrclaY3xsjRa7Ya5Y/Z4sHr8lnL0dD530tfHtkjjL5bvr5h7umK6ulmsbcyY7ecc5OfMua6OenUz06c1tzdU1oxZyk2oUss2ed/WPldq9Y6useA+Z+khjREs7nm2Z1bm3Z1dLbLPKUssnkSmX2f9r/NXpGWnn053m9nL4+/nz12MTucztndktDe55u10+bbcU53lvTN6pyOu8OcdqcehjNmbVnXI1rw3qeM9HLyvXn57r58O+WPfOjWa7lWRsQkVgqZjXSl9lz7+r579f5r6XzzbjJZTWLaqy+KtZq0zaiOhjPb4562ZsSwssmkrJU7GisiKxEE5xxNOF11txe5ydPN2FlQ3ywe/wAHJ93gu4evrfM+ru4dsEcK3yPqvzn19Xa4ptyViXmnLmcsaW+nN9Ka6Odbc705t0s82MqVTTzXL0uPT6n8zryO2O3x6eC+d95DmpTVnPVk1dm3Z1dLbmziUPNlmkLL7r/Rv5ZLcNI87k8fo5PzvocPh9OubKyzvnu82qJ0pyimXWOP2zX7c8L33z/pyt8uNjxdzjru+Xp1fLvfyVL5bvPBerl5Xrw5PTlGydk0mrR1FIoqEepanXmu5m97letzbczLvPP6Z5vSRNeVOpmucu2jWPYcOPsvLz73O9BbCdSRgA0VCA1rjnxxq4m9X5vb5XpxstkQ6cuf7fn8b6HzdHD19j5n1tnD0cvLz274j168X6ekVy24awnOObMwzd7fRm+i1vzrZm6s6tmiEqljKTTledeo8nX3vi1g786/H9DzPh+oLOannduNW5ts1fm3S2SzylK4lks0j73/AFL+TljojP4+3K+R7+H4vp4serHO+S9c2pXvHQcYHD664vXHh/Xws+p05vsnL68X05+V4fM9Djp9C+d6fVfP69DjqlfPdXjPTz8t25cXfCVal0y22TJEyRK5sLGdFzqjbi2FuWlKdZ528pNmW7C+SNyazvvH1PHl6HOOouu6kO5WsxSq5quVTsdk5bMdDnujGskY1nLtjVbaOq988Ps8HF9/zL+Pp6/zfsX8O/GxfM9NeH9m/Ld9YNXmW8o5+ZmLZerN9JrWuyXXnWiWzNJqKxlJSUlJpzXv/ndvT8LPrw8h4fvZ/N6ZZ1PNsmrcbtzq3OroulszZw1lmORZEfef6n/KlABT5unP+V7OR8/6HN5eznY9uadVqT1jHqcrrnzvfHhvT5/H+vy7/Tv0vs9WjOV5OGjly6/LevhvpcNb+amuN1nk++PO9eeG50GyTUaEuTVGrM0xdJomdmZ0MTdnOiS0VldldkmbpLkukuSxJazbvnp1jZvOrVs1ALI2VXFGsZtYpuI3LuXNWc+s+fWfPqQyds1mMZVrOL1eHl+z5zx26Hz/AKb83p4EvD6b4XbXK665HW+c6XzJyomvTa6s1sW4smtEujJSk0mozRClavOpZ19T+X31TPQ1jk+D7HZnMl1ZaWWks6qzu1HJKJJDOr7zqzu5zjNfSP3/AOLjz0sUKuG8Xz/Tg8vrz460c+ubPeqbS5L05+ted6uD35msWWba1SSwuzYywSmsdcreedqXSbMS3Msk0TN6XM3yaZnRJYTktzL7nVc69407xf0w7EjslZKx1KyRKxMqwsnqSp2OxWK5hc02U6xVrELFrJEuXQ59J46TWcE0SiskQ1jP382P0eVTV/l9Wbj281dYdbgtdtGri6Xm9Nc7TBrWW2FsbRqJUas6cpLBpTThTRLKW/G/rnyevJ7Y955d6uHXvenwWXNlzn5dtXTgpcnH0bevmrm79cVbJLunOGdWucda3/b8kOWq+W44tPn65fN3q57dRljNRlhN0zefPapqpqJBarEVTVU1z2+VpzdTJrNWpok9DynX550ZwkGXZJmes2MySSMaWXN2s6euNXXnb1xXc54iXazds7BJWFiSMpY0dRAlcvUKildlbNdkLITNWdVY2RKHLHO450pZ5tubJmHblV24rUfPrysXzfXVbUbYqlSixtjLVbVbTrVFuLV5+7i1c2rW1CE0poglUvofN1+i/P1i6Pa+Hvhx6fZ+r5XX9Xhu6c7enOVjsdjp2ABRAEIx8vTDluHLVXPdHHrTjoVbvE9YeolhndeNPSMtOOtGeled0zpVNZJvnOnL0orPrOfWKNSVz2eeetzzqzhwXM0ncWJYkrJWKAdzbZdvGjphdMYU4657NtzsubdR2TSdy7HVms3aOyvJSysdjp2Fi1mNkbISRlrSsjCzYy1ileNT57lBcrpk3InPl4XRgu4rGlNCpRSWFrlnLJolTVbWese7zt3nbvP3rMtcRX3Ph6+u8t0469zyd8We3W3w7Ho83W7eLRrld05XdOctxoDR6AWAAcv5/vUKVS1Y6V40Dp2IjNKWE1PWUtGOuXn2z46550zLm0x1i3Ob0zzemeX1xy+3PDvHQ5zr8WvMvNDOrM38m7ndMlyW2W2TZnqW3NlkrMsvGb8x31m1OhnPQmbElc2s2pazdrOjeL9ZdlcRyY9SNjubd5s3HrMLIpGIJCyJGM5mWCuITVCtNCaDRNabMKeN67y63PNlK5pqNBi1YW6+d1Y0K1JXNkpKKms64dXmbvo/N06OHa4ddPk9l++Vi2bxdN9Pv5ul08enfG/fK3pzncvUlqAqACOD8j60Vdk9Zeox2Oo5ted141XjcZqKi0Y6Z+fbNNY2sdY95zXOHpjndcc/pjndsYu+Ezr5N3Nty3mqZ1Rt5tmLdJaW2TZdk0nUrImbOue6ZrS5NZLFctJXLZdktYncrWcus0aNLCrWc+5KZ0ltzVpTZIkmesmkbLFpTB0tGkpWV2xskmrDqcr2eaJwN6wdLh1qm6nm6c2/NtzoVLZlPNk0K5Rpyk05XNSmnLLOp5sLfoXnS5enqd/LRnpdrHO5evn30z1ejfJ1NeHXrz375XdOdvTErGhQEeP+D+irai1ZrE9YlZKwFmxlhnQtt5265iwzqvOqs9KZqqK7K7KyuyKG8reRHDRkrALI3MN5jcpHYXIOxoWFlctNqR3LslctJXLsLGhYXOfU43VyOqclxGzLqUaRrRlHUzaQqyJsozaVaSS/KUU6Z9ahU5LEgtawVmvm2YX4sLcetYN2jVnnWzF289X5s5XmtXLJXNDRK5ZZ1OalLPGsvR9V8uO325Rlg1W3Cbg1VN5p3rmtGeWlw1Xhp1xv3ys1izeJ6yV8/wDz362dzGakzK5FisVjNQmqnSJ1unh39fI0rm6s7hncGoTSIrEiRRWRsrsrSCxsURImLtx43t8tHTEkklkk0ui9LZLEnZNLEmzKySTskk6dy0KzpBKNOTtxOxLqzBK9M2mXYs1ZkbMmrl2km3EDLq5Ng05mrCusXS5tGaMaDn7vO6Mizjrc3V43scU5cW9crow70TW3nelzurGp5rm5K5W08pS2Z3OWzFydZ9Y8c73p4GToK5rPjtnz1SySesKWGdLOpzFlxded15yZ8b8j9HfvkSwllZKxSwmoTVbdc3TOmm8du/MqrzuDUJutqE1EharIJBK0jZXZWzGxXLuRI1zPTw430vBLXO+TfjO7OdMaZL4vSxmdkkmOx2FggFisEVUSUisDJqcPpcW3QxLQKbOT2cjs1Zm3OYHP6ay6VanR5NWZVpg3cu7dmbsSyKNMWtc3ppCty1ztsd1CVYvfxOtw16bhNnO5tXkddcrpcutbMXrcdbeerJZTTmrcrc6smp5R6Z+s+WdTvykjUhQClhnVGetGOsJqVkrkITcJqGdLN43zfoz1m1zoz2EncyQFKljbCaioCxVCljbGagKlqRYjYIrEkbkIWJhWU3PD9HLmfR8WrXCbNfTnszno87v53blfZJGyUxI7kpoWOhCwQsilJAmka5e55rvcXR0MKqhVNlG5j3duJqzILk3MO7Ct2M68VWYN653SRrTl0uVS8rpeZ0vN1celVsDE1IMaZ2+bs8r6bzu9xspeP0vG665vS68Xv+fezFk1Zi3Z1zuk9T5+nuuXLo9MadyzUJSQlWQMUsZasbpz1pnSE0ljKl4Hy/rXXlYxTOrsmjuVKlQKlgsWkqFKqFjBSqFlVlVk7myYaOxVFBlWVJxe2fC/S8voevi24xj7c7U6nF0eetmVjMrJMlOmyI7ALHYWCNFYARR0IFVnmu18X7L0ubbiYOs2ZwS6MvJ+rXRw3czIVns43W2nW55y6tOlK87d6XO7czPpzOl8/wBdZ7MvPpTprsjjVubsk1c9aksmvW+V6Tha1itmbbNXZovK63i9NYrftfzdTc+j0mvSwsW1JUQQQClIEWbDOq87rzuE15X5X3b7zmzCarahdyuWKkqGBFYtKWKxUEhqKosQSGoBcTS9hpXUCFiM+pwvZ5a/Z865kuSzZhpzLYlZOSVjR02QkjR07mVjQRjpoI6EAFWbU+f+rXmu99VxzOOP2b2IRs5b8V6Nd/Gbbmq22Oda83qMzLJJRJbYnm25aJdUb8WZqt6+nX083yta8+66GWyWteti+S7zga11uWu959cHfTm9seD9V9P4+n3ThnT0xk56rxZ22tRzZS3XNly6JCVZpBIlMnCXx3yP0MonrEkhNZnap0suK5tFtxVdgIBK6UsVViuUJlXIkbI2Oy1mxiTKIGe6qrm9+fO+j8exLksktZnFiO5nEqkjRsyJWMdjuZAyx00EdjAKYrAgnN3fDenXZc6V9f5WzMkGrVq7i8sqy2yyVWLOp1KWdkz5F6J4zV9Pl9M899Jq6+qdDWXlI5uzqspLTypKpSIykrFKppSxzbADNWbDOo51XnUZJy2WTuWizVKsIqZEKa8t877kc0llclQm8WfRTellxGas1iMqJsuRGbUC6WRGxFdkUdzEdhZJkSSSsilFI4HbOf6XyNDnclszNJpNJEiaNlpKmjJWO5kNl0wLEt1LdLqu1bKe5eW1bU9LCy22p1OxiIwCFCAiilSKIw6UeR26bXo9S/dYQAsiEQxVCzUrmhSAYpWOUmlBKZrBTJizYZsMarxtSyseQhmxlWRikeT8H35TM7mDVU6NOe9OVq6Z2znKyrOywoZrspsrQSmq9NGa5qFV3LsmzYFjZsZkllyWZzm9HB+j8zpb8sybM0kk5dFzqrRrN+po00al9l25dZdV+1lltW6XaSokikCmMq0RStEUy0maWgymRc5nlpXPGVcpQZVxrlMlZy6ve45fSnL1HWatmOlKU5ABtQy5/Cylu0mrG2AEAKK8pSojnUM1xJXDCVBks2GdQxqObHNJVhHnY415fxfchNqLbiDSUTG75bbIq1IWFmmGyXNNldzWZbMHXOTpjoc7rxrXjUFdgKxyWsskkkLnJufPPo+c9Hk7OM+k3y7np53aWJC3PLVFEucyy8+6imaKDKYjDWQx1h1nFZlTPZRVFUpVqU3NdQ1KyrUqKrI1EhKiROGtkWF57TPP7BOf2PE9DtZYKU4ABXRABCWrnc/KxiSztstkAwUCVq4lKBLGaWaRIcIAgyYoWas2ObHGoYvivk/prXODVc3G2bNE6CJabY7zXZGzn9Mq4iKKguM2poyw7mTrz+hd/D6Lr5tEsanrMrLNZ4fl9fg/n/VlZpmQprxXq5/MPoebz/bjy+/LN2516xDfKGuZDCVkpWtkC2RMJZlubYW5u/F2ZX5t0XS3ZuvKc1rjZhbGrLQlhKro05TldNLbOlrPpu/D2Pp4dbtizQpwAAAASorjNmc3nnCvb3roascWvFrzVNSJEqY7WCuJTQpDVyuAUohBBKZjELNIIUvzz4H6+idLrzuc4rndbLiDUbLEptx7wbyZvA9HHN0yoqjDqWZlO5fM/S/b8z2vXyu5dKwHXO4dvA/H+3km6iOo7Aplw7zyOkx6yWXRqk2IZmg1SbZiVYrq650zNlySzQhFqTsaNNGJbcEjsLLEKaQqdyrLbLUdG83dcbO2dvfG70Y0dJPQAAAAAjFOZlw5vLceet/TNu5Op1bqzsAljnUcWOaSqGshgNSbY1csoYDtIEcAQQQpVkZKX5p+f/aSkzul+uV7EZrPdGsKynS6Tm9+WfUvxc1z531cq4jJj0x6zsuftXo+X6Xr5xkp2FBzsb+JeX6fn9X0vnvpvJ20VCWAlSyZ0M2XOtzlZYyy/XOdgjsdlcprOy85FapJ6hDsaOipXKppAnvM9ZiYvP0nZZ0zd1xq7Y2986e+LtnYDogCgICJVlmw5fK8Llvzs65sb73Oet68tnTNu5PazUlTpiBWGaSxzpQSkIJRskJqKuJrImOGADGEhClWRHyf85+60sSTG3frnYjSuye8xyjqZumMm8VVdiec9XHmdZnshJ1dY+xdvmd7pwNZQ0KdQy+beT6Hmp143bNWs+pxj7H6PnshKpYkFElY7J2NJajR2PRo6lY7HTsEKEdA6B0wpo6AAo5alZPaVjAAAAGIhFWZky5fO8bl04vLry8bwNc3q5fXXtvLj2V4djpnbuWaKUESqzSdSqVSpwK1GoyLJSkqlU0pXLIayhjmgcNHYZABBl8Y/PfudJrzAqstuVc1WCPWZhc5d5wdcZtYpjh+jlT0nuJ5vpHfwbuvKLJTozk0npJK81CZdCaOuXYDp0CuWCFMNRghQIAEgNGKgYDoAYUwAAAAABEZK4rkz5c7F4+OnHx05OOmHHSianLAUsapOZu17xk6562cesxz9LOfbTdqzqQwmnU6nU6nTV0x0oJVBKpqU1KWQBLDO1mhJGhDAIcnxz8/+3z2WWbMrJIBvKua0VhrMmRKNzD1xh3ivWfY9PH7Hv4LdYVjsdSZdT1LrHYKWNksY6SAUkKLJDJAGoQUhCQFBYxo6dMdAIBRBTVIpUleUJYZtOblxrJjeDO8eOmLO6sWOFedW5s4aziWZMmOIVAjEaqrLpz+s53fJ159Fz69593OeqbNL7ZklCSyqY9JjqSiO0mjOpyyGJa86hnZKxyOAEcjPj3wP2cii5Wpoi3Cqo6zHWa0jZHWbESVbxg6Z5nXEevOjrzq6Y0SWSWpswXTFmueu5ZpyquVTubmboqq/FrGkoWs6ctuL0eV1Y0yKwlhLGWIxySqSTuXZO5dOyyxpFaqrWaXyMURzac7rzqEqlUorVEYjUUiNJXIjQAjAIZEiJWiWFZ7cPSc7pMfWHTG28+xM9hnpJsq22ZMlUlla1Y6krlnbIlVUuWazzcpqc1bJblYzO5Aj4f8D9ldcyKNZsLEvxVZVvELIWV6y0TNVldmXeef0zj1nFvOfebM2tNObn3M284d5ivTyimfUsqW8dPCnpjpYdDCWs3yGs6Jl2acOlxvQ5a289WZ0xKpCnUkaSuZpIaAVNLEIy2xXQljLpjZKjLFa5YqgHRCWJUtdtdhc2I5HIpYKSBIaBG1ERiWtcupi3MnSQ0t1nbc9Fjo3PQTVpdbasqY1ayidtlFZLOT2zj3rRz1djVuOm3nrXhdcA0+H/B/YmsXxRrMbLEtNGbVZn6ZpuI2K5dzGSFBl1nl9cZ95psgAktgsrsy7irblCyGkN46EmiRanQxJ2Gsw1mdlyW5iS46XDXQ5Xbzt+bOV2pBBCpJJJJMAGWJUsVvRJGrInclRE0oRWZ7aNKtIjlcVaU6U6sS7LVhbiTklISiSGjpiEoRlgsEgsFrrPq59obzdrOq56FxuuNdmlbCbU1sq2lZh3POeuc3topV0fPr0ni11Oe9BDOoy/Fvi/qVqXMxspsNSyL4uzaNZy9MU3MdQZViBEmbUw9MZtyqxIrIl0XSQqnUti0VkbJ7xvxCrLNEjsq1mO82JfmCWF+Fktqbec287dNMmEOpSNZJKGBJJCAmQIVZCqGjScQKra2qbqqyrWVZOJyQqvSApqzO9GNShxOZmjkEVjJssJQSxgWMsZWK2NtdUW0VTtHebNZ13O2TTGtdll9j1MunP6s+7XpVpdm9rza63O6Fhjcc6+L/ADP0bSxm2lZVYWWxqwZn3jLvNGs16kUjcxSQFVzm3nPvMKo1I2ElpfKAk6dzOJ2aJmwkmmJIWQ1k1LJmZJmyLMo6Q1myy/MsWyLc26XRlfi2QSitJIyQIxLBams22fcdlkldUXUpbs2/mvktuJIEQqJLOksZYrFYrFGCCFzEJBZJIlk5RUIKRG2C13UVrWuq6r0srXqbpNLOixVVqwqvSGl0dPDflPOozpKPj3zvvu5EtLSFlOslzbJdmxqjWcu85tZrsgkdZlFhMimXWcvTNWpVcodWy6oskdkbm2NCXS2zMklVkkqLGwwsLJJORlW5DUViuSpJsxelx1v560YTUURgkgUAqlqu6bat4hqIiKpS2Z1fhpmbmbEkAglQitYrGWtqAhKhWREiBCJRKJwxwUhVG1LFY3UQahWa2nUW66vjXJqLksgqdapNEDUZqcfKvD9s1JSOpoyOpTcFkkslglGpm1mnUruVY0lFhNIGfWc+5DWatBLZdEmrKyQ1mZbF0k0lTQsKlZZMiKlYrmVjiNR1IoWKwsmm7nehy1r52+W2LYmFJYgERIqlVJAFEFIFmkkmkiQyIClQgK1qaioMCBBYggRhKDknAAACjStSgim3Jq4tbzasKtL6vjRJoNEl8WkppyyPnvk+qWKxhZIbMUr3I3AW5qKNZqRakbCySCTSUKyqzPuQ1I2Mti2TRmWrNJWWSTQHUkaLUkSZEjRYItZKEVFgiRWFl0a+etmLq5tMt0WSyUACIhANQEFaIiqGMmkhqkSoJWkVpapWIyaAiCskBWtcsFSoIYxkoAVxICLVLWXWsrVGlRHWkOp2Tq1bCyWSzl854verGisdCSZkQshrMLBmyVFdkbJXKRUWCWSWSxshqV2LUrsWszksytlsJJNGhUkkkxjskSmWLUVhcpForGgJCyNFkiyL8NeLrxbs2yWyWZMYhKkSqUFaIwQAY0agoAoVSwWpqhYWiTZCKqJE0kERmq2o0liAQDGSyksgEqmo2wmorStGrRbTVNtdRpKWqnNQ8XrjYIUMuhJoWQsjcxsKcCQuXYBYI7JxbEgSNkbKqp3hajJxOJxNHYI7JkxpMsymkkaKjULksRFFTCxWBGwACyNGbfi6cNGVs1MmTGEJYqCVLELGCACUlBCVEVrILC2IKkSg0ckhhKwlSoQgBWShjlBqlFQmorCWu2DVC0W5qpqm2pYrb5vQVFI1NkQqSTRkbmOkUjcq5QagjRxMcSsnEhIrI2V6zXVe4rGlmVsTiVhYydSSROSySckh00LCwSNKkjEIVhZGoUkVSW7LTlo53Rm3ZtsWkiROBUsViqIipIlSxgGJQiBEiRWDULYCQGjWUMcrhgMQWkOGCitQcoITUFiRtjLBqK1rTLRbmWC5OfV0IqEmMVzNLIdkUjqRuY2QsVCNmSzSeRSqcjkYWQsrshqV6kNSySyWySwkkg1HUokzOSRMkjiQ0VisQUqSRELQFZCyFIKcWRozdWGnK/K3NsW2JxJWKVNRqBBYiEsSIlQANAUsLUsSJEhSSSyiSygGISgDiU1IasAVEVTUSCwWKxE1AUta1zUTzzpdKDRhZJlgkyaMhqK4RGo3KsaSWSTyBVFJFkhRcxsrshUNZWpOLJJlsSHYU7mUkiQJIkSJQ4ELBFSqNkRC0jUbIrG5VIZIsi7LTloxb8r827NsJjlYxSxaiQI2xljUVgIBAABCpCIqhArJBKhVFSGOWUslkSAaoTUVgQWDUCIrqERthLBQ8s3fFs0MunI6aSGkySKyFyhWRuSwJLKSQSKxK0kWIIrI1CyNDJqFkkkWLNJDuRGNJAOpQyUTkcCKxaIVJEQ0jUaQrEKxwyRIsi7N0YaMW7K6W2WcSGpApKiNsIjUViJYiFSEAAIiCgwEsViRUlZKJK5WSGrBUJqJBa7YEVjdRiu2CqUs8jNXxfLYMkjAkkgR3LBIUkLlBqIlErHIABJJWThisEhRcqo6yqkkqsJySklYQ7GAwoJJKJJPJoqVKkFRIpCo6IKQhoxgiJl2V+bdm6ML82cNZyyWUNBSVEYjaiCxqJFY1EBKkSg0AItRIrFSUJEolElatQBKLEr1qJBYrG6iVlbSBPGW6c22LpbIklgBTSUy6YIkjYkKLIhY0kSydCMlZIlIUIIrFRYrFqOppJJxPMmhaIx2AUIE4lJOGAhUqBEajURWJREFhQCRtSSLJbsrs27C3NtiyJrOJKxw5VAqIELYld1AhbEgFEFOGIQhKKKxyuGSlaglQrqFsagsKhapYrCIWqo14ZdGWmLs2+LJZWTgR2McOxiSLJSCwFYkLHU4aSJEkY0YkLALHYD1DWZSSJRORjHQCAWNHTlkjykOABUKkKRGo1GwAQqVIVkRWtJxZm35XZtslmVss4sJStZRKEpEViV21NwqJEhURAMBCGBJWoSlcCqC1EWlUVhqxWFtZC2KxiFsbI6nhZq/N0Rqyui6JxMkMlI6ARoIkVKo2RsaRpsyJSSWZIkjBJIEqcy6LCx2KxjRk4kMaFgCFMcOnkyQ4YKAoEiqNipWKo0ESNkSNoKxCslE5bsrYuizNtyuymspZqQ4JUsZqFQK7qFIiRCkiVCACVSlkNXAJYqpYWwtjULpLArthUCJXbXqQrwy6M26TRm6crpb8pxJJklaNAEkNEKyKQ1FTRIWSGTiRIkjSRNJDSVNCxWJFSsZMkOSQ0AQGOnDABkpZQwQUkLRCkRsjoqjUSJEVRRUqiioJxZFst2Vubblbm2RYPNlNEqWIiJG2KxUEFIikLYhTJQyQ2kRIrFY2wWFVtQtiQqCw0gtVU1TueKzbpL5bo1YujK4uytykTBGSEkiaCpFcwqNisdgAxk4lZKSxJEkmTJI0BWRsVRpXLqY0llJZIkFaFBIIAGSiSuHBQNABVFI0tIUiNREiIWxuY1G1MlTlti7GrSzC3K7Ntic05VAsQAILQQEWlUREaiojUEqqKpURWFtawtrtrIbVrCoFNubSmz/xAAxEAACAQQCAQMDBAIDAAMBAQAAAQIDBBESBRMQFCAhBhUwIjFAUAcWIzJBFyQzYDT/2gAIAQEAAQUCr/MraP6qUfhFUl+8RC8Mf7iEL2MkTZVfzn5QvC/aiIZVfxcP9FZlVlQ/9pFqUCh4rFx/248siJH+z2wOtgnck7klWcj9xeEbJE7uMSvyequeXcipXnU8MfnHhmMtHC0u28Guy5rfCY0MY14n8ztUQ8VGN/MReZeYkfL8MmTKn7r90Lwv2oeJMrP4uJfoqsqEz/2mWpQKPit+1z/248sWRI/2Up4J1idclVyfuLzukTukivyCiVuTciVSUzJnw/cxIij6epfpONh28lV+ZsYxjMCrZqW1QjWJVypcHf8AMa53CrHcOsdoqhGqRqHYKZubDkTkTkTf6osQvH/lDxMrMuX+ibyVCZ/7SLUtyj4rFz+/Hlj+0SP9jJ4KtQqVT9/OR1UipdpFbkUityMpjm5GTI2ZM+5n/v7iRxFLqszgofoGMaGPxRuP1ULk9WTvSV4K6I3Qro9WO7PVkbr59UQuhXZ6oV0eq+fUkrknckrglcfMa4q5Gud53/FGt8dxOuVqyLut+idUnMlMz80i1ZQaKMj/AMrP4uv349liyBH+xqMrT8uSRK4SKt7gr8mkVeQlM3bEzJkz4z78iIlKLnKnHrp1m40rCn02AxjGPxSpsppowySZ1sVNipM62aMdNnXIUJGsiKkJyHKQqkhVpHcyVeRKuyVdnc8xrsVcjcHqfn1JSusL1hO7K12XNzmMqxKqSqG5Sngt6xRuPijcHqPivc/FxcfPGVclg8qBEX9eysytL5dRIncpFW9wXHJYKt9KobZExMX4MmTJkyIicVR7btlVbzximMYxj8RsBWR6I9Az7cLjj7efbj7cfbj7afbT7afbT7cPjj7cfbh8ePjh8aS4w+1sXFs+2M+2s+3s9AyNlIdpIlZyZUs55q2Ux2MyVjMlYzPRTI200U6U0Q3RSnJDrSxVrSKtR54mTzxr+IMgL+vkVSuXVx1uvyWCpdzqGTPhMQvwZMjZnxAgvjgKP/IWcO7kJ/MvDGMfhcehWCPQI9AegPQHoT0B6A9AegPQHoT0J6E9EeiPRHoT0J6A+3n24+3o+3I+3n25H24+3I+3D41D4tD4nI+IHw6JcKPhUfZT7Ofaj7bgfHlTjipxbb43jdSzoaKBEX9fNlQrnN1cTyZMiEIQvxPwiKKcThaWtqcJDap5Yxj8K1FanpUemR6ZHp0enR0I6EdCOlHSjpR0o6EdCPTo9Mj0yPSo9Kj0qPSnpD0iPSHoz0Z6Q9IekPSDtD0h6M9GOyPRHoj0Q7IdiOwHxxQtFAhEiiP9hUZUZcy+OVq9l5kyIQhC/ExsRAgR+XaQ67apLSnxVLpsPLGMfhIx4wYMecGDBgwYMGDBgwampqampoaGhoampqampqampqamhojQcDQ0FAURIX9hUZVZfT1jWnvWbMiERELzkz7mzJEh+0EWVLtuEXP6owjpQ9jGPwv5WfZgwYMGDBgwYMGpqampqamP7BlRlZnMV+uhkbMiERF+LI/ESmimsHCU97kpQ7r6f7+x+H4X8bJsOY6huKQn/Ax/ZSJsryPqOvinnyhERe3Jkz5fhEUU0RRwdLFI4iPZdv8Af3Pwv4WTJsbDmSq4J3SRO/SKd8mU6uxGX/8ABVGVGXEvjnq21y/CERIi9mTJn2PwiJAicdDrtG8Lhaelp7mP+Dk2NjYciVVIndJFS/SK3JpFzzCir36j63xv1F3y4683UJEX/wDwNRlZl3PEb+r2Xb8oiRF4z4z4z5yPxEgiBbx3qwWsbj/8beHXa/hX5XI3NzcdRIncqJVvkivyaRcc0olz9QIrc3OoVbmpVKqLS49PW4G/3ja1d4RYn/fsqMrM5StpRlPaTkZMkSIjPjI35z7URIEDh4b3aNey6n8fhx+VslIlUwSukipfFXkUV+WSLvndS45qpUKlzUqGTJkqftL4f09yGs+Lut4xkRf9/IqMrs+oq+lt4QiJERkyZMmTJn2oiRF+3A0/hHGR7eRk8y978LwvwskyTL6661X5hIuOdRU5ecydzOZNjH7JMqos7j09bgOR3jbVd4xYn/fTKjLmXx9SV8zz4REiLzkyZM+1iEREROJp9dmcDH/j/IvwsmVf25bZxvN1UbNjIzA15YyZL4Pp3kdJcVd7xi/iL/vqjKrLuX6eZq9t94iRF7MmfGTPnPhCICKSzKhDrp13ijY0+mx/GvxzROJf2+Tl7InHBgSMDQ0NeGMaJxLat0Vfp7kd42tXeEWJ/wB7VZWZyNTWFefZWz4REj4z4yZMmfahCIkTi6XZdorfqm1pT/GvxyJl3DK5O32V5Q0m4+MDJD8MYySJLD+nr7qqcTebwjIi/wC8ZUZWkc9X67ZsyIiIRn4yZMmTPnI38CIkREDgKeaiLOPdyVV5qfjX42TK0cxvaWTkbYnTHHHhj/dowSGMZURbVeitwHIZjaVt4KQn/dzZUZcSPqa4/Q/CIkRfgXlfuREJEPg4OlpanBR3q5y/xr8kiSLuBfUMlzQ1dSA14aGvDGMZIZwHIdc+Ivd4qRF/3dRlVl3PC+obje838IiIyZMmTJkyLz/5ERERFkFksqfXb1ZddLiafTx/8hjLmGS5pF7blSGCcfDMEkSRIkMkSKFbpq/T/I7K0r9kIsi/7llVlZnI1MR5K47L1VCMyEyLEZ8ZMmfK8ZEIRERFLFjT7LqJc/qgo9Vr/HYxlSOVcUi7o5Lylq6iH5ZJEkSJDJImcDfdVTh73aMZZUX/AHMiqyszm6ulCtTcpOLRGRCZCoRkZ9qEIz4QvERfvA4GnvdIhDuv67/X/IYxor0y4pF/QK0cOXskiSJEhjJoo1HSqfT3IZVlX7IRYn/cSKjKzPqCf/FOgVKBOkfMSFQhUI1Dcz7V5iRF4iQZ9O0sUDhY9t/OW0v5LGTjlXFMvaOVfUsOXskSJkh+JIf78Je9U+Hvto057KLE/wC3qMqsrv45j9dWdIqUidIlRJU8HyiNQjUFUNzYz4z4REQvERfJxdLpsnJQjwUHSs/5TGNFenkuqZyFEqrWXjJImSGS8MmilUdOfA8jlWFxvCMiL/t6jKrLn/reLetOmTpE6RKkTpEqJKkf9TtwKsRqEZil5iLwmfuRLWG9enHSN28UKEOnj/5TGMnHKuqZfUS/o6tsyZGybJDJeZjOGvOqrw19sqdTZQl/bMqMqMvJYjNZcoEqZKmSpEqRKkTpE6ZKiNOIqmCFYhUFPJkgLyiBwFLsvkTXZd3H6f5jGNFxTyXlI5CgV46TyZGxsmxsY2ZGyZSnpPgeQOPud4wZF/2sioybOQn+hjGhwJQJUyVInSJ0SVEqUSdE1cX2YIViNXJSqfpjLxkgyP7fS9L4RxdPu5as9qv8tjGTWVdUi+ofHI0MGTI2SZJjGzIxjGcTd9NXhb3KpVNowkZ/tJsqMqM5CeWx+XEcB0yVIlSJUSdEnQJUCpSP+pCZSq/EKhGeTJAicFR6uPR9OL/j/msZIuKeVd0TkrcuIddTI2SZJjY/LGMhPSXA8gcddbxhIjL+0qMqMrS+LuW1Uft1HAlTJUiVInSJUipRKlElTwRnq6dchWFUKVT4o/rqW1PqoXM9Lfjqfp+M/msYyayXVIv6ByduMbGxjGP2Pxxd101eGvsqhV3jCQn/AGdVlVlzLCm8y9+BxHAlTJUyVInSKlIrxJkWRqOKhX+aVU+n4epv0XH65VF10f5zGSRXhlXtE5O3Lun1zbGxsbGP2sjLV8HfnGXW0YSIS/smVWVGXs8R85/BJGuSUCpTK0S5GiFM6/hxwU6ji/oqjmaLSn6jlbmWa385jGSRd0jkKBydsS+GZGMfl+zjbnpq8LfZVvV3jTkRf9jMqMqMv5+X4XuwTI/tIqFwysdOSFuOh8VKJSo71Ppi26LFH07De6lPaWxsbGxsbGxsbmxubm/8JjGSK8Mq+onJW+VfUdJ5G/D97IvD4S+OLuto05EH/Y1GVGVX8Xk81M/jmRfxOZVqFepkl8lOnkhRJ0iVIpUsS42l02VWp1UeGh6XiNzsNzsOw7DsOw7DsOw3Ow7P4bGMki7pZV/b/HJ2pUWkn4Yx+1+OPuOqrwt78WtbeFOQn/YVWVWXEvitLaf45MnPBVqlasTqZdNbOhSI0xwHSLC27rtLBev/AI7lq34/sOw7DsOw7DsOw7DsOw7DsOz+IxkkVoZV9QOStsnIUNZeH+GLw+Fvji7vZUpkJC/rmVWVWXlTEW8/iySmVKpVrFWuVqxDMnb0ijAURxNT6ft9r417+S5uulddp2nadp2nYdh2HYdh2HYdh2fxGMkSRd0srkLc5O2KsOuY/LH7rGv01eGvi0rbwpyIv+ukVGVZHI1fjc2NjJkz7MkplSqVKxWrlauLNR29EoQILy0fTtLWmcHDu5W+uO673Nzc3NzY2NjY2NjJkz/FYySKscq/oHJW+TkbfXy/wo4a9OKvNlSmU5f102VWV5/HIVc1Nzc2NhSEzPhyJVCpVKtcq1irWEnN0KJRpFKJBGPPF0uqylLSPBt23FZNjYyZMmTP8xjGTRd0tlyFA5K3K8Ouf47Ov01eHvixr7xpSIv+tqMrMu54jdVdqu5ubCkJiZkcydUqVirXKtcnVyRg5FGgUqWCnAhEXmnDepTj107+etpff/U4DJkyZM+EL8WTJn25M+/Jkz4YySKscrkKByVucjQGP8fEXmr4i8yqFTKhIyZMmTJkyZM+MmTP89lVldnI1dYynmWxsJiEZHMnVKlYq1yrWJz2KdLJSolOkQgQiRM+eJp9t6XP/Jc/UtTWZn2IQhe/Jk2NjY2NjY2Njc3Nzc2MmRsyZM+GMmi7pbLkLY5G3Lmn1z/Hb1Oqpw18WF1tGnMjM2NjY2NjY2NjY2NjYyZMmTJn8WTJkyZ92TJnzIqsryOWq/p8IQjI5kqpUrlSsVaxlzdOjkpUSnSIQIxIr28DT/UcbT9RzfNVu7kvYhCF+FmTY7DtO07DsOw7TtOw7Dc2MjY5imbGRjGV/wBuQRyMC+pDWPycVd6vjL4tbpSUKp2G5kybnYdhubm5ubm5ubGxsbGTJkyZNjY3Nzc3Ow7RVBTM+yUjcUhPxNlVlzP45WpmQkJeMjmTqk6pUqk6mTVyKVEp0SFMjAjES93DQ0sz6c/TTnJyefahC/DgaJImx1MCrHcdo62DvO47iNTJGQmbkpE6mDvRGtkjM2MkmXVTVXtbJevYvIFSn8uBqYMe3BgwUZunPjb3BZ8h8Ub9FO5UiNQVQ7B1B1RVTtO07zvPUCriqimbGxkyZ92DBgkydXB6r5hXyQqEZCfhsqSHMhMhLxUZWZdz/Tey3rJCXhyJTJ1CdUqVRyyRhkp0SnSIQIwIxEvGTPmPy7aHXQvajo2v/wDg+ms+1CEL8GpqaE4E6RVpj+GvE2ZHI7ShPYjI7CVdI793WuY0yd7CTd+6E6PJJpXsWeqRK5Rc1lNXlNsuYSLiLKkFmURwHA6zrZ1M0NDU0NDUtazg7W9kW93NlrcFO5PUnqiV2iV4K9PWfDvB3o78V+UrrYpTyQ+RRFE1NTU1NTBgwYMEkVEVycv1UZlJkCPhlUkymymf+VSvIv54U/1TSMkpEqhOqVKxOsbORTgU6ZCBCBGIkLznzksKfbdHIfrPqKp1WufahCEL36mpqOBOmVqZUp/OMDJjJzKdPc11Tr6k7zBcchgpX+tG55H5lyWZVLre1hzbiUOcyU+UcynXlMislaksXdJF3D5uIfODrydArY9KemPSHoz0SPQo9Cj0KI2SRRo4KSwUaupC5PU/FS7wVL0lekbtsV28O5kSuZnfNlKc3KzhIt4MpxFE1MGPwYJIqQK1IqW5Ro4KVMhAUTA0VkSXzRRTR/5WK7OSn+kciVQnVJ1SdUnUyKOxTolOiQpkIEYiXnJn28FT2uClD1PN/UVft5P2pCQhL8GPZNFZFRfMhywSkSbYqbzD9JUrfFW4K96oq+5VJ0eS2tq1faTrvadw1b06NWoWlrODsoZVKDQngqT+Lr5LqmXFMlT+Y0hUhUzrOs6zqOhnpmelkKykKwkR45kbCR6GZG2mhQaKlLI7WbPt0pFPjJEeMkfa2PiWxcOylxLi7ay1KVDBCkKBoampgY2bGxsbGxkaJUx2+SNvgjSFAwYGiqicSkiA/wBq7KzOSyyUWicsFSqVKxKsOWSFPJTokaZGBGBFCX4uCp6259OR7b64ffcaGhoaCiJCiJCRgwYMGDBr7ZFRFWJOBOI4iiYwTeCrNlxJl/XZOjORx8HB/apspcLKRcWMZKjxaRCyUSNLUpVGhSTJsuC5KsMnpm3C1ZGzZGyFYisCPHkePFx4rAjYEOPIceiPHxFx0T7ej7ch8ZE+1xFxaFxqQuPQrEVmeiR6NCtERoYI0xRMGDBgaJIY/ORPzqaGhj2SKiJwKcSJN/FeZP5K9Ddzsitx5cccytx8x2dREKDRTpkKYoiXiD/Bnwiwh12lar00eHpu0+n6PHOoR4eA+ERDjIQPttOR9npj4yCPtUBcdFD46LI8aj0ERccsvj1EfHo+2/p9DHzkySZNklkdMlQyO1R6bA7cnalSzLmzch8Lk+xxPsyKVnUokrapVUbHU9KdA6I4k6qgVbwdec3C0dQXEbH2UXD4FxQuNPt5HjxWArJCtD0grYjbkaJGkKkKkdJ0nUjpR0nSdR1HWdZoaiiY9mBxHEcBwNDQ185MmTJsbGxkZIcRR8VZfFxPw4jgiVBMqWakT4xMnxKZLiB8Vg9A4nppI6mjQ/Zp/HhsXsyW0O2ulqcpN+n1VOwt1AhbRkekPS7RVpKJGgypaOa9NUzG1yvQyRG1kiVjtL0ZK17IxslFRt9T0NMyORsbkpkpGfOpqaDgOlsO2R6WJ6aI6CHQR0odIlBFRqJUrE9pjt8npkUbVZo20SFJI6jqOo6jqOo6zQ0NDQURRIxFHw2OaOxCmjZGyNjYyZM/hwOI4mpqajQ/GTY2NzcUhSMjGjAyvIqvMvdg0Q6MWO2iOzRKxRPjsk+MY7OcB05L8HC097s6/U8jP9qdRqpa3ClFSFMzn8uw34Y/C8Y9uBoaGMY2TmVJNko5HEkMc0U6yTpXCIVsnadh2I7Edhubm5sZMmRMUhTHUKlYq3eBXpC8I3R6g9QeoO87juFVNzc3NzsOw7DsNzYyZM+H7H4yZNjc3NjYkypT2Ktu0848ZM+M+Mm+DYz5wh0oslawkT42DJ8SVOMqIlaVIji154CniBx9PatcyEhS6ijdkaqkKQpmfx7GxsZGYNTHjJn2MkyTHIkxyJTJTJzKlbBUuipe4KnII+44dLlsEOZSFzORcrkXJEL5shctkarFUFIT8oXjJOZXqFZzm1CoU9kRmztZ2sVVirCrCqiqirHcOqdp2nad53irHadh2nado6p3I7R1iVdDuUepR6kd2O9FekbtCrkamRftXSxdvWSrHajYyZMjJibIP8LimStqcyrxVGoVOCLC39Lb5wuMhinP9UqFFzlP/wDRoVdwKV3kjVTFMUxP8WTYyZ9r8ZMmxsNjySySUiUZEoSJU5E6MypazZOwmyfFVJEuFqMXBTIcDIhwTIcIR4fBHi0iPH4IWuBURUxRMGPZkciTJwydJ0nSdR1HWaM1ZhmBZHUwSusHrRXR6kldjvGK8I3Z6wlfHrxX560lejv2fcD12SV2yVzI9VIV2z1J2tnbJCumiF6Ur8p8hAq3kZKr/wAjnRJRaOxoVwK4FUTNvGojPjPsyZ91X5jb2XXT9EW8FRXpk5SoLWVuToSTjVnTKd4QuExTFUFM2M+31A7g9QKsdp3Hadp2HYbm4nk1NDUcRwHTQ6Q6R0joI9Oj0yPTI9Mj06OlHUjrR1o0RoNGTY3Ow7UdyO5Hed53DqGxn8WCcSvElLQjXybjmSmx1GKs0KrsSTY8mzFVwbqQ4jQ5NCqikmaZHAawb4FUR8SHA+YnqJRPXtEORTI3kJG0JDoKRO0JUJIzKArpxIXiZGumKRkyZMmfGfOfZvtyFrWczUcBwGn4xFjoxZK0FQlAUpRI1iNYVUUzc2MmSJg185Njc3OwdUVb5pzFM2HI2NjY2NjZGyNkbo3R2IdVDrI70d6PUHqD1BVusFTkGn65s9XI9TI7pCqMUmLIkzDNWaM62dbOs6zrOo6zrOscSS8SpJla1KlBxISNTQdElRHFxI1cCamOkSpNHyiNbUjUUhwTJ0hpxFVaI3B+mZKhknQaHmIqzRGspElFlSkTpuIpTi/uE4Ojymqo8tGbVxSmnQp1Cpx5UsZIdOpTFdTgU+QRG5jIU8mTJkz4z7eL/wDsfUFpD9a84NB0hwaMtHYfpZ1pnW0fKNmKqKuKsdwvOxt5wYHEcDQhLUVU7h1jtOxnYzdnYzdm7N2OTNmZY2z59reCvIkstISKcExUkRpoVNCpoVMVNGiNUYMezJk2NjI0Y8YyVLdSK1rgUnTcJKZqjVEqUWVqCHJ03TucmUyUFIlbkoSiKvKBCupCipE7ZMlbtH64ON5qRuYTNYzJ28ZDtZRX/JGU7vrO1TKjQ4FZ5VHMF9wmp0OUnShb8/8ANPkaVRa0qxU46MirxJPj61M7q9B0+UKd7CYppmTJkQhePpOG6tFiKfuwOCY6I6R8o3aOw+GY8ZZsNkX4Yv4uDBqamhoajiSRcfBusxGKvq4XJTr5IVCMjY2NjYyZ/ExoT8OKZWs1IuKFWgU+Ree/ZOrI3kzRSJ22TsnQdO7jUMjwydCMipQkhVqtuU+QjUFLYaiydsmVLacSFSvSKXIxlLtUhqBXt6UivQ/RTtq6dW6Sn+8ZS+MR2r1VNWjzOrff8nH3rI/Uk6crTn6Vcp17euTs1IrcTSqE+CaJWl5bujVvEUVVqDoyiJYF45Or0WHBUPT8VSk4KNQT/DqmOkOkdZqz59kV4YvZkyZ9mfwsyKRt4x7MEol3TzGo5xrUP+uCpD5pxKaIkWZMmTJkybGxsbGxsbGxkyZMeydNSLziIVyrC54+VveU68dVis3FxsrhxpWGat3xNGUf/tWjoXMbhdcjVkqO5OwpSj0XtvKnyMHU7N1KSy+uRO2tpqdC42tJ3cydObJUHFVpMqU6Mp3E06Ktrnot9K9apbKQuPrM+z3EyPFX0KNX6au6UafCXkreH0/dp8Xb3dnLvJSyLwoyFTmyUC8nStqNDlres+SqRvL2lYPq2nSdOtGXhSaFUM5/Dg0Q6Z1nWIT8syZM/jyZMmfOPKEajXnJUhsqtoiMcCQ4ZI08EVgTNjY2NjYz/CaKtGNSPI/TMSny1W1qqpKsrtXVe3ha3NAdvyM1T4/knX5L6VhWJXl9w1S35a0uYS5K3R9wyK7rlWynf17rgeQ4+cLDlHSXE8mz7Hfs/wBfuWf6y2f6vQF9M2iF9O2RH6ftRcDTKfFVKcFwtbMOCuJFH6eSJcPS63wP6qvE29tSpfbahGyt5P7dSRTo0vUekpROW52HEXD+srg4/nry9rXFKpMlx9xJ8hwd3c23+hX7f099GXHFXVpFxpSpqRK0izplEWfH7CkKRn8ehr41HAlDBnBkbFIZkyZMmfGDBg1NTBgwamPCkKZ8MwOmOmP4K80vCQkJCX8txL3jaV/Rt/pCvbXL+nNnU+kqFaFraq1t8Ci2TspV40uEjSUeLqI+21j7XVFxMj7Sfaoi4ukfbaCPQ26PSUUK3pI64I+qP8jUfpnk/wD5inUfH/WPOcq7N8jVVOjIUceZ0diNosKkkYQ0NI1izWJrE1gfp9sf38YNTQ0ND9jY2/Bnwn8C8ZJSwSaZnB+5j5H4wYGh/BuzsNkbG5ubG45CZnzjxlnd8uaZyH/50/1RSEL+GotnRUYrSuxcfcMXF3B9prEeHq5XDsXERFxNI+2UELjrcVjQQreijrghJIyZ/Lgv/pOx5S+tPprj7MhQhD3ueDtNjL849+DH4HEcDU/Y2NjYz5yZIyMiZsb/AC5JlX9J2aELhMTTPj2bDmNmxsS+B3GhGopj8fJgUjKfnLNjI0mSpouKGadlnqx4yZMieSNKciNnXZHjrli4m6YuFuWLgqwuBmLgRcHTFw1BC4q2QuNtkKyt0K3pISS9m6RK5pRHydpE+82CHz3HxH9ScYh/VHFof1bxSH9Z8Uh/W/GIl9e8ciX+QrEl/kS1H/kWiP8AyNAf+SCX+SKg/wDJFyP/ACReE/8AIl+yX+QuQJf5C5Al/kC/H9e3pQ+r+TvZcfZ89fFjw1WkqduoGMfhkiWYtVDb8mDB8oz79TQ0NTBj2aD+BedYjhkqUlMVCNPw4tjdSBGsbG2DOR48TlFkosqUnJQbpON5Kmo3tOYnstjJk2wdh2m6fj5RuVZovLydjcw+qLeBR+qeMZb/AFPwZH6r4BH+5cJE/wB54tD+veOQ/wDINiP/ACDaj/yDSH/kIn/kOSJf5Grn/wAiXJL/ACJdD/yFdkv8g3ZP6/umS+va5L67qkvriRP63JfWSZL6spsf1TRH9UUh/VER/U4/qeY/qWqf7JWH9RXB/sFyPnbpj5m6Y+Vumfcbolf3Uz1FzI3umaXbFa3kz0F4z7TesXB30lT+nL+s+P8A8a8rePiP8UWluWPAWnHwjCMfx5wbolHY6zGDIpGTP48GPxYNTU1NTu1NlIcEzTU3G8pywb5Gng1M4HiopOVMVz8OvgdyOpMVbtKlVpwl/wANWadOmuw4zg01HibSDdhQPt9uehoHoqAraijlatO1pVKpC5IXWx3ol+3L0JVS5g4uqsk4jVQ3rI77g77g77k7rk3uWf8A2mdF0xWV1I+2XguLu5H2e8Z9jvmfY71n2G+Qvp2+Yvpq/Z/q99mH0hcyS+ia4vomoL6JyL6HWI/RECP0NSxH6JoIh9HWqcPo+zP9SthfSlqL6Vtkf6zDH2DUjw0RcJJC4OZ9imfY6pDg6hR+lJydr9I0Yu34mhbKNOMfx7DqE6+CpdM76zlQqNowOJjxkyZM/wAhs/ZwqGUyR/1FPI2Ko03+okSIS/VNZJyw+zDjW/VUnoOsOuqsaV46Uq0J1zg+C9JD3Xl7C0hWdStUlhHVGRKm8qEsqGCdDcqcXGoVfp9TP9ZWV9LwkL6ZilS+no5/1e3ml9MJEeAnFQ4jBHiKZPg4TU+AiUuIij7JSmfY0j7JA+yxZ9mhn7JE+0Qx9ngfZqR9liiNhqehpnoooVnEVms+nR6dHSjoSNTriRjgxFnXE9MmKzqEOPyQ4+KI20Yigl+LZIdZIneRiS5SBG7dQjmR1HSjpQoJGPZgwY92TJkyZ/gZGk1/0IVVI2G0z/qb5NjfV/8AaLm03LLjPKuYtk/iUpoclUjJLGUqlGzq3tbh+Ijx9L3X17Cyo8hyEuQq/c7u0LL6io1yEqVYdCY4s0Zoz5EReRRWcI60KJgTNj9MjriaHUKmxRMDpn6kaZNWas0Z1yOuZ1zOmR0TPT1D09Q9JMdnPGridTZ6WTFZSZGwRG0ihUoo1S/Fsh1UiV0kVuThAuObwnzNaq51qlUsLIpUdBGxubmxsZ9uDBgx78mTJkz+SE0KR8MkknGeTJkf6X8MaTSZOHxKPxCeCbzGsZ+O3R5lOVpxFW+qcZxdOwpe6pLWPKWt1f1KnBXBV+nLtqH01cnFcbWtbuFGLj6aB6WmejpnoaZ9vpH2+kehpnoqZ6OmekpnpYHpoHp4Hp4HRA6YnVE64mkTSJpE0iaRNUaowjVGDHvdGLapJGq/Fk3Q66RO8jErcvSplXnolXnJyKnISmTvMFXkPiPI4lYVYVFQu6cFTuYyFJPxjxsKRsbGTJn3Y9uDHtyZ/C/0kaopZJLJ+wpjZsbnYfv4lDBMdVoqR3Vai0U7Sdd8bwBb2kaEfcjGTridUTpgdMDogft/DwY/i5NkdiO5DuYolexRV5WnAq8/TRV+oJsny1zUJVq9Q6pHRIdsz0rPRZJcdsVOHFb3FCPfeUpW/N16RbfURQ5unMp38JirJm6F5ybGxsbGTP8AByZMmfLQ4shNoU8n7n7GfH7GTfB2jnkwVIjlqbKT42MIqFeEF6tHqkepieoR3o74iqo7EdiN0diN0diN0bo3R2I7EdqO5HcjtR2I3RujdHYjsR2o7UdyO47juO47TtR2I7Ebo3RudiOxHYjsR2HYjtO47z1KHdDvEiXIRRLk0T5VkuTrMlfXMjtuJDpzmejFZI9Gj0iFao9OjoR0o6B0DoOg9Oh2cWT4unInxCRPj5wIuvSKXKVqZS5wpcvCRC+jIjXTOxGV4yZNjY2NjJn8GRyNjYz7simmNjFMjUN8mTYbMmTJubknkqInmIr24pqXLX0T77fRI/UF4f7DdkPqO5z/ALFXR/s9WJH6twL6wpH+4UhfVVNkvq2mhfV1Nn+20iH1ZSkf7TTP9qgL6jUlL6hkR56ch81UFzVUjzNQhy7Ici2K9bPVs9VI9XI9VM9TM9RM75ndM7pHbI7pHdM75HqJHqZHqJHfI75DuJnqpkbqR3yO6R2yHJseRxNDrRpE0idcTridSOs0Nfbjzkz5x7HFMlQjInYxZOwHauInVplPkKkCjyxT5KMiF5Fka6Z2GxkyZMmTY2NjJkTMjkSmdh2HYKYpe3JGeRsbFI2E8+M4GxsyZMjY/klHJJaknkmhZT/c+RPKaeZ0skqQolNFSBpgcSH74ynDDoMwQ+Bs2FI3IXEokLwhc5FVyZ8Y9mfGTPjPnImNEkJCXjPtwY85MmTJkZk3Ow7Dcz7MmffgcEyVvFkrQdtg63EjUnAheyiQ5AhfJkbpMVcVU3NjYyZFI2ExyJzwV7pQUbxSFXQqyI1MkZEX7GR/cYngTFLBnJIbGzJsbGfDJxyVKZgcfGPH7iROjklSIrV67LrOs0wUl8aZFHVx+V5TwLxkhNohXaIXBGqKZn35M+M+MmRiMjkbCkZMmTJkyMch1DtO47jtHVHVHVO4jVYpmxn8GPOPGBxNB0UyVuOg0ayRGrKJG8aIXxG7Fcirncdgpm4pEpE5F5T3OrA9kd04lvctlKrki/GTYyZExj8x8SiSXhmfORkok6Y0a+xPJgcB0yCNTUdMpxw1EcSETU0yOBgj4wIQiNRohXI1cm5ubGxkz4yZMmfGfY/GTc7DsNx1B1DsGx/JodQqR1DpHSdB0CpCga+MifsyZMmfZgx5wajgOkh0h08H6kKtKJG7ZG7I3OSNcVUVQcyUifyYGjVMpwRSKbNhv2Lw0Y8RYmMkiS8fsZMmwpDGskqZqOJjzGRkwYM+MGpEwJCMGpoaiRgwL2Ko0Kudx3CqiqCYjH42SZsKZ2DqG7FITEKIqZ1nWdR1nUdZoamDBjzn2ZMmxsZ/BgwNDgdY6Y44IvBCrgjXFXFVyN5MDiamCCKZFmxsbeULw/GRSNhjJDMmTYyJ+ZoY/ka8xYjU1NRISNTBgx5wYMeV4fsybCqMhMhIjIT8YMGPcxjQ4mhodZoaCiJETJn34MGBrxgwYMe7Jkz+DBgwYHEaM4O3B3kbkjciuBVUzZPwiLNzc3N/OfGTI/GTfBvkYyXuRjJqSpjiYHESFEj8CMGpqJCRgwYMfmx4UsEKpCoKYpikZ8YMe3BqaGhoaHWdZoY/K/bgwYMGPZkybGxsZM+7BKJKJJDMimKqKsKuRuBXAq53ncdx2mfGTJkz7UzIxmDU1NfbqOA4GpoJYMCF4x+PHnHtwYMGBEZtEaxGqKoRmKQpGfx5M/hyZ9jGzYyZ85M+cDRj2ZMmTJn3uJKmSpDpGhjxsbnaKsKud56g9R70/Y0ZMjGvZjyjBqaGhqOJgQhezJkz78GDBgwYMeMGDBjypYI1CNYVQUxTNzYz+fJkyZMmTJsNjYzJkybGxkybGxkz7MflwOA6Y6Q6Q4DXnY3Nzs/Jj8aMGDBgcRx9mfGfYn5Xtx4wYMGDBgwYMezIqjQqwqoqgpmxsbGfZgwY92TJkz4yZ8v2ZMmfGTJkz/GwOI6ZKkOmOI0NGPwZ9uPx59uBx9if4c/xVIVQjWFVFUFMUhSNjPswY9uTJkyZMmfw5MmTJkz4yZ/jNDiSpjpnWZM+9fwM+5mPCEZ/nZFUFVI1CNQUxSFIUjJnzgwY/Fkz7cGDBjxn35Mmf4uBP8Of4GfY0Pw1+bJkyZ/hqQqhGqKoKoKZsZNjJn2YMGP4ODHuz7Mmf4D8pif83BgwYMezP48/jx7ce1MUxTFUFUFMUzYyZMmfZgwYMGDH8LP8hMTE/wCbgwY8Y9i9uDBj8ODHjBgwYMGPODHuUhTFMVQVQUzY2MmTPswYMGDBgwY9+f58WJiYn/RY/h4MGPGDBgwamDBgwY84MeMikKYpimKZsZNjJn24MGDBgf4MmTP8XJnznxFikRkJif8APx4wYMfxMecGDBgwYMeMe/JsKYpimKYpikbGTJn3sx7cfyc+cmTImJkWJmRPxn8GP4uDHjH8rBgwYMGDBj8GRSFMUxTFMUhSMmTJn3YMGPZgx/ByZMmfZkbGzIpCYpEZCkJifjP8NfiwY/n4MfkybG4pimKYpm4pGTPvwYMGDBgwYMGPyvxkyZMmTI2ORt4ixEWIQmIX43+df2aYmZMmRMTM/kx/Cfsz4ZkkzJ//xAAuEQACAgEDAgcAAgIDAQEBAAAAAQIREgMQEyAwBBQhMUBQUSJBBWAjYXAycYD/2gAIAQMBAT8Bn7kEeyGahISF0LsIQyT62Ikyb9CRIe0DSIEdmam3iCQ/s6Fp2R0SOiR00j2HsymyPh5SIeEivcUYw9i931IRqOovaTxg2afrchCFstn6sgPabGIWz2XYQhkuzJmoSGPaBpkCOzNTbxHsSH9ko2R0iOkLTrpxbI6NkdJI9F7DfbSEjxD9EtvGSx0TT9IiEIQizO2achzJTJzMhSMjIyMhSMjIsvaihI9hsb9ezM1BkhiIGmQI7M1Ntf2JD+xSNOBDTPboUGxaYopF/neW3iHcq2/yMvVQ2QhdENT1Iao9UlrEtQUxTMzMcxagtQ5BTOQzFqC1BTMx6g9QcxTMxTMzIUhscicjUl6DY2NiIGmQE9mam2t7Eh/YwRpR3oUBRKov83vtoQtpPKTYvc8TLPXeyEIQtoxEMkjAUDAwMRxZgxRZTKkep6ibFJimzNjmxzM2KZmKZyHILUHqD1CeoTkOQ2NiZBmnIjIjIyJTNSYmansSJfY6aNNFCQkexfwEIRqvGDYi8U2L1k3uhCFstIwMDjOI4jjOM4zjOMwMDAwMDAwMTAcB6ZxHGjjMDAwMRxHAlAlpsekzhZws4mKDI2RbIyZmSmSkQZL2JIl9giBERaRfRe19xCEeLlUK28TPDRZD26ltnE5YnNE50c6PMI8yjzKPMo80eaPNHmjzR5k8yeZPMI8wjzCOdHMjmicsTlicqOSJnEziZRLR6GKMUYGBxo4zAxK2aJQEqJskS+wiQF8RCEI8XL+SW3+QlUVAW62Qtnrj1znZzs5mczOVnKzlZyM5DkZyM5GcrOVnKzmZzM52c7OdnOzzDPMM8weYPMHmDzB5g8weYPMHmDnOc5kcyOVHIjkRmZDkSY2P7CCICH79V91ERba0stRsXqzxks9aulC+DZZZZZZZZZkZGRkZGRkWWWWWWWZGRkZGRkzIyGxv7GBAj8NCIiJvFXtH09STy1G/o66rLLLLLLLLLMjIyL+xRFERegn8C9kRIiPEyrT21ZYaTZD27L+NRQomBiOI1/oUSKEf12r7KIi28XL2jt4+WOkoi618WijEURQsjoNkfCk/CtE4USX+hRIiJd6+hCRES28Q71GI8fK9RQ+TRRRQoigyOg2Q8KR8Ol7lRiT11EhrKZ4jSTVokhr/AEGAhH99hvtRIkRDJO3ZH3NWWes38ejExMTBkdFsh4VkfDJe5UIEtaKH4i/YlNskQngyLzia8MWNDX0F92yy/gIiIiLrb676IkSIjWdQYxvGDkR9bfxkiMRQsjoNkPCkfCpC04RJaiiT13/Q5yfQxnhtS/Q8RDJWNDX2Vl9xCEf0L4KIkUIR4l+iQzxksdCiKpfGiRR4fRv1ZUYj1Uh6w5tjY+lk/U0p4SItSRrQxZJDX2l9qItpfCREiLbxD/nt/kZesYfHiQNL/wCPQ1L2vofQyQzw2pfoasc0SQ19tfYiIQ/fddD6nte8SJEQybt2L3PEyz1/jxZFnhtT+jVjY+mt3sySISwdmnLJGtCmNDX199i+qIhd99EUREI1HUWxnsmxfylKXx0RNCVMX8ok410vqkM8Nqf0Tjkhoa+Xfxr7iF33s9kQIoQjXf8AHbxEsNFs01/HuPtoiabpmhK0TQ0Pd7sYxkkQeMiErRrQp2NDXyL+hvdCH7d2+hCIkRCNd+tbf5CX8FAqvkITPDyPckhrtMZ4bU/pkv5IaJL7tCEN+vavZvpREiIW2q7k9vGPPXUfz5KEaEqNOVokhofQx7MYxkZYuzTlkicf7Ghrvv599d9KFs5eopFl72WJ7WWN9KIkRCG6GIvPVlL5KEQdM0pnuSQ+pjGMYzw+p/R7oaJL6q+7eyF0piZZfVY30IiREIRqP+IzUlhpuRor+N/JWyNKZpyskPrYx7Mi8XZpTtEkNDX1t9xbPosTMiy+yiJEQttV7eOljpURVRS+UhEXRozPdE11MYx7MZoT/o90NDX3C3fVZZYmWWXuxESIhC21PV7eOeWpGHzEI0pUabtEkPpYxj3ZF07NKdoY0Nd9/UIXYrfITL6URIi2W0vXabz8Q3+fNRF0zQme5NdL3e7GaE/6E7GP7K+lduitsjITL2REQhCGMk6Vmj/K5fvzUI0ZEHaJol6dLGPdjE6ZpTtDGP7Zd1oa2sTEyLIiEIRL2GeLnjpM0lUF81CIujQmM1I9DHs93vozp0Rd7P7Vd6ijEa2TIsixCFtN7ePd4w+gRoyoi7JokqfQx9a9DTne76r6H9UhfAaHEoiQZFiYntN7a7z8RX585CERZozGakd2x7Pr0Z0yMtn0vt39BfQhfBaGWQYpEZCYn6khsg89SU/noQjSlTE7RNGp6F9D6mJmlO1s+l/S32lsvgMltFCEyMyL9BmtLCDZoKofQIRFmlKxmpEfp22aUqZF7P6W+8tl8BjEhIraPuf1t42VadfpGNKjEoxMTExMTExMTExMfhoQjTlRdomjVjtfZYnTNKez+wvZC2XwqFvE/oZrf8mvCBiYGJgYGBgYGJgYGJgYfDQhEWacrQzUiNU+t7vbSn6kXu/sF8JssQuiHuMZ4dcviJy/DAwMDAwMDAwMDAwMDAw+IhETTlR7k0asf7630MXoac/QvZ/ZLu3tY2IihdENpOlZ/jIXDN/2YGBgcZxmBgYGBgYGBgYfEQhEWaciRqIkqfU92PbSlRF7P7Vdiyyyxs9yKF0x9tvEyrTZ4PTw0kuqjExKKKKKK+KhCZB7TVmrHu6crQn9LfeXZvosbPciuv8AoZqrPUhA0/b6RCIsgxk0SVPpfXpypkX9mu5Y2L1FEXUtmaCz8Vf4L26rLL7NFd6uhCEyD2kjVj3dOQtn0P6KxssvqXasssuxIS60RL2/x0cnLU/fgUUUUUUUYmJiYmJiYlFb0VuhEWJjJomqfcg6ZF7P51l9hssTF0rsXtZZ7iQl2Y7ajqDZ4GOOmvhUUYHGYGBgYGBgYGBiUUJCiOJRQhCEMkai7unIX0t7MkzIiyPSu2kJdv8ArbV/lUf00lUfg2WJkUKFj0jjMBaZxHGcY4DQ0YkYkYjgOA4mJQhDJEkNDRRRRW1FFFFEfQiy/lX3GajMjTZHpXbXYvdbNkFnrL/oXt8CyzIjIjMhMXqPaKKKMEakUiSMBaZikeuzY+ixjGNDHtW1bVtRWyYmX3rLL2ssv4MjUGzRI9pF7rux2Z4NZTcvg5FmQpEZmnMhP0LvaJZkclDlZihrb3KGyfsZmYmWXuyW1GJiYmJiYmBgYmJj3mxsyLL3ssssssssvtSNQbNFkO7QuhdmOzdep4KNRv4sWabIMiJFbMYixslMhK4jZZOX8WU2JMj0saKKK3reitn3mPtWWZGQpCfakahI0SHQutLsLsL22l7GisY72Wiy0WWWWWWWWWX0ogyEiMiLLH0yY0zSePozExJ+vojjMBRKKK2Yyiiiiiiit66rLG976WUV1V1IQuyzUQ4mmiHQumivgLZiVyijPFD1xeIHrnO0ebF4qzzQ/FM82x+MZ5qR5xnnG/YXjGed9aPNS3rZIiRdCmLVFrs5zlFqIuxmJiYlbUUUUUVvRiUUUUVtXU+iy9r7NFFFFFFFFFGJiUJCXTRRRQ0SiOBGBFC+StmR/wDqybZLUaOY56ZzxHqohrqJzQJa9P0PNRZLXTI+KpUeZFr4O0PxTbJa2R5rUKKMTEURRK3syMhSFqUcpyGRfTRRRRW77zH8WtqKKKEum+pocTHdfIjsvciY+hq6Y0OPexK2Qtn13st0UVvZZZZZZfeZRRRRiYlFFFFFdt9dlllmRkZFl9V/IXttH2NNf2SZ7k9K/YcGhoce5RRRQiyy9qK6UhISFvZZkZGRmZmZmZGRkWWXtZZZZZZe+JiYlGJRiYmJRiUV8ayyy+lfGroh6I1JEfYswUiegSg0OI4ldqiiiuzRQkLayyyxsb+MhMsyMjIssyLLLL6LL2vpoooooorrroor5WXqcxqPMWpSoU3ZHUFNMenGZPw34S0mhxHEcSuriFpHEPSOM4zjOMwMDExH6GRkWZGRZfTRRRRRRRRRXRfYsyMjIyMiyyyyyyyy+9fxr7S679RosyMi1t6oWo0LXHqxkNRY9MemPTMTEoooYmXvRRiYmAoIemTiOJQkUUUUUYmJiYsxMTEwMDAwMTAkiXoWXtZZe172XvZe9l/CfVXbr40X7sb9OnJi1BSTKTMT+SM2jM9GYj0x6RxM42Peit7LLFIyJKx6ZxHGYGCMEYoxRijFFFFFFdciXvvZZfRfwr7Fd++pFi2rsX2V/wDJPsKTFqGZ6MxRiepe1IoSJIXyL7Eia3f0VdVdLK7iGV3ZP02l6vt5CmLUMy10t7Lpr5TRND2fafwb3se99XqWiui2L09x7Xte3oMTLLPcsbLPTrsyoUhfynQ5UXffyZmZmYxrdFFFfGvq1V6fQWV+bKhsvZSKv2G6LL2Vip+/Xe3qep69PoWev4W/wtlstlt+xUhREWh+u0KiT99rL6a7mRe1ikKRVlCQ12bLLL3vpaPUsyL2mv4j+hfqtl6b0xJnqUyiiiiiuq2ymU+xfwrEy967T2Z77JCTW3t2b6q7PoVs/Y1PSXx6ZRRRRRRRXxq+fZZe1GJRXTJbNFGIk0R9dqPU9ezZlRkn3bLNdVLt2WWX3L3xf4Yy/DGX4YS/DCf4cc/w45nHI4pHFI4X+nC/04f+zh/7OFfpwr9OKJxQMIGGmY6Y3pr2Q6/PpLMjIsvpyF6j3tilQpdNl9N7+xbMy0+1RRqRs40PTHpy/o4tQ4dQ4NQ8vM8tM8tI8s/08t/2eW/7PLL9PLx/Tgh+nBp/pxaRx6Rjon/CXomWkZ6ZyQOWJzI5kcyOY5jmZys5WckjOZlMuZ/M/wCQrUMNQwmcczikcUjif6NJf2X+H/73b+dZZZkcd+xi0KTRlZij22qvVF9V9nV1/wCoHPP9PMan6eY1f08xqfpzan6cs/08Mp6jtv06pKxoaGhpjUhrUP8AkP8AkK1DHVMNU4tU4dQ4NQ8vqHl5nl5nlpnlpHlmeWf6eWOBHDE4YnFE4onFH8OOP4ca/DBfhiV2PQ9C0PUiS1vwcm+9e1bWX9DKBTEe5WzQtnu9r6ZP0s1NVy9F16enyMioxVLroxRgjBGCMUUj0P4lI9C0XFjoyRlEyRlEyRkjJGSMkZRLRkWZGRkZGRkWXt67WZHIZsvvKLZiV9LR6nuOJQl2X1SlSs1NVz64xzdEYqKotr2Fq/pmjNFosssvay2ZF9FsyMjMciyzJFoyLRkjJGSMomSM0ZxM4maM0ZrazNHIcjHJvv4ti02LSRikORfxV8Flbe/Q+h9D2vaclE1NRz7Gm4RRnEyRaZqSSj6GbORnLI5ZHNI5pnNI5pHNI5ZHLI5JHJIzZyMzZmzJmTMmWWyyyy+3fdsvZJsWmxaaFFLokevza7z7L3aLJ6qXsSm5dmyzIyZf1tl7U2LTkxaLFoowRRRRRRRQ4WcY4GBXzqKK673faZqNjKKKKK+HW1FFbUUUUUUUUVtXbplMxZgzikeXkeWYvCnlUeXiccUUV3qMRwOMwKfwF1Laiuqum9n2Xs4WPSRxHEcRxI4UPROA4ThOA4GcDOA4WcDOE4DgOE4TiRxHEjiRxI4kcKOFHDE4YnFE4onFE4onHE44nFE4kcSOJHEjiRxI4kcKOJHGjBGCMImMSkehZZZZZZfwb6KMRwMDAr4SYuivhX3K+FZZe9/Drv0UUUV8GjEcDjMDErutllmQn8h9596yyy/i0V2rLLL2ooor4NFGJgYGBiUVtRW1FFFFFFFdFFF9x9b7r667Nllll/Hrau7RXw6KMRwMDExMSiiiiiiiuhCX2lll/Aoor4Fl9NFFFFfCooooooxKKKKKKK2S7l/6K+3fcoorv0UUUYlFFFFFfWv5ll/Q1vRXcooooooxKMTExMSvurL+gr6Git6KKKK+4r5tbUUV2K+hr/TLL+LX/lN//wBCr/Y7L/8AXF/qV/6lf/i9/wDkn//EADARAAICAQMDAwQCAQUBAQEAAAABAhESAxATICEwMUBRBBQiUEFhYAUjMlJxcIFC/9oACAECAQE/AdP0NV9ib2gIY91ux9CEtokFu93syJp+pBGmiAiZrkx76ZEQhfs7HMlqDmN9TmkT+o/hFT1PUjpxj6bre+nU7tIRBWz6v8cNJfwh9UeyNVktorZ9C3Y+hboiMY92PaHqaS7kDTRDaZrkx76ZAQhe8r2zkOQ5Dl1OdEtb+EYyl/yFFL0K8bZ6z2+i0+TXij6qWWrIY+lwpGrEcBaZHSOMemcZgYC0zAwHAcTEooXYRQkRXYYx7v1HtBGj6kexAjtM1yQ99JEdkL9iyTHLqch6l+hg5f8AIUUvTavLp9++3+mRpy1H/CH37jH0z0+xqaRwEdAWicQ9E4ThOE4R6Jwj0R6Jwj0TiOIWkR0haQtPsPTHpj0zAwJxopsUCEDRh3IwIQIxKJmsTQ1vpkdkL9jIkx9DlRbl/wATi/7FV6FdNeFk3SIKklt9OuP6Nv52ez6JyJbJoyRkjMyRkjJGSHJFobifiJRMYmETBCghaaFpowHpj0x6ZxHET0jgI6JDRNLS7i0yMRRKJo1YE4kolCj3NPTKGIX7GRLduj8p+gtH/sV8eOuljJd2kIRrfhpaemMez6HI9SiiiiihxMTEooxMDAwMTExKEZGRfQ0ymLsRZpyoWohaiFqo5EOSJ0ycUT0ziIaZCCGiQhe8r2jJDQ+x+UvQjpJepRWzH4a2rdjI/lMR9Lp8mqon1UstVjGMfRZfVRRRRRRRRRW1dVl7X05CkZi1DkOQyG9khCZKQxC/YMe3q+pj8a2YyTNFerEf6ZC9XL4JSyblu9nvRiYmJiYmJiYmJiYGBiYGBgYGBgYGBgYGBgYGBgYGBgYmJRRXTZZkX+zY9o/O1bsfkoZIkPt3NJVFbfRLD6aU/nZ7PZ+KiiiiiiijExMTExMTAwMDAwMDAxMSiiiiiiiiiiiiiiv2ctn6CXbdjGPxrZjJEvgWzWH00IfPSxj93aMiyiiit6KKKKKKKKKKK/ZPaRQkUMY/MyRIirmhGlDOaifVP86+Ol+6yMx6g5mQpCf+BMey7srd7PqrwMkM0VbbEf6bDLXT+DUllNy637SzIcjIyHIczIyFIixP/AXvFbsYx+VjJDNFfjt/p8cNOer4H7CzIyMjIschzMtlGyGi2S0XH1PRkWRf+AvdLoYxj8jGMn6EVS2guP6SK+epj89jkZGRkZDmZHcxsjpNi0RQSIk45I1I0QYmJ/4A96KK2Y9q8jJDPWSQiKt0fU/jjp/HtmxssyMtqI6bZHR+SMEuiIjX0z0YmJ/q69oxiEu+7GMfXXUxjGRVzEfRQz1omvLLUb9qxjGzEWkxaYoi6kRNSOUTUhQmJi/YV5nsiK3Yxj3rxMYxml6Xt/p0cXLUf8L2zJDF6kKEV0LoQhH1GmNUxCf7KvI9kR3Yx+Ot2SGT7IgqQjRWH0l/PhflZJDQ0achdK6UIlHJGrAiJi/evZC3Yx9FFFdbGMZLvSEI1vw04Q8L3XjkMkRZFl9C6kxH1EP5JKmIX717JbsYx+Kt2MZIXee2hDPUjE+pllqv27GS2gy+lbLoRNZI1YCEy/3b2Qluxj8j2Yxku5p97e30Ef8Acyf8Ddu/cMYxMixPqXQhGvD+Sap7J/u3tFEY9jHZjGVtXRXR/Ixj2aJOjTVISPp/w0Jz+fbvZjJIRFiF0LZC2W0o2jVgVQhfuntAjHsOA4jiMe9bVu+h7PZvuT79hCJrDQhH3D2YxiZBi6UIWy31ofyakaEL909tNC7CdjRKJKI0V4mPZjJH/wDS2045NI+qf518e4YxjQ1tFi6kIW6GrRrQGqF+7RpLez1JQJQHEorpe7HuySI/8nt9HG9VGo8pN+5YxjQyLIPqQhdOrH+TUjsn+6RpenTeziOI4mJXSx7PZj7EPTb6b8Izn4X7FoYxkltBi79CELZboas1YElWy/bvaJD06XtZ6mA4DgOJW7Hs9mT9BET/AIfTpfPu2MY0MgyD3oWy2XTqx/k1YfuXtFC6mUegpF2OI4EojRRIe7GT/hbRPqe2MPj3bHs0SQiDI9+ldSH3NWBONbL9xBd/GnQplpmFkoDgTh3Gt5DJf8ttCOU4o1pZTb949mSGRZB9K8GojVj+4YiHXZfTdEZHqOJOBKJJFEhi7vb6X8ZOXx75jGSQiDIu/IzUiaka/cIiu3nUhSscbJaZLTHAnEn2TIoiu5D8dBv598xjJIZBmm/LqI1YjVft4i8V9SIjQ42S0ycDXVKhETV/GEYe/YxkkRZBkHfSuujUiakf28V7FCJSMi7JRTPqv+VbaccqXya7ufv2MY9oM05C8k1ZqxGq/bQXsUWNlkZEtSomp3ltoL8//CUrdl7WXvZZZZZZfsmPZjIsgzTfkZOJqR/aoiu3kQyuhj3n3Jd5EVbIuoSlvZZZZZZZZZZZfs2MY9oMhITvyTRqwJKn+ye0UJeWit2Me+o6W0DV/HSiiyyyyy9rLLLLLL9oxjGRZBmnLyMnE1Y/smIgvJRRRRQxj6Nd/jtFdj6p06+CzIyMjIyMiyzIyMjIyKK2r2D2Y9oMgxO/JNGpElGv2SILyUUVsxvq133raC7pH1ErfRZZZZZe1ll+2YxjIsgzTl5ZxNWI/wBikQRW9dVFdLYxj3ZqO5CPTJmp6/o3sx7QZBkXa8klZOJONfr3tBCXgoorpb2Y+iXoM013J9tMfr7Cyyyyyyyyyyy9297L2ezGQZBkH5ZI1Yj/AFFFFdKILw1tXk1XUdo+jZr9lXnosyMjIyMjIyMjIzMzMyMiy/AxEGRIu/JJE42TjX6hIoaH0I017fXeyXZI13e9eZ7WZGRkZGRkZGRkZFl7WWWXsx7wIEH5ZInAlH21+wRFGI4kuiJFdVdT8er3ls+zNR9/YtbMe1l7WWWXuixvoW7EiCIkRMTLL6r3Y0SgPTMemiiitqK9lRj1RIGJNEuiAvbMl3dkFbG/Ul6+wUTEwJQHEa8VliTZHSOE4TjaMSihIiLZbLayy/A0NEojiYmJiYGBgYGBxnGOA4lFFFFFFeFFD6YmmJGoS3Rpr2+o6jtp/wAsn2j7FaZxnGS0yemTj1VvRiKBpaRHSFpD0v5Ho2S0CWjQ41smRYhbWZFlllllllll7tFGIoCgYmKMSiikUiSQ0UUUUUUUUUUUUV0WN9MTTEahPdGmvb677Vt6QNX09jRW1GojUQ+tIURRIwNGFoUaEiVYmSG0arJPeImJlllll72WWWZGRmZozRkWKRmjNHIciOU5UcxzDnZfXRRRRRQ/EiAjUJj2RD2+t67T7UjV9dqKGiitqKKK2ooox3ss1DURJdFl7IRFCZp6qgzmiPXSJfUZMesPWZKdj3QmWZGRmZmZmchyHIchynKcpyGZmzNmbORmbMmZMva/DfgfiRBikTZMYtlIzFMv2k3ciKt0ajuTJd3tRRRgYmJgYHGYHGcZxnEYDZZkZE5E2NFGJiVvkRmchynKco9U5DkMzIyLFtnQ9U5jmOY5TlOU5TkMzMzZkZCe9FbUUYmJiUVvZZZfTZZZY2Nje1lllmRkZCkRmKY5Enst7LFM5DkM0ZFl+WfZbaS72Se0YWPSOLscTFpslpNnFIjotn20haEkP6Vs+2Z9u5KmfaC+maPs0OQ5jmPUJag5F9FFFFFdFllll75GQ5DkWX5ULeyyy/Z3s/DZYmKRkPZeGzIzMzkOQyRfg13+O14wYyzTkWKR26b8DmOQ2xjL2RW76H0X10YjiNFFFFeGtkyy+iyyyyyyyy9rLLLLLL66KKKMShxHErqXhorptmZyHIciMl0a771tqPskMSERmXZYpF7WWX15GRezKMSjuWX0sfUkUKAtMWkcY9MekcRxnGPTHEoraiiitqK3Ra67LLLL9nRRgcZgYldEUYmPUtn4smcjRzfJN5O9tT1GR3yojMssUjLx2X1PayyyxsvosTEzIUxaiFqo5UPVRynIchyGY37Siit8TAwMRQMDAxKMTExKKMDAwMTExMDEoowHpjgU9l2E9qMTEryV1Iq+5xkI4nH3MB6Yo1tZZYpCkZFll75mZmZmZmZmRkZFl7UUUUV1WWWWWWXvZYiijExKKKKKKKKKMSiiiiit62RRiVtRRRXRRXRW97UV0Yo4x6ZgymizIs7MwHAx2ry1UGxMorpoorezIyMhSMi+qyyyyyyxMssvx2WWWWWWWQ7kdM4zAxMSivFXRXSmKW9lmR6lHoWXtRW17UYnoWXtRWye1Dih6Y9IxaLaFMUy0zFMemOLK8ep20kvkS7ldFFbWWdiiit8jMy6L8NmRkWWWWWX5tMju/b2KXTYmeo475bVY1tkZbYmLPQyMjsYldFDQ9NMekh6bR3QtRi1RTTKTHpj02ivB9T6qJDwUUVtZZ23ord/oIEd6MRr26ZFp+o9P4MSiixSKUhx3sUikzArZSFJH4scCjuJsTLXVQxxseiiWi16DjKIpNC1WLW+TODGoEqRez20lc0a0rk2R7F+KjEreyy/0cWabHsmMfuYajQnGY4NbRM4jn27EdRn4yHHEssyoU2XGXqOHwV/R/8Ah3LkZL+R4iaE/wChf+Hcoz/gva18mUfkzgv5FqxOWI9aBqSjL0RXXBOToloyRH/bg5spyKFt3LL8VFFfpLNKe6Zfu7Ia38MwUu6O0RYoTi/UvSQ9TTrsQ169T8Z+hxyOKRx/2YR/7GUYLtIWtGS/Iz0jl0vg5tP/AKnPH/qfcf0fcy+D7iZ9xP5Oafyck/k7n5HcwkcY0xJsx/sr+yhxRRHSUldnDH/sThFK099Oai+59xA19Tk7I018jiOBTXVfjrorzV7HT9f0MZOLtH3H9H3H9H3Mvgbt3v3PzZjIwkccjjZx/wBnH/Zxo40YRMYmK+CkR0rVnF/ZKOnH1Z2f/FCiyt3GxaZgikNIaRSMYlIpFLw0YlFb37Fe6ooh6i9vTMZfBxy+DikcMjhZwv5OH+zhRxROOJxx+DGPwUvj2Ny/hlN+rFBLrbMjIvx0PqoxHErayy+iy97LL2vw0V0Y2NNeWxMi766MWYS+Din8HDM4JHAz7f8As4F8nBE4YnFA44fBhH4KXTZZaMkZIyRnEziZxM4mSMkZIyMjIyMi2W/guR+RUjGRkVOQoeJj7F+avBRRiYlFFDGumuj0LvfsUUV0Uf8Aomeo4Jj03/BVeNGi+xjfoYTIr5Fh8GemcsUcyOdHOc/9HP8A0c7+Dml8HLL4OSZnqGWqXqn+6Y6xhqnFqnDqHBM+3kfbyPt2fb/2fbn26+T7dHAjhicMTigccTCCKgf7Z/tmWmZ6Zy6ZzQOeBzRPuInLOXojFv8A5MSS8rK3svyUUV4KKKMTEwL8nodmYGJif+lUJD9ewvUaIaX/AGOKB9vp/B9vp/BwafwcOn8HFp/Br4QVJbUUVtpSoTLIsTifgVAqBUD8C9My0zk0/k5dP5ObT+Tn0/k+40/k+40/k+50/k+50z7qB91A+6gfdr4PvP6PvD7xn3cj7uR95I+7fyfdP5H9TL5PuP7Od/Jz/Jzo5onIcxzHMcxyN+h+b/gjpzfqLRX8iil47LLMjITKMSumy/JXtkIrZFDXYSsoqvQcbILvQor169TUwRKbk7e3cT3sWozmOc52c7OY55HOcpyM5JHJI5pD1GzkkZyZnIzkZyM5GcjKRlIuR3O53O5RRiUUVvSKKbFpNi+nF9OhaSRil4rMhzHqUcone1GJXRRRX6Rdz+StmiAij02/gj/ZGKXXKVDs4oT/AIJ/Std4mMl6lMrwVtW9FFFFFFMxZizFmLMGYMwl8GEjCXwcUvg4pfAtGZwzOCZwTKFFi0pMWhJi+mFoJC00il48h6g9U5bHIyH+Qo7WZGRZZfTRRXkv2Fbp7emyezF6kdqTOwhLrbobEKSHqV6CqfqLTicMThicET7eJ9vE+3ifbxPt4n28TgicED7eJwROGJwxOGJxROOJxxMImETGJijFGKMUUikV18cTBFLx2ZIcxzHqDkymzCxaZgSTQuxkX0WWZFll9VFfokULsJ7f+EY14KsxRgjBGCMF+ssyRkZnIOZmOZmZGaMzkOU5jmFqxfqZwY4xZg16H5IyL6bLLLLLLL9nfTRXjTKEtl2MjIyMiyyyyyyyyyyyyyyzIyMjIyRkjJGSMkZIzRmjMzRmjNGaM0ZGRkZIyRkjJGSMjIyM0Zo5EciOVHMjnR9wj7hH3COc5zmZynIzMzM2ZMtllllllmTFqSQtdkdczTPxZid0ZFl72WWWWWWWWX7Gtl4q3TEzJfyKcPguBcC4DcC4ix+Sl8mP9mP9n/6V/ZX9lf2V/ZX9lf2MQyyyxyY9RnMznZzs52c7OaRzSOWRyyOWRyyOWRzSOaRzSOaRzM55HPI52c0jmkc0jmkcsjkkZyMmW969vbFNoWqLVM0ztvZZfTZZZZZfVXsF4Uy72TL6LFIvZieye6ZLqocUx6Y4Fe1v2VFFefKhagtQzMkXtZZZZe1lll73tGGQ9NmBiNdN7+nnsT8CZe9liYyy/AhocRwMSve0YmJiYlFGJiUNexsszMzMsvayyyyyyyyxM054nIZIqLJxQ97L8D2XhQn4b6Eyyy9rLEyyyyy96HExKKK9lRW9Fb2WWWWZGRkZFl+1syFIyLLLLLLLLLFIyMjMcxvey/GheFMvprrvqssssve9/UxMTExK863ratn0XtZZfvbLMi+iyzIyMjIciy/LW68KF13037OhoY/MmWWWWWP9hZkWWWXtfjryJl+Cyy/aXs0NFfsK8Vllllll/o0/0FDiUV+7sssssssvpor3V7X7K+i/E0NFfvbLL6LL9/fsr8F9dFFFFfvbLL/Q3+jooor9/f8AglFFFFf/AAOiiiiv/gVFFFFf5vfkoooor/ML8V+CtqKK/wAzssvzUUUUV+wrqr/DKKKKK/WLooor/Ar89bUUV+lr/wCBrwf/xABLEAABAgQACQgGBwQIBwAAAAAAAQIDESExBBASIjJAUFGRIDAzQWFxgaETNEJSYJIFI2KC0eHwNZOiwRRDU3Byg7HSRHOjsLLi8f/aAAgBAQAGPwLYC8pOUgnwBeRm1M53NpuTFAh9sxrdybHXlJykE29czSq869/hiVfcQXY68pBOSgm3LlCq8+2l8USLvWc9g351OUgm2rlCarqCINbuQdK/UNRNqIJte5coVrqTExQ4fvOr3DG8xYtisWLFixYsWLFi2KxbHYsWLFixYtzNixbloJtfdqr37kxInuJzNixYsWLFi2OxYsWLFixYsWLFixYsWLFixYsWLFixYsWLFi2OxbayNRdXV29cUWJ27LsWLFixYtjttdezVmoQ29g5dyDe74qUevbq0NO3E1nvORBjfip+r5Xu4obeptV+KlEbv1d79+KK/qSibDuX+B0bu1dieIqiv612BcuXNLEnwKpEXt1ZqCJuFTfQYmvXLmkUKuXEi9QlfgV69gq6s3sxQWdsxE3a3cuXM0q7lI1VE+BHJvpq73+GJy9TKC6wpcuUKu5lFG1+BGM1dvbXE+N71U1lZCoq86jVE+AlHdmrIg1u5B3Ab1T1lRVTnUcNr8BKPdvXVmJihM3uIbezWlF53JmJ8AxF1d792JqdTUF1pRedRw2vwAojN66vP3lxRYq7yeyUaqiV2+pk7tXht7BztyE+tddXnUcJXb6kRe3Ve0hp24mt95yEJvjstGqoldvPXsFXVsr3UxQm7iW7XV55FG1E26rd+rufvxRInUlEFXZiNEJ7clq0iGniK5bIPiOo5dgKS5xFG126urMRN4ibhU96kiG3rXYC88jRKk9tKLqzfs1xYOztyhjPdTYK86ija7bXV4kTwxK7qbJP5jtgqKvOyEqIu2Zauze6uKJHX2prsNedRRKibZXVmt3qMZuQevXZO8ROt2w1F51E6hK7ZXVoe5FxQYfvOITOzYaii84iiVErtddXfF3JiYnuoL2bEUUlzqCbWlqK8lqCrLSXFHjLv/IVd+xFFJ86lRPgxN5Bb9mY9/upMc5auWk9iqKLzqbhBNqrq0NnWriQ1iaT3o1CBDT2q7GUUXnUqJtRdXR3upigQ06kmJDT2Gy2MoopLnEEr8FxH76Yo0T3afrxIz97tjqLzyCbSUlq7O2orlshhGELSJJV8V2SovOoJUTaS6s1O0a3ckh6dbs3iQ4af1jvL9JslRRedRBK7RUXVmdlcWDQkuq5Rg8BPYZP9cNlKLzqKJUTaC6u9/ZiXryERP5kddy5PDZSi89ISuz1JavPeuLCsKXtcKq3XZSi89MSpcvs5dWkQ27kIippSkneK1KK/N/XnsKSYq07UJKt7KXLl8dCxbl25lC/KviuXLl9hLq0Nvbigwf7R9e4wTB03ZS/rjsG5OekLUkRM7OZnFzSKalYsW2curq73UxQYfuNnxHolmIjdgXFRFqMWfVvJ4sIX7OT/LU7FudsWLFtgLrDnb1xYVG+0vlQiRPecq7AkmLIdNqLZ24tQoniNgw0mxtVdvUsWKFeVYtybFixYsWLFixYsWLFixYtjtsexbUYadkx7/dSZFd7Tmynv/UytC5RSqFsVsVuTVMU5ahJCqFjRJNtuUk5abi3JuULFixYsWLFuRYttaxYsW59jd6khIaaURyNIMFO8ljlyLUJKXxTnIuScXL0LahYty6fBlsdi3MIvu1xQG9TM9f5GUT+Ab7XtioUKtLY4j/DFEf4CNxJIrsm5f4QsVabjMfxEYt8TeIpOVNlWLFi2yb6jUprOSl3UExSJ46Yq7BsWLFixbZl9k4LCTtevh/9MlfhGut3JFS5UoUx1Kl+cf7sJiN8bir8JV5deTUvyal0xUx52OZJCZRSalTOKSLFDNKoZxcvzGEYV1PcqoT23bV6E2knUKY6lDsKrJS+PNK2LyXHYzStiSrI0i+KUO5nCtMrqJTLkkUzlkgqTVEFX0iyQ3oIj25PaZr0LIpVsj6t5TOKwVUrDVC3IjOnJcmSKQmaM5UET4TqUopVquZvNJEU0kKOJ5UkM96KhlMdJxVmWzeZqpPdiuVUoklPq5RGisc9GOTeT9OyR00+4q9VNB3AV2DtVGt6ndZ9ZARnaXYhlemYTWMnEylXKURjMxErQSN6JYkJfaaZCNcq7pH1cN6FIT18D1d/AdDhwJZXWoiwnpE+yo3KcjYs7TpIn6VEXvHemj+lZ1IVQtLFY0VNB3Azmy7x0RVlJJiIj+JgWANWfpH5b0T3U/XkNyeozk2tPX5OSaCxcGVUenUNg4Tg6ovvKIrWMye8RiPbDTsKRcspYZlOT0c85D0mDO9DF7LGRhsJXQ/7RBzmvVUbeliiPXwM2DEXwM3BYq+BDZF+jX5LlrE3Cf0BrY0JPYcoycOEx6pndh0kJpXCmp3IZ2Gr4IZ2GRFM6NFXxPbXxOjVfEpg0+JTAv4DJZgytbuRsiaQGovgWa3xM90zJkhWN/COiRcIVGNqq5JmvwmJ3QH/AO0zcGwl3ekj1OIv67z0a4A9n2nW/wBToWfKJCb9HxcIVW5WVCh0Pq/oXCneEv5ENkX6Ni4M1/W51vI0VUpCUcxIV0KMyfvEXComfFVMls1sSclcdCu05bBdDipRdw1zcPd6BF6OVyuFPFa+NEruIcJFVyMSSK7FRJmS6Ar2r1K0lDwJrO5iIUgonAsnEuxPErEaVi+R0i8CrnntL4mh5nRIdG3gUY1Pui4D/Qn4REa1HK5H5CVJM+ilVf8Anf8AqJ6L6EbDavtRIq/gIsdkBnY2amdLk51eTYtisWLbbyk6hF1aynRu4HRP4HRKaKJ4l28RZxGyKxU4FYq8DSee1xNFV8TokOiZwNBvDUVwmPBhxHqiIuU2an1eCwWr2MQo1Nn0M6nOUUziiz5y2JTJX2VVvLoxy9yHQv8AlOhcdF5oWan3ir2IVit4FY38JWI4u9fE0V4nR+Z0TeBSEz5SiInIqqFYrE73FcKgJ/mIeuQP3iHrkH5j1yGett4KetfwOOncv+WppRV7mGhhC/dT8SkDCP4fxKYLF8VQpgjvnPU/+p+RTAk/e/kUwSH4vKYPg/n+J0WCp913+4/4dvcz8zp4afcQ9aYn3GnrqcGn7Q/8TJg4XFir9gRYmGRoLf8AHUT0mFYREX7UVTrXx2hbHJSiYqOkbypfk1opRS5JStUKrklJKV5rcOyc9i9R9Y17fCZnPTxaZyt+RT2f3RTyhFPSfIaMVfA6GMvD8SmDxeKFMFXxeUwVv7z8imDw/nKQoPmaMBPBfxLwE+6dNBT7p63DT7qH7QRO5G/gV+kl4oftF/7w9fi/vFPXIq/ecesRV8VNN6lnL4FGO4FGOOjU6PzOj8zQTiaLSzTq4Gl5Gn5GmvArEdwNN5eIWilGRVOiicDoYnA6BxJuDuVRMtGwG9ojsJVcId22EbCgtYnYhRNnVpjouPcScUqUpjsdmO9CbXErO3Eiair7vWVWaCRcIZ3Q/wASbYKN7jo0OjOjQ6Jp0TOBkshM9I77KUQr5kq/6i9f+FS9O2huMprcru/IW6G83F1NJ/E6WJ8ynSxPmU6SJ8ynSReKmlF4qXi8VNGJ5nRxF8DoIvyqdBE4Hq8T5T1aJ8p6u/gerv4HQqdCpJYUvFCeWxPE6VnmZ0ZPBDp5/dKx3fIViRPlKrEUtFKwnu8T1d3Epg0/E9UbLxJtwVjuwpgsNO9kz1aF4NNBje9kiyS7jqVO405Gn5EpOVOwnORN+cZrEQonPURTsK7CpfkVxy6idsXYLuxSkuJaCVudos7n2etD6lqq51EaNi4RnRupvU3l1VMpbIK9y5SqLNDNJ3PxN3+E6l8j/cVRvi2ZoNl3TQ6KH90lKHPthmfBh9+SdGxF7EKKnAqiL3IfkXU0jSJKqqV4lF4lUyV3oe8UmhVJkixbxQzV4lWGiUTkdRYoWLFseaku4395UsW525pFE2RPlyxdmObSciaE7rjRsFLiK5cuL1uXlq9wsR1+rsOky29pKMmSvaTYpoqaPKtyrFinKsVSZRFKoWNEsWLGiWLFsdixYvqFzSM1FUX2SrlUmux5Y581LEuOYkqN3iI1K7+Y+wnUURDR8yrCHNmZ1iULFixYsWLFixYsWLFixbkWLFixYtrVyr0M2alKFXryEUkX2NPUZNQRYlewonM2LFixbZ9yrkKKq9xmsXxNxnOVeVYoZsz2jOM40i+1E2tcuXxUaqlGHUhVylXO5di3KsWM3FR3EqVLl9oZjpGkh7JZpotKw04nR+Z0XmZ0JxZxZxZxZxZxZx7R7RZxRqmgpo+Zo+Zo+Zo+fM3Lly5cuX525pF1x2LavbH1oUUqVxX/ALg6ci/9xF+er/2Mr//EACoQAAMAAgEEAQMEAwEBAAAAAAABERAhMSBBUWFxMIGRobHB0UBQ8OHx/9oACAEBAAE/IVGqEkEiG3g4i6EHwczYIUU4C6LmrAVYHvo7CaHbi+J+INrKUYbHZiXhiCWjixQ+w2hxwXRcXF63l/5ayhEBY9h7jiBLuNBDHcsQ65K36QyaZ+w9rPgbGwRBB9zkICVl297RE83Wf4X/ANElguJBMUJ+8JbRpCoVRAfRwExvWD5FgguBYXM3PRxdEE0wpoIiVCzZzZw8DwoE0hdnY5sTGiD6wXA30PDeL/p7MKPc9o3ib5GgmJjKh0xI6WbQeVz9DxN0bKw9sTZRjVCbaCQvnkhP0+f5Gp+wggoomCAToQpsSlyJr2Q+QgWq2I8kTkWJG+RN5EeRBIgaUgULEUWwu8jps4nZC4HAQ2zTzitzDfBcHD9mPY4j4zA+kMJ5uW+l/wCliJDWzfIWhELEoeS6NBxH9cxQPMMN47CxoS0KROTQdpm9HwTe73s8foPYgmJBCDHsRnkXua3so+TZycOzTyKlyaeR7cnlGncT5FXkmcmvkSPkT5E3kXeRcbojmxV5N3IoKJsSdwvuJR7CLkTwERwX5FNDkOhIiQjyVPkKfIZd+BqDaHE9dDL0Xof+ipNYlveO6HeRF7wLRsdysKwMUYbws1QhpjdzkONvcJt2kXLHU+QlWia1/wB+RiCCiiYS2JEKA/8AA/AcoQQgnwMj9h5TPHRH2aZGIsp7i8bPYLXsleAoP3s90L7jt2QNnLk28lNJmnpdPZZFDmRI2LZqbF3kS9xtbL8izbGG4HG0XLf0H/mrqaI0vI3vAoE2GXYvLKas2yTiZRsbGy6ExDgMMbMeiVo8AJ1ibpDv6D/2D9IQ+RMEFEIPQeLfbETDZ/QWqCBBwCBb4Wg48YTZaFeBt4GBzTD2HiDIKg94erTEkJvhja7Gu5jnknPDE+42jvAkaJOBuuTyg5eRlRulHPBMNrFL/rG0NUzkytWK0T2PHuDd1u/IgngcYTKUbGylwemFhCCHwLOcf9fySiuWWz79v5Hp60MeCCCEzIoqShIUCJAvQTeOoKGTJkwgNA0MYRgkMk0QDV/+B4bgWKWXv/Ipm7+o0VML/wAChDcHtcGvgSBpiHoTFhsv+t0DaZxY6UNj+5jxMN0ZMbKNjZwUo3BhMWvFTfBqyMpEeEbV+NDdbfkYxMEEIL8CfAvALxCJKF4BeIXgF4D1HqPUeo9Q/Ees9A/Fg9B6iScTXwR4I8D9T4nxH6npPiMGgQGgahoEhHgt2EthYtEcSC6G/wDVU0Ghk3JHsGGG3jtkRSlGyjy0NBzE3Re/ctx8GmJL/Oz3CM+eaYY+gITAgkhIQQREXWBJJJJJBBP0MLUkakEDUaeBr4GvgaeB+AaeBPga+Br4wSxJ9G/6S4fGLUyq8Klr74D2HG6AmXJSlKPjoBEOIZyNQbt4EaLn7Rz/AAdqtYYx5EJghCEIX1liEw2kQUg0MMMMMsMssssMMvGgsr/pb9B4jYcpZ95F01YY4iKJlKPFww28IUwNi4RsQ7s5+Zxfszt+Fljwa6BCELKxOuEITFRqPMEblMPDH0saIQhMJ9F/6h9ZhDHtx6jZax82omUbL0iwajY1yElS2LlFjvy23xz+o9Z5YxoeEwQhCwhfQhMmWHiJ7hTyNA5CF08Sdwx4Yx9D/wAFl/zoLp0EkTc3fwuFzpgxSjfQLvFGxzx3L13PJTQU80kqywKbfv0MY+gQhCyhddxYYeLyDO7lwXy2MobKfokpX7ltSyT6FjGPL/27xaXin2JDFEx+iUQbGE8EUYaiFKCDuXrEoOESH+x/Y+f0Ox9XSxjHkhC+jYQGHi5Ri7k7oIXTaK9Sp+B+9/k/ZoI2NUuUJCkbyj6H0PL/ANnwOY1m6OB7AdNnItxXiw0wTw2FLgmUbG7ghxwrnxvg27Tt/hfr+h8OnQxjGPBCF0rrtHJnfxK7wWtUVv8AeIqX6RY2XyOd/wALD5C0Eofvj5Q6KbpSM0l+k/rP/TNEcxOm1NldDY5+ZzNBsFp0jsIuhsTubma2TbXph4dEnzzf4xK5eGPoFgvpr5RwURsS/gv8PbLtCrbtHtKJSDZZFGOrtLuIASneOgx9FxS/4F/zb9BtGsgzN2exhvZqbZaJ4MsbHcQmU5YMLR4No4TGzCKPcV8DYn0PDGPCFkvonJjOnW2NVGLEo9TcNbH+MbmtiVCUOZNo5yXRrDbm9F6L/sexqNTJh5wtBvGizosE68HsXIhvC4OaY++2ehhHP5VDXvRJT04T6GMY8LBCETqeCghdCTQGDYjjoNeCbF146NOUaXcUAtW8s3ilL/pb0X6zesGkr/CbNzWo6E948MEPYujh0a4b0U2FN8FNHaTpxE8dRryl/wCw9Q7CxehjHhYIQvoIJh0DbrE2Jb8YIcBfgncTeHwNfBZGrbgcxLJPDaNlzcUv0bi5v+RSl670PEc5zHnyRYGHz6mmQ8SF2JiFUQnvBMX4FPWqLDu1N93z+kPjWhMWLljHhYIQvokKBINzQi+Cwn/IRq032NlGE2cTmb4WInIKVCFbzb6b/s7iCZzEUJ3bJhraF4w0WD4KUZSieFKzcIWitH5hBw/YI7Tm0+FpDo8nRC6mPCwQhfQbELJnORvRR0chBkxdgaFFOZoLUcx/BjEDXSeNO5vRS5uaUbzf9PBEk8Wm1xEjE6TRBxYWWHYthFkJ1FKvngpoNGUZQl5Z5MVntMc8m/cCWEIvU8IQhC62M2x3iqesCyG7yNfkWj2PymgQ57w2FGLezIv0FK3vHpzc3F6KUo8v/UjWym8D3b8D3l0ijFDnAo9h4ExOnaJw2GNRsapFn2D/ADISIX0YX7/we4VghC6XiiQhCF9BMFL5zET0PYfhFGhoSr0UxKzgepY5DYBHrFUhrGnc3FzS4vS/8m/XbRtx7SDQ8t0ej0zTg0m4noo2XCB4bCY5uxcHAbgRKi667gh6l2e24n+n6nxSwWUXoeUhCEL6DQgmKyZo0OQtruhR7INYaC+hcd0d0U18MSYra3ln0UuKXqv1bi/5T6Oc5x7Ue2PWekanwX7B2SqKCQSsabEJ7KMUTz7FjG2MWil8vBDWW2fAlr9z2MxYXU8oQhfRY0ILiucoaL7F2J79je/Q6IKeHYS4s0Gosj+DkJATvhos0vXf8ClKUv16XquILBuml4nqPSXuLwipJFD3CVYqmJ3Btj00y9ka7HaoaijaoeZFNsWd3fpv/kIRcLN6VlfSaEEwSM1vRr0Lwb2yqkfcajnM1Z+ofJwwKTynipalcFei4vRf8G/416tRyDRih6O/Wek9Yzk2L5Mc/Rr5hoRRlG2csVyaBOiIcvqghLwkNYVsvdX/AEfch/77IWEXC6WUX1YIIIUjlLrQcxpcHYOY9cWgfbG5Go37G94ajXD12EjIQmBOjLi/629DazFd6Lb311++of4HNnxo8Ii2QNTZzjyp9hMtvQxTPeBAd7R8f/WJqCsLCxepIXSvoNCCYIGVT0a9BzlizvLrLPgY5sGKXGiqBS03h1lxSlzf9a+jmwzV5yXzfrwek9BzaL9hiEgvA7wccylcJMcHsHnoWbXp/wB9se5CUr93/eCp9wQsJ4XS8JE+sgghcOQpoHv3Bw/Y1mPSbBt4WGGuLRjkO9dsPSjAqGUv+x1GhkUyTgQg0UxevoOqeg5xi3Bp/IatcGse48wfEbRLTLRqMDw0tv8ABb0vwcve2LCzellwvrMTBCwXTNbQezUTO7pAmGxsb8ZXIOUyiJSVyFzcX/Xanig58IHlCDQ8KmL1ZX0npK9qnciP5NfJvEP9grRdjR5CWvub+kfcaX6s9LC/79cIXQilxfoT6TEwUmaOTRRaE3CHmnrFrY4w2vIw2PuU+FHodUjPGUl4bLov+Tfo36jejQbC6WGHmEIMU6IvWc+ieyBqm6TfOhALRTciKloeCxNvavS/rR5fVv8A789CfQil/wARiCYLxZaFFoWC46cBtHsN7GNj4wes7F/MLU7sdDn/ACb9e/ReLBuxN7GN4LoaGhoRB4nQSTE5HNvRe9y1svsU+H2N0ox0vA/++cO86X9q9/wfYpFKUpSlKUpSl/wHkopZQomUWh4MFbmG9DDDIaDGxhjGPSrpnISouGiyv9c2jmOQ/MKPLh0wY4YIHhpZsgu0KSIcDwLY3UbYrQn+z/8AcJQ7/wALT9CoctRdd9dNPl/gPJcUKRr0OMDmRB4TbwbGNlLk1b8GvvC1G+TgKnP+w0YpuarwMNjeEUpS4ZxICRNcnkiVCrjRqWsvp62cR59hvvsU7VrC76Kkh+S//c3zPl0D5HyPkfLH5Hy6p9J4eCC6xcYLaDdqNYn5GG8G2MMeHgx+9sV2iGLgt1r/AFDNBykXKMeG/oMUIxefeHcXAa+CRV2DoGhBapOFon+uIc/wzQs/0f8Amvof/n/jdseGJnUzVoSaDXkuORj5Gox8DHhjGbIlPJHO8g1XUv8AUNrBtPtQqZSjF0Uo4ICV3PcSuyjO1GpETQWY1grzExCffLOvFcT/AEGM6X/L9ofM+WT5Y/MXv1Q+eIQmIQhCdTH0FEzepwKcRzA8GPpGPD9iJssU7eOy/wBc8RznIRdeelUUKUo5RMjT3nJs5NioEJo4SaFpE2aC6LcIzuXxVQr9Jcfq/I6paz47HyKExYmE5Yisiy8V9djyxRMVXCm0Km0uBjOQxjHhjGNHSbVvaOZF5s0C2sr6t/zdTNZBhEroKzFCkiZ7zQ9nLspdnoZx6IJEcTHcZJ97YVwwrEf0df8A13g5yIooIIJ4omXK+rcNjZSjwUXDdNeh4Q1q7Mbo2NjYy4fQxixPHcqw2X+sZqNZeKO98i6IiMUe4ndnJs59jw92RmiZECRqdjuMQ8vgJR8KDlG4LPafpz9hlzZSfw/+PyPIQQTGGExdd6BSlLkpcNjY2PItEEwVCnA0aE2yW0Nt3QwxseL0pxj2j2sVIQ7jVL6IBegUpSlKXF/waUuW1l6PwiwwWE44pIET3nvPaNbQ56zj0SJkhJg2JjUeU1fYRW/M/al/6I82331/z2N4UTExhhsExf4aPgyy8BlFdBoXTXob9BjPA2N7H9FqF4JG4a1/iP8AhhSlKUpSl6gUuWx4UU8No2s5ya7zobwguKknj9xyFB6jG4NEOxpwwKUokce4hFLTg59DaFX8J/A2UomN0FCF0tjDDDg+fS79j5CxKu4qwnkEoiw+MWJaCWL0c0Hz1PosJTPaNahB2J8iV9xUK8dOk5RIlEovoshlhv0j8yPJTvgVlw2SGN5bEka2QCf0DSg5JnsPfl1BtRqPWSySKUqERu6xYh/vPfjt+Fr7DWqyt+XhSiYmMNghCzcvA0E+iLbCbFi9zyhjJ4LPJF5hHcVxeJcEcmIc6OvQ3wxuMtyDRCdAQExjWDyQ7oLfcWNBRLue0+RPkaLufMVdz2FCoisCwUuGhorBiR5I4innFcNS4ZG43FCmsmmSY+JYZnBDB7j3HtHsNY4CJpJZDjBhvZRJ9zZ2Zj+T+ceWl+rR6A1v6hilEIboSFhPNFgYbBwxCBTLEMG5FDl8jfgLSGnk7ics+4i8+2rSQhGCzctvIbNhiog5HHYwbYXbG1WxiVRYvwPwR6hoFrsV46I4EZCk5HO0HaoXFsSJDVdyTkT5EUJHyI4JKGwNnY/G0YUV09aj9c+g5ilR6D3Co/QcTaWzMamVTceyA4Ikj2HvLdygqcOsrPCkKN4UuDPMCjublHwTbf7fkRqTUa/72GLhCwToywuhY5LCBJIoIJGSQiQa2xLqIiDkjYRory20LpzGMqoYh4J33v8AT9h6lxNEc4bOXZ3joy2hFPSFVEQYdPZAliBIIw+xL7D8Q/EeoRGtCJGtrER5Nh5Jp5H3k2FO4HuCi7jaFYdlFoPNYgsUxo1mYXRYY0OowQmCCzNLxNuFoNybGT3+jgkiXc5T2nPgPc4dEOx6OhtKFwYpSlPCoJksVVD8V/4n9hLV/wCp/kuEJCWQTAkJEIQhMIQaNBpxlg4PZwRuGNOxa5C13HAWQfI7yVHZ3G7WvD5/sb8u/wDRu7FDw24PJoTG6xKOzFiGKNYuijY1tjS5jsac17CYT+CnYTQnBjsPPg8VinZkOKcimL+By9x7wxdkzwDwBRwZSFBiBaPCIdunTkQYhY0UEg0YtlhXYK7CXYkLP3RsNwuh4dbLsKmLydoHciRy7Ldxrj32esn26MYnBRsuylxSmhf+v/RM8YpFfaP/AAP5L9ZdLE4mEw0aMyqK6e1ilKcSy6GUCWNUH9hIbspu3b9Dxt2jO9lUcsNfu/L9eGOfuPw90Qya/wCFwibYwbtkNCUmhIyj2JyPrRssUa7YfWxDfAhhKKiooKOhGAjN+Y1PEqBUgnwLwCUEOwvsEIQhZDyBcKUQYJkpvikQlEiYQoihF4xY5jQzsxe6KLg34hWxMcDlmPX0WrRV4uGIb6G7Pu2fc1Pax50IS2/cWfcSuEeW2EHWzxihdhFapDWQNjsguCbiCfaj8GG7IDYlUKR8I2XpEuLiewczFDQX06Fu3RiQzJf0FSZpwipFsPVpCy0XA/Qh2FLFG+TJYcmJ9hSOwEgtFKErsKXYS7CQkCfBDsenCs9OEeBeIQSQJfAk8E+B+mKzExPoCHhESKhFBISSSJSRxbikxaJMUZ2EMa7HYDshaw6vsGNoeEPtTuEP0GVuhcBsomPSDVPKEIRJpLSLNib+7/RM4nLkl+n8FiefAm4HHGxrcjUIRRqGmJEwHvQLi0Od7xkToI/ViMt/I+19gbmXiMvIdsNDDLy94QwQHhngEOwhdhJY7mk+Md/sfiOxCE1jpIwgSCTwILOWTJFkSFISFgywEEEjwXouZR5iyw8FGKiYoblFDcjbAhqiKeQbGUpRo+w29jsgx2Gew2grsOwWIc6w01jgeaePkFNz2t/J/wDDHrNpSL4EMpRo8ihJhcLRS9bDRtjMdjYwhCIcGNIgQKERBIlEeDucHtvZARIdIUOeGhsV3CQk9wvIJBKJBJ0NBpM1k+5HuJfkXQh9xD7k+RL5EnkS+T5nyLCLLxpD5dJssMNoZDeHg2vJAlEpI08kklUPssx0jG8lwowwSsQuGzlCDg7KOwiX4HnRzYNdsh40B4rHkZtKaEfLPANReRJJNsUc5JXil6L1h0IMoWbQww2UTIEZb2ku4nyJEoT3CF3FdwgHxgqNixbFeR7w55GWBw3CmJjDCY4NQ1JjvNIWdrGJJYlgTCpcV5PYPJov3I8i9j2C9j5j9x+wwjyMGvkQjyjyj3EeRK7iU+Rb7jHcS+4juIYqaI7RFbE4Svv0ZtDtiQ1iZcXp5ZDjljVtK/Q3n7CDZp024Qx+C2fLV/LN0IiUgvDwXhKm0JhN7O8nuwpZRPqhfeKCMFRo0NoZDaGg8DDh8FwfwnRjyeSOxs7wwBsd6EPJLlMVgCikrQhYVIgl7EL5KQwtix4J8ZaLzXqEwyD8PtwNZXkYhDyhvca+5p3I98SX3HS0xqFO4n8lpyOdywY9wxux4h7mNT5EJrZFsbooiuzwZ9sXIt9zvgkY0mQIkIUQpS5EzuJlEyu+CX7n2EloS3uPccjawM8NkOEUgfb2i/gdxFvviQxJheiGsCoaJxMX5PkfMfuP3ErFDcYaiRHjAR4E+Bp4GA2D0D0sJC9R6BeI9GKBCGiGhHkjyew9h7iPJHkh9yw6wTLhZhohBrPChrcHAa2fEdNDTEeBlyx6eHglQh3GMO4KrUY9MY8nmuPaJoeuwwXy2d0TvIb2G+Jj7qi5uhVtncCXk4sZ2Q5PaOaOQZ3MS+4s5YLhRMVYTEN+6n8CFP8AnYfeBiVm0ahIb8oeBqnwMNNnNDe+FT7ntEFlNod4JQrRZSGw2GwwcuRdLE8RhloQQMPYPzD8h7xiiIjwTQaDEPJALGThME7uM9x0+TARIUCBAnE4isPwEIogaYtKXQxIaCRrQ1fKK8D12HhwWIuwtrWMaFvL9hPs48auGsR/8MU9PYocCk1Bge0K9nLpHCtFbHRJrR41Hu1NozRAQxa0EOyKeRxqHWqPdGMawojgxZKFwTEIRv1tp+X9D0yiHcY0aGrR6xphryRRfEHxCAlD1gT5I8mgmN4JCp4YeaMuE0MsY2GDB+YeSH5Dyh+QfmPIGNmzeICq+jGxFsnCXgF4iSCI0VEkkEjVlhwJjQKNIc7aRMBFzgYIvAtX+xrt/IjT9wle0F3ZfBZ62cQoMQj2iYk8/YQaqOMRPA+Rusigmhvf2Fnl8Dk0hr1ThfsEvaS6FttjkAVhm1DStSER/wCpRF+sTaSvkSVZsWS7wHGzZynBS0kFZ3qRiwIbDGxdHep3+Vdfsye3MwLZBnbEPOTEQ1zCHyhsmvZkaGgsm9Gsb0csKUuHhomKUuGMhCDGs7ynlNIlnBNhGkMg6V7MYDH2GCYssryX5wpSlLmlETJkDkW7Qp0GDpBDbYujUqGoPijlSc9sKp9kJ7V8HcbGF4O+vsLubCqNvAxfdTmVR7yQ05fYcSmg3VfYlFFafs74Ma62VdDCRwXsYfUJqu1LgGs2ECBI7Br5fcU1FD0xOzE2ELGbdlX2HI9vtRGS5tD5tyfhjdDReD9koWL7ic0GpvWcT2ngN6X7ioFIn/PyQpwX50J85ZOiVDIp42yGgcDuOcDKuCZRvoFw3hS9Dw8KTwIaeExB0aC40TCTLaU2RDI9cCK6so+l9mjYsQQ6UWNr8I9wfYIs8AxpoESQlzYOUzb98KHamK/Qa0+5ns9iH4wtJ+MfPmTVJ7ByIuSCTSuyNConBVPwmJCTZ4ekcXiNRU7uY638VEu1dkSXvUCimgVjTTQ8DVT/AALpiNvUP17Cn6TjnkH3mtmxwsNMGr7LJxBl2g+S+TGwEU8J6DE8Q+FK3x+Ax/QcewQWr8Zs0LTYnCm/gLbur9UPuEGfALbrXlHfo/Yt8M9keRQS9xfoQxswv1PiNMrQ0FmUpWUuGhjLkeBsaozIQUQYjiw2wFNISsWBaj8JMTiYTicosrKJ9C+nQYF2DT2jJXpl0DnEkMrexqhkEd3rHOhPYg6JuD9VRCKLfN5BYvcXCTDxCfbwJN0Un6RLAhTjKupr54JxFEtf1OLiDW3o5X/t8C/5ufYXHYftvU/kRG5L7pS/eOa378LqIaE5l4AW9mqjGG1/aP8AYcJq9UNsiNeBKZQXaf8AoSGzwC/Iktl4qOiD2f8AKhBr5DDaXKq3fpCW1+Ixc2jLzW9Xzr9h9zx9cG+cabj/ABgpJfGH/wDJNuKEbNt81SJpfOgFSAcsjcJC5KjJwOUNB/koL6SiCgSIEMYPMJWbjeMPoxqjTH9DleghwQEFCWXHIQbB3EukASEhCQhYWV/gV40amrzo0IdqtlTw9jWNxPtP/R+AtG4TFd/8w0jU5Q+CGvTRttMeEJ8qofo9wLx32idz8z/wdmfhNij+D/2eVg7tfhfwLsPkK1N/lz+00/pMfsGQiaChEVJad0NK/Ozv9h4gW5B2AfKf1mxFKvsQjoXUdy77boiErsKg27FSTYSuwaOw2OxRJOwq7IX0bNlcsTBoxqMP1Iwm7iQTTJllLlQNlMbmCE2caKwme6i0YW8EVtGzTEoSdyw0HkwJeFqVCSYvcjXcaRsoaFgns7DJT3VxIQJiYmIWaLoXSjhnfCFwm/ecL+afyw0v5GOfkKJvKQnubhb1+h2i+L/kT5XxP8nc/pf0LmWEP3jnB/dbZwv4J2ivtOASXwix0P2fsPLxT7lixz6Nh48w4zS3f4IrWu6zEMU+EcdPGJsN2e43lCE6NjUvVBLwNzYTdBbLkVCBCQ/IcBHGLerdHz6DgmmhBaj5DaNF9iYg4xbKXtEPT0M9gTHBeRXVhBpjXYXBiW8C+xNiULyFfDOQDfGh6t7En2P+c/nFaF7CCk7DfwfqTJw/5x/J6h478/3HNfPH81m/4O0Xwwj+n/0d4/hJHPfPT+jul8udgv5Y4X79HBn4QQ/gEUbcKPbS+WfoXqP1pP7x8j/v9nN/bVnO/YbZ3z+GfwKe74/oP2hf0nH/APN7s4P4I0eV8pAjzPhH89L/AEJ8OS13PsOynyv+Dt582O39wHCv5T+RT+iPu/if0DnKr4/qGuY+39BBevZE/wBkUHvag1PfuvxRHz8lCTh9GqH9IQxCl+hcNGPEXmVZeImNGMMNi/BRBpmxaVJoRsNTtycPdCY96FJyIm+r2J9LaLVxBhXN4ZzGnk2JNBIuEY++mvY/UxafYSj0/sUELlVRMX9xM1PlNdyode8hbPvg2Tp9MbpwLfY2+V6PVROu35wXeC3wz2NDbvH86E7b+wY8M11DUfLohPAaFn3Bfr9X+hcSxX8Gg434S/2J8Cu7XyXeT4V/ArsvkuHXyz/gr0/mv5HIg7H/AIPZ/FK/2Ofsf6TlvgEfxyr9h7l/ww9y8WMVyZHl/lp4sf1kP/Czw/dZ2DljhP8AvwMc1x+oNnY+zO6KOxR3OL/1P6P2IIjzk8m/IuKz8mu+ETEzKv5C/lzDg1FeRVeXgazz3ev8IbL521/EXCnZaEsQvptcBAYsJ1gV/QJlxKNTUSzOmEDwvHP7o5yYhY4bfuVY9eGaBtMY1/JcDRpfnuK616HO9ndDS6qfdMtKvwLa5fssNfyGw3UvZZ3T9DElabukRtBr4Fo09nYyTOP50OVV9d0I062URnLZ24+5D1Vt/f8AoUZvdr+Rnn7zYvH/ACxJ/uYq832OA/FFX/1cCLaWq2J/37CFwVf9+f0EQijv/B2+452V4fzcD7P5+n91oSp/4q/yNZS+4zxG7lhLj84+Er4ca+xk0MuF50XtA8lDCJzxWOOCZ94v5wwkKmLzQqOxyGpWheXoSZvGs2eG0JK7dy2S/UsyX3/sMtXf4/8AAlw/kZNfxf8AoiPY3/yj/YeUn9m29skl/A0adflf0LHtrIlzU20tvb0W5QDzRvLb/LEvgPg/YL1nlf8AKFlwu/D/AEQ+12NSM7v8BPsLVPgahJNQ18X+o0kn3Mv7O5jyIl28cIVF8ejgDqpc0aLuJQnuO9gqkCVMZaNcDPAuh6UomUpSlxOiE+giDezT04ZT5+RcgbXCa9ibqk15E8Nl9yr4NaP2GSRt5Gj2aO34CFRqvIpGnH4Nwj64QknK0Tpy1sXZMGWOLsFBaXn2LyW21DkRcGQbNEef/g3m9FH3lvks0FFanP2eyJcPgn6fLsIcmvy/9Ecl8mEbT38ihuv1/kX7fiH6uRo0X6L+tsY8/vV/ggCO2V/IhEfgj/gdxeiEVUt4TUkly9AuU6l73Kl+4N+yfka3TYj9MZPpvwglI1rxuaNnkDS49+4lbPnGngR90L1HhnlYV7KfwMe/DwkopOZ+ULuqn8EmrQl2+0NtP1hLT5+DZpfgV5U7JPg47pCY8WXPWxdhJfYXMX86Dr1EDP4MG7E/B2L6NH3B3oUbRCraSN+EO9k9iK2xKxeAXhOMw4ORofQk8UvSCFKX6LXQlaTtQweRvsaOG0/GPbDbNrg0aOXHI3TWjlvbxi6pBqNiWSKa2vcSXcGxl4ENVw/BIthsp/gvkoCl/wCDN6GfyLgaS12RbqDsUi/SQXpJt8QQ5Yvgj234LdmNC4IrGafEEdiYjT2E9efkrxr4IeGLcalpmvInMVbZoevI3tSK+BiXBb7MbamTJyKfAuwDyrF2mPKxo0xDvKcuz3EeEu9QTuwp2FGcLJQmdMLFY7jZ21HbRcBfRo0dzvGLcun5EuIgDSetjJJg/Yx99iEQzQiiuhasUi+gFxcLEFgpSlL1sXwwNo2qRfXwRHAXkIvQRBcdyVJMI2RR2HcTQJU1G2vgUJpeiYH32R7oJuWUvSxzXJ2muLyHUZiUMNSWvDK+bTwKcW3G07T0/wADZ2/gSu38Hr/gS+z8CR2fg9T8CxtLASwS8J6z1Zg9AXgHpHpHqR6ER4RHhEXjEzyUEOxCR2Ep9CkDR3O9HLIcIPuJphZcF9m7YwpMoK2uIsNNs3YkzgR3w08NopY1jXRKXDQ8oTCiCXSX6DNoFzbFoJUSPmwvtYoNSpacqH2gkevyMIbgqNcCI+35LjFVIUH1INE2qeiV7SmMVLyFpClKX6sIfLCEJ7J7J9ekeSHcae409zuCO0DhV9yn9hYDOKaX0fvLCY++Bk+cbdgtA57Qkl+SZUbf2OjhQ3+hw7L5If8AIQ6VnfygyZwVp50VfQqi4hCEEIhCEmb0tcOZzlzgJvAoWmEWky/YQKpD9EA0F8of2DRJkWSO2Ke1HtRQWPXkGe4vKe0957T3ntPee8957Sfc9+H3i85Xv1BA8kSSSQT1UVPSARigaIaj8gjCHaIpw2/hH8Ej+cGcEgfBXwcpPuQ8BB6xHsJXYSMBHjA/QjwNnY5MIeJ4QNak/lDXZB3IPqRXLgdqndUJhYogisCCy6XKLhxkpMaUuLhzggZKwhhcRuWiiNBjQVLY4EoIZ2AzL8Q4h32F5j7DC5iX9xrgu4U91a9Htx8NEW1v4Nv8IpqBv/FhqX9BpuXwOPH4F+PwFkBDj8hwpdmgffAdwgZ5X6n/AN3IPrGL9A291hUCwWsdJQXlREQIrDFAavAx3E7uU7j8h3aI/cMfP5CfLf3P+XmT0HqH4xquyH6GpEhNFRrCDNCSogxMM5BHPIcboeu1ON+wPOU7uwu+HAhByJGyfOKKz0FkqMLldFEUE+SgnehS/eRRZ4BCIY8Qke+WIwkHPoo2N5RsFodERj5NnxY7ZDB65NWionbg0iwzqG0J6HSNuFb4a7HeaeYX3CgleHlcLihcOTaFioa6CphS4aQw+kyQNB0ORhIJCMW82LGDWZRodkO2iux4h3ho7sc8ed413QU+4p9yHguhiyaIfqxVyeeMdyotgpSmxoHtC8lREoJ6IXYmGHZTYlGJniT2LQ/QSjqLPs7LWBMd6DmDUbeDcVQZMuMCUHVhoaiGyOCZzAh9yohiRk6ExMRQog34EOI9OQugCR9GMOsUO4/c+Yi08iBWCsC4grKXCLiYQawaGEeBq+x2Y8B2D8HmnlznhT7i2FvuK8iRsnzgWS2TwbmCL0xLDZDULlg5wVLHWbsSMTE/I5LmOBoItwuKUU6bB4/EkGixivmLQsbSOhr4xQ8No0VIaEdBDFIxI2EwpcDU9ohBY5IxQtFgYRogxZdyRosUmMBI6KiMN2VgnA7wew5GsTCuFMMuC6I8bDEIMMIfbM2+h4XjlORD7nkFoWRYpi3GwoanEEpBZMDwaU5GFtYdHiAtoVCCJrgbDGuTtnLZPE4NYkbIeBdjEJGILrGAtxdDGuWggkNDHPJ7BFe+G3cubYbysc5Y2ONGwwoeNw1lBUxLEk4voVrlYdEhZjcXOijvmEIQaGHgIfYZ9I0N7hvce8V5E92E6wMP1x6E10zkYYToo9CExVnUc1GOQsXkZU0xQE8MEosVp59FY9BPJYJhhM2QxlKKRQX6ONPoU6ORBcFz4DYbYpsloIIXEy0PC8WdJ7Q30ILHcXDGiD6AsSHgE67nvPaLeUGmasiIPMonMExYmo3BFAgfFH2HsaZGImLBixEdjSGSUViKzxFZEaEIhCHAmUaGiEwmIryU6MEFHg+g0bYR1c+ApEhFwi5uXg1hEwYeFBoaLguqK4ZCDFMcy0WmNQ1dz3nvPYe49ok8nzF7nzNGXoC1KNUaGRYrwQZ2yQmxbIbdH36YCQSYpUQSEjg5FiEgiYQmEzOkNMjWLdKm4aJmlKXJR5TEylyXLY3hRpRiMvAYgy9RqXCw1SxUaMGyLRtCZCcTeRnkZ5E/kXsfMYn0cYGPY8CgVdA5JhBIXB0MMvF0wYYTgmLMgmJlEL6IadIpYYGo9xbKIKhG3EJmlKUv0wbyZsEE4y8CKCDLRjQ0TCZpRYpcMaoyhi3jcNREaZWUiz5nzLcWHJxhCGhogwszohBCZHnIEEJwQTFlcCYmMIWJhMj64EJjXGP8nvLYkUUEasNdAaxS9QXJRvBjXSLisyl6IQhCZpczLGhoIFPA5YWrBRc0uUFsg0QeVKUohYTghcNDQ6wSFE8KdFiCQnhCCYmXMzCdbQ1ijUTPZjt0AILKYMTFL9ANcNEGhoaH0ssVyXovXS4YxsbGMexAliGbyvAsVuExPFGEQaw0NDWUIWhY4EKIY0MSjCGGEEImU8ITEylF1vNLh4pROhy+jIIVg1kYaIQYx4pclGMaGGGXgghdF6BRPF6HijY2NjYylGMeXaUtynlBPDw0NEwsIWGcCCeGUEyJQTExMTvQhMTKUQRQQvUxjHobH0XHA1Euk7QXRVoxoYYYaGhrqpczDRMiKUY3hcUQQpS9DGMbGxjY2NjYv0DTOTgTKPY9FHvoQsXoTEQgw+jkIogt4nRcITEEy9LIQaGGhkJ1GLpOvaWEF0NaQaGGGWWGITppetjZcKUTEylKUpS4bGMeGNjHi9ECYhYgnlMuGiQhM0pelC2QmDwTGCFkkQnQIQgiiYiYMToDLLEJkgiExDp+kivBFYriYPAwwyyw1ijxcLljfSmJlL0XFGxsbGxjGNjKXqUEdsrpebh5osrC6YQhMELCFmdEyhFxtkmDyX1IToOBT1IogigsVLiYMMMMINDQ1ilvWpiEIQWaUo2NjYxsbwYbwxSOKmSomc9K+jSiYnlC6HiYToSQhfUomLEwhOhsvMmExBomEF1pAEFmUo8QaEGGhoaGh9SlxCdEw2UpRsbGNjYw2NjD6mmWAmIXpuUIQaLh4TKUQumCRMkwmUL6aEy9EIQmD6UPA0QhCYovoUERFZVLhjQw8DRCZJ0oTETLwxjYwwww2NjY+kO48xqw1ExdZEpBoa6VhYITExCEJEJgxCExRMTFmfRpRPqhCYMQaGhoaIQg1imgiukAo0zUo2N4f0BDL6ZFKUY2NjY2MPBhhhh5t8hPYwYaPwNrA2CYmUue4hnLmDXROhYXYWFguiZeELgTF13CZehPreWMfOIMg0MYxnYowYJhMJhg8QTKXDYswmTQ0QhBrCw3obKNjGMfGDGNjDYw2GFn/9oADAMBAAIAAwAAABAll+21coy0vFeFKEWPhMii5lhFfY4z8ClbECSUP/P+0Rstvt7zeCRg/wCgAglao+Pv/jdFXI4iKfVMije3gyDEoQCz3NfPgh3EAdr/AHRBLTe6w4pCj3m/IIAI/KE8y/l7gldeomCvQs5OO1zyFxwIO/Y3nRi7myWJlfRf+1i2pPzS531FhJaW09zHx2L213ifX8zvZNKkvpCmAPv5hk8a26WP9Eq7NiXqVYtOCFWnv+VITSai+ostAW80fHEORjS2BGNIzO1nS/qZA9IItv7v8L5N39ZvG6WWUoia5pRyts1+1422Xb4otL6S6Ct7ViJOn8FKJYH6JfSOt6StLlo/po5oHIbUmKugGharHXpKQWv4TG26uTZfRwnvQSSeO8HA++YYkBVnyJE4OAY0RgLUUsJaEkX2YV5KpGE8hsM9/C2G/wBumlk9svf3iovSik5WPPJUKRLSlNgN3mpvBrepXAZ7162EcHlqISeJnRuw6NhOUyh/t9vdt9+/u/lrM6wgdtGfYo2SC0G/uL6NFp2vChHWdJsP1DiIgNJovre5sfa4pfWeJ/vv/wDLfb/rt/HTg1Fxt6ogKmDSiPxlkF97NKO5Q+kJkGHR3ZmSyVesbM5r/wDD1jb5tDVkNtyFVL/+Q07GWtcb6P2v47DYgmb02luq1DWMTx3YrlaiHZvGUJDtkLzeJt17YQJQaTaLSTaS23+1PwskeOvpfUldB/pOzjDNdRePsFIOBsmvwuvfC4vlOsxXyuNkvUY9ZKBDZJJFaYIp8O2HlztuhnT39yU+om5ZeZcdpeNuj0hKTc2zoMwkcn+vzcM89hTbCDTRINvLaSYgBCQvmEEU4VDwnzEAK73ymvNjfc91dDcr4MmpJiw0H++oqox9W9kpRYAQBI4jJJEk+2ohBV41omC5Y0tl5BTgADYRZa0WOVkdz0j/AHtKVglpjF8hh1BjdpMWkACCS/tf7XsvXLfoCfxbTK6zotW8XyI/2QAXZ1Q2HIfIpOTBstahCmokErUYKZhPm2ACKJdIqr+0kk0JPYcBBLaBejDxv8HENEqHVyUnmJmXbTmJJGnybGJI8foYD73ML40yQD+dvbeCSSygwElfIwxJLem3JA6F09aDjBDIrBnSU8u8m6cH6O5JxSdQfc7/ADzSzxBgVyzRVNgS0AAACYJl2d6zgSDqG/ggBO4Bv9ZEVzxUgnj5MTXkmA/TF5Q8mL/j3RmsAFmHQREh826+2Sf0+F92Dyq4Q3MknfZxyAtYWZjM0B5G0OMRflzYAEoLs2V1UEcbtFkU3aKonF7bb3/bbe4QLsgwEoDQLlXnKWjZyXRLrpXd+Q2NESFRvQ1EI0hzV1CopMgACHW2EhHaALfP/safaWVBUGaSr/0QKGlNqToYTMZO+Nrj510EAEd2dn81+6RstgI/sAlRShEz6OVNNsFuyS9bWEqfDpMIaeeJjyGR68Ws7bardLbwkQdZ4WeRHe2qkXDQHbNA+D3IKaiOBb0fLNcqy0eIM6LGJTOxIqCb25twaam2injgdmixNzf9EW/SF/oc76e1Mh1aIhPfRNYf6wEjLtV218Qr3M4X4OcSH5JF8GIXm2BbJl9kGQTPt2t0DVwnnReh73Yhuy8qb5xpKy2WW2xDbAS/zE9kis6Ha6cILNX9F0+rl8gg90j9PomgTdvEbVeJqa/ZYzCfkjbAJGm37STa228vE2b7I3ImKx/mSMVXyRfFX/e3Y3BwdN4PhSJRWYg92VR/u+EAeNEbelngSTbLPYSJWSb6qYN1flfhHYbwL6fqAh6RbYpXLwBzntdjm3iOKHqioefPjTSYn3qs0yX900kHLbZSY7p0hFoawllfmuI4VCFjqHJwgMlJl99oixt/4AWOAKNfT4vi1B4+Fgn3NxAAAEthb92wIwluLkV8CV4ye3FzuTIwODXm2GKv/qpczXnlQ8kiEp+oq3gphxYQ6L1AEFtyAAsPDW1g1N2S7BmNQJV7sUXGIpyP2208Jv8A65dyFgRMmljWuGg0EsIa74AmxABRhrJSOtbbfkkYjJe0eQT2nnIW2tSwyu25x2zTeD75plECDYHGxslNntmRmuvammNBRKe0k7aQwLnOK4y+VifReHMHktiyTJUH7TTHbeKsqW4wKFfZu++IKVDC2n7niZVxAYL2tt/8k3wXt59x9u3qUawdm5tIE5w8Wb889+gkjErPsPug+hhRRmdu7Ep/w0wp/CIJU+m2/W0l2tFv+Cj3YvvziMe6P2PauYzdYiLQ4P5ErBj/ABcP1TkW9FrcD9A5WwDtUAavd+tp/k9f/wCLWujL1VB7tNQ+l/jl1WcWwaHGwSEkCuNG0ib6sgGygbMdQna5lSbAg3+fJ/8A+/t7cv4dPQUuzBFm2v8A7ecxt5KnqW+pCq1MQKjwqPjY9V+yqQJoPMvemrvCQ85N3dtt/wD/AP62vqv6n1k0ml/eRJaxIsUFl+9asyia3QtlMwxBj+eATfMpkoe+OUN4hWo26n/zSjC/9SlSJjSMhx4uY3BkH/q80gQPlDX6pS3572I/MJDlWplxCpXg3l1J/kVMP3X2DJaqW02538wqd2LbAlYSx3+nyLMizcNnNikpzfhIMIje3PXg+uBvlun2lHZ5/fPz39D2Tb6IcHnp0AamsOBAfOligfkzLT385JcwfZ1Ud70QVBfeGZu+KLLsz8vl8McqGR/NCxaW/wDzJrdJZDlO8WpSjpF4eI5m2/8AsCh10vytPNXKr6ACFZr8JlLJ5YZyAT1AqY8xlt+AFLk3B/ObXkxT5iFEo1Y+kns92Db/ALvTVZ+8A19aZpJSMwwcizfnYcL6tEMUueM9rBQgS5UT1zbfhxB0lmGx2o6ButARNv8A+5+qMGo8kniDKn8I9Zbazzt4cpPATWTaNgJIRVqMasZZLVliodA07d4x9rPaXxkIFg86eY64xFN94fYletYPYZhVrvFel1oMySUTONCyALjnpOc8/S2torYMi4oaV8ZOtZcjJdtLJ9wvvVeUMgbB2Ii+UYnHSJlLbJsZ1/7VhZDfaU1ovbwT9m2o0U9fTRg30rfZlH0snBVODzOXTu7tNUTq0danbqEVnZW+ywDSJPKDeQ97SUcfQQU8HL/jXzZX9YS5z68U7xE+VqLNTJELGarCAr5ORmjH46y4d/ZPUzFSroNOU2yeefchjkiOUzeLTKU46vNf6NnmWTqpqxA6jU3tWz6TQ/6HxttYbmHDiRsuMhSHmuacbr9SyCv+V9xOFnrlRkM+hDTZShTaD2jTK6uiqNGgqAto9QHW5JgK3Un6/dfngi9U3VaHFYYOtuxrMZR+Ml5D7s/8I2FHEHrfLZI8N5ZRU51m5yBgY/xB6WJDQPBcwQLWP7M3eyfxvCHmqDwGv9vkoWlBZ4u/cQu7DA90RHCrYwJLB7hIG5CMxAXF4/PWSNOZfZjyuwiFY1Tg959Xww9ApkCq09k0A29Fak5XglIiNXTTbHPXgsXXou7a5Kq+Sum5IaafkhlDl/lljaepDTlj3oKcs1LIrTQFyAJrVtcLG4757CSo5LuOu6LTJ33g/BjzfnIU6GRKVIeoiImIjt9i4aNatKcIrvWcGwnlfFSx7KYotwUoN65qWVEnv27Jv26c3Hd0OaHjX8+R7+uKgM4rrorH56ugWH2OPy0DNxj934697+v4MuIqcvBNEbXfvufhikjdvOEOyTWYox27KHd0IsMTrf28qFYjr15tSH2Wgq9Vr4X18sGB9iQuCfhpfPw8qmR0Zz54zD7mlj/LvYOsNQfYMKjdYwxcozk5DSt5LorrLzzebQtierpMKotAwv3Eg5fcn5Ph9o9vUWAer4R+Gc0M28igokq0dIspC2SZOtjozGRdR28JChFzkvDWxoaYXyborUfZ6vw2Uz0n2JiEiTGx2KmB1z7Pp+Ca5Au5LYP+ydXjiibF31KJ4gQzCZUYu8uBMfKhpZstY18lLH7GfCojH3GIyuosKaHNi3LeqlA6rWQN3WFIbm+tOvj6oW0Hp9M3WHKMPi+wO/BlobRDBJUDVKGPnRv7j1JNr6d2T21XIbce4jAInotTE4LWxwAd67XvYohPVxZNFo7GzR4qLnLfBHeyLR/4vET9QPvfSv6A82/j5EcYqz+c4z35sKrpEDqHIXUtH3V3QXnl2BQjDLNLByUcfRtv/wCazsHZz59M7naBN/PX/Nl//wCCVxPeM93iKhmd6WSU3Gsy4CA8oVhSeFgq8HPff/8A+10qmdjtB8/t99rJcroE6fk6mI1xdmjOs2Osa8Nn/hy6Jccz9LYOf196G959Ufls3vyRagdgu2wpbxkgqQUtvkkKw9/28KlZuBU497cXWa5GeQ8fwTMDwlnQQluhyt//APbTpKvd2uGJtFpeYrSNCNNV3P8A79MFqNssanOB+F6NuXz/ABdOqn2Wk3Ixc2utv/8A/AMJKMVFUZy7u0t1Fx3566pCpv5OQjqi8k3nTiqDSBl5EB47FAFevAFDqYDG7LoOP+ENVL6YHYUPN7wxaAqiCjqcuWyJEyE+K6F+mumJ9QFPHUuFJVzpQ6s8XDmE+puiKFeXfojp2AdHeW+a+PZqPVLbafPGiHOrzPXFUZuQTr7+obKeu2ab16Z9hzD2TpTZQ3aYTuPzdaJEwhMWbXXcBN7mrtS5D0GUFl/vqrfgQUp7ZULC51q8IqyF4OBIPiTn/feWQ1TNmTPth4m0y669SycwjSNQtSR6dmjXr0IfWVlPDmNEzwjcLVVLprumH0fHPry2zzb9+FDYjj7c+Tva3OZzNJ7YEaNq8lNJDO1/0cE76JQwt3bQntzUyu4/Y45cY4LFoSbyfT20mzLzywqC1TSR33jeVnljf7QDPiWwDNR/k7chUWK0exbIlJRhs/TySpNTmPHsr7yPsaWoBdqdXKhl+xCRbunvZWmBHS2lbAZiaetG71jRI2Oa8OZdf1hLpwGPQ+RqiFe5mf6NIJeSbbd2wFEgAluAqdf+C3dZ/LQYKZTcAuhMERkabzna24Vz+zavwQzVjdujESfff6Tzcr3YMkITZNNkEpRaNBGVXuvCzoiq2dE6yRaUHyfm3KB+FDzb3lVx4r/eBPQEWxsszmDLbU7sBFWdjBnWIwT94EjELAO10fEcYlO9qfOerARe7o/cbJbyLUhrkNV//G7TJNEOWq+UTMotBXwPAhWy/wBWXXRZFDlLXmLjbW3Mnizpgl6ygT3rWKEn4yHR2lmIyAeNz66ITzcPTBBaVI9YSsssL36Ir1uWQN3mGSAii7DEvDxD2vy690USq+L05ew34iZgkAbpAL1Ls0LCAR6AmKLFA0z9oCD0pOBWX/l1Ov6pcmkKtNDCSlJ+1cW12p9o2/2Ac1oBq3d7Iej6gJTZb7TARTsa33soKAWsj6raQ55bRpUaKEFcaTbdhLVUbaWQIVrfe66qhuhI1YXQ7SNV12LbSaJLYaixswQ03Oy7E98ZbEPaozjkbXfdo41Y7Njy/or5nScJ/aZnq/GQ04gjNHTg0SQQjbBJuf8A7R5PPqEybJRcQRsZjw40onRKJtJmyTvfEkxd299tnGPMxyiEflxiVfN7xwcsoQkiEKrbcQ71B1N4567E7BcJJ7wVAyrzQ5Flv4bd8w++uyl2WtAYKE5A0a+FTpkEq/wk80CpTgYusexxn4Evq1CUNgTmMfrP54cvgn1vMh1svm2gBKyutHxBX18iscSu5t0rC19s2xMZPKJyUzsShpb3AqrYoxp2/wDeznuM9mQ6kfZbZsQIYZHJJTHwQmscu542yQyWL80ppArQpU3RmXD1bPa166Mxlxtc5TJJfy/VAPS+dp7tI3lJY/Vd7c33ouDt6HXP4mbidzVxuzLE5tNj/8QAKREAAwEAAgICAQMFAQEBAAAAAAERECAhMUFRYTBxkfBAgaGx0eHB8f/aAAgBAwEBPxDsHXBuzodnoRBeezyEJ+ijffGp45WiYmec9aSWRuxhxPvDQbo8Mbo8i9C9cmYiYkTFqFqy8Xxv4HzYxjy2NH1n1nlRtdEUxoQfio8gdgoqiDdjY8I7EiYgpfDgv6sds72xhhhMYuRRetmrPDCSgh9DCZR5BEFh458y9iENY3EPZlBruvOFEPA9j9HRlwQUaITJqQsmImIX9O1j4PZdxcV5HS6Q+xoaF4iG+R27L4Ax+RjZcfyJUmeybUFPb/0Nl9e2JBDDDDDYgEiGKSLndiIIpfkaC7ZQLBILtCFjRdNImmJoWvwQbgw3R5at2Ieo3WPez6OwTsUaEsmzVi5LnS8XxfGc2izKCF2zrwH3kb8Yi15xGiDbfkbGxsby4lrEhRdI6F6QvMKpC66GGGExMpXB951dDWz7CYuu3sFCr3qSn3HYJGu0mifQux9mJEMSLFHgUXmoZUYcZCl2WDjdjZ2CdijxImJE5LZixE2f0d4Mh3aMTsR7FojyH1Hb8jcGGxs8jY3dW9hBCdFB9sWoUXx/P+CYwwwwxRyFaQjGsbNjV2JxONxrwLQaRmqeTgQGTQ/sdBvkWVt4Y/uQK5U6OvGuPaRVPCLK5UekujuwgnY0JExLmtWL8D/oHzSpQghFOLnkJ34P1xsbGxjcG6d6hYhBBD6YQ3UF9FQxaExhhhikhFhiBIJCPggkgkSkncakeDrnsIeLCkdT9I7HZY0exrHDyxi6iWLiqICQ3Hse2NHrYIPzkJi2ZOC/Deb5Xg8fFRZB4kdj2Bu/B17E0UYo/TG4N3lSieFzA+b/ANdi67LT2xIohCExDFJjQMHjtYweVBBPyR8kfIl+SPkgWR8nHyp7PtKez7z7B/MNo0w2ZNfQ1Z2USISISCRDQllCQioMdhohOE2ZP6vzr5J2IeKH0kJTsbKJlKNlKN69ZSnY88+BT4q/2X0TLz/P+CRQQsJjab8jRk2exk2D+Y+w+wfzF/J9x9x9x9h934ogIEwmL+RP8n6z9QvsL7i+4vtlKIkC9QneynsaMgWuiuT0erJxmT+uewmCHnBqiExsTy4bKXaUZ4GUo+F6O4I+y/10JBF34fz/AOCEIWGEyjY2xtlZWVlZWUpWUpXwFYv8PBRQnE5YnE4nE4mE3yJhP8if5F8wi9I3XsxZOfj+qYnwmIQQ8xOtsoxOFG8bGUpdb1s8sp0L0Vfihr2PTf13raEIQuhMo1jHzf5LnYk8SFExMTFggsFggiiggwy3xnCc6IpSlKUpSlKX8c2l4J2dQkQ8U6Cl9bSjY2UpSlHjFEFEELC+eh+BHz/Qnl86hCYhaY2MeMb4XaXaXaE+KDMCGIXBCEITLlKXiuK4r+rp5ycIJ3tbhT6KURSjY3pPKUbGxsXnKii6O7+8Uah9ixJahCFpjGPXr5XYIorA/wAEOPA1rtCDwNePFqYtWJ4sQn+XxwnCjKX8FKUvC8plx5cXsTLdpFLiEMbGN6mPvDdGyiHcLk10foToSuE89fz/AKe9QhCFjGxlxj168jEwiisj3hHrhr8rPJfrG6dEGE+sl5ELULViF+SfkuNlKUpfzThMXYp4nZjdYvFspQveXWNjYn2IKJ1l4qxjG99i1XxDWIQsYxjxsb5SiKzQTPA3HvKO/CU6R0qHdEPLs7nciIMu7CELZiFiFi5TEvw3DdK8peVG2+8VirjeFPPIvR4jexvZSlEzzkBv4KNiKNjDYxYQ8M3C/T9xEfp8eZ9vghCEIoxjxjHl4UKDuiR6Ia/KPMZPhhoPa2NlKhei0yKxERPrKCxImzguC1fhpXxoylKUpSly8KxMJH+RexesNxzwxC1sbKLEMbGNjeEyUTo/UveKSeyCiExYhCxjHxfBCYKQF4awmvpHkGdg1G89jYwjDEMkztdBCWTEtRMWT8dK+FLlL+Sl4JtfjJleBukhOCYmUQ2NjZRaxsY2XKiCHhFIvhIZRH+fyk9D1YhCxjGMfNZ7tHQmVtMuL7HRr5EHjG+DshYTJvKE1IuPKR5CE4LJ+Rst28GUo2UpfgpSicKUpVxuptC+QneaHmdh7wEUbGKJCEMYow3pe8IdEX/kxKhX+v5/zYQQhCFjGMY8ePUyBEuqIqhGiEIPsYfkmH9CYP8AhCy7pZQEthCbMWznR95S5RuFKy/nv4CbQhVwS70fstbZRMTKe8fYkJCKMeG4exKvJd/phYv6AP8AaMfFCELGMY+L4OMdYPIuGpjIsMg0NfAxOhIyHcUc+nlIIhMhNmTmxvbrfFSlKXlSlLwuLt1OCaerZOJsXSxMTEyl6xYkNjYxRqLyKIdAgpOPl/6H2feRFfsa2ExCFjeGPHyh0GIh0gto7BITKY0friHhKUHqYsugYwFk/FOFG/gmPGxhu7fwUuUvKl1MT26vR4tn8CixdCZSiYnrYw2NlPIUUrDRF1T0v945JexQS+BjXFYsbGMePnByTTPCNSReEg6NUaH9jQxrsfW/tPAqOPTxhNmzhOLfe243BvbtKUpSlKUpcpSl43UxPimeWWOoiiE8pRsTE8XKl6LpBNvRQC7cIh4/j/5jGhrksbHj4PijpioRIoeoSHkbjJ3BExB9jd47vMvgyUsmQmIhOTY9bG7xo2XaUpSlKUqLxTxVt4FlFonDuvJFUJlHA8CYg8FKWiYogmYhA38Hc6Vv0P6t1/P2xoYycELGx4/wrDan0lJ9iDGhjGhBNFpaZJ0zopuYNRk2EIQhCa3M8Dxsby62PKXjS/gJid4UTu3E5xCcG4qOvsrQ/NZNj+hMTpcU248hOhRCCDfeFMjyHtjWNDWTVreMevjRDCYxJi0EqOgxujPoaGtYH8CexswhRbvhUIQhODxuDePGxvHxbLlxuF+ClKUpSl43FupwvFMQxR+oNDR2hYUJwmKiY2N42LvKcEjwIfR+sj64Q8ZCcFrxsevkhMYYqp4BShDseeSDQk2+Q/IigndJsp6U1Hk4Pi3j1vm2N62X8t41iep3bngIXg7so0Qg+ihFcIMUp2O2EFPEaDbg9R+T4r8/z/OMeMaIQXFvi9fBMTGGLhYWQrTH8jZYNjF4+g12eoKNUIgnspkudxvgxrtx42N3bybKUpSovKl2lxPU9Ex+SUmyjDRYJlqon3jFEomVjwxqovk+Pk/n+2PHrXCCRBsfF4+NExMbUGoHor8FGywcbjG8N8KDoQh8BRqCHj4t3G8bhcePGxvLxpfw0rKVF4XgpRO6TExHjPZNlGGLBsLCgnRBOGYeJjiWseb9v5/sePhOCxvKXXjx5cTExhiLmNRCVBu+S+hsb9jjDDZRsusiMzthMfB42XW4N3g8Y3ybLxu38CZVt1OCZRO4QhakQaxohCoxM7R061lPHho3Qwfo/U6n7kYMeNDWzk3lLrx8kxhiwr0LV1k2UbGOg2N6yjdj8D8ncCIeHjKXG4N8H5x48bHl1j438NLxpdup3CE8upzYQhB0MvA+h+0+i6RBD7Gqj2/5/sXSg0PH+CDG/wAD5ITExyonL5Fo2MbvDH5GMbGM2TRImeV0NafeXHWsb4x48Y3xY9f9BeCYnzLqbxCxEIQhCanUaCca5W9FmMleifz/AOYx6xohNSx8qP8AEYck6WUYt7LMdQ38HRhjFH0Nj8UY+hkmUQ+B404Nxc3i74H+ClKXblKXgtXBRxohMQhMTE8QuEINE09Ozt7HJCHlahqzoPuJz+f4HjGiZBkIJcLzfe3ihsOREKsWEqofQbGxsYfk8j+RvrUHRBWoQfnHjV6xu48bxvrXt5Uo3eV41ieLt1OC0sJiE8RRMTReEueDsdRd5dV2RGezsMY+l0RX89kIQaIQhCbCMu0vK5dWIYYgyaMWosmdkKMdGx0fToxjxodYVH2IPH1jxvLjeNwbvBuDfKl4UbKUpS4vGo8jKJ4mLvCeE+K1CepzKhod0Vz6EKaITgYz5LT+fsR09IYsooorFcRrb+VYhhhywQKog6UfoeBlGN94/Ou8KJMTqmGNwb1+MeN43B9vHrevg3lKUf4U4Jl2wqergLilKJ3E8WJ0QxFG4N0aoghBnRj6RDdiv0C7/n7HiPkgx30vH6Mb18bwWLgYbsgdQLUWXZYGNj+xj0x4ZCC6H2h6uW7YN43xb269o3ypcvKid1PU8UTEyidxMTKJwTonRCcLRjaJnctidCRKQShuxh/pUn8/Z5ZfD3+Dwf50ISEMMMUFSFUyfTH5H2Nj7HhjG8OLdhI10PDxuZcePxrxvXrZ4xsbvG86XinqepwWKVamLELKXFo4ztkFhLEL3RvsW5vSo1vst/z+9II4R5s9NI0pcpSlKUvJCELDD4UUYlKpopIY0MYzwx6UrVK6Y3wYxua8bg+8evXxf9AmXKe8TE/Wt6xMTExPKtGHOKVEcLEyYPpH6k6Oi/CX/v8AnOhdkJSSBoPh8EDy6+S4IWIWGJEWJ1dDB8Q2Mb7GNjGMfwPyMTaYtBRD7HwuUbuPX3xeUbG+D4twsKUbmE6uC4JixC0hNC7KITE0UUbLBjsF2xSEJidEeRdIMWvw3X+n8okXhRM84xohCCxj4T8ExIhCCEMNhVdi3ovdp+aNjG6Gyj+sedQVRRsevH44vWyl1jfNuFpedExbS7RCYjosWIohMTyjg/lgj8skghdiYtP4G7he9F/z/GzoiKJlpRC0UY+DJpCMjIyMhCEIJCQkIQhIxhtluose4bGN9dDY2X2NjePE52XQwxODcXt5N+vwN/lRLFQnqeQIXCiKXDDkYMZDsQhCeXCdj7DfZQPb/X/7P7FKUTLRCLxmwL8v47qIJCQgw1OJWQpUYylG+xsYxjPOPcqecMeNUamPGh4yHZSjY3r1jfBvEFLUdcG4RZOb3UxcKy4nD5FwzR84gSEIWJiYmeBeqfp0T7+L+/eJlohOYnx8CHiQkIKxVyUnLG+HA8qaDcZQQWD9DdC0uhpob7GMaGvkh44JlEMMetj4vHjHj1jZSlG9r5UuPQ5E8yrGxC7KhcXZ4ylLRso2UzXQjvEJ6mIXQfgWFexCQniFixYnwmMQQ7hTEYEOhUSDUZmLRTDCuKFuQUYZDHmLs+kZbjcpDXZRRWLGbFRF42PjR9jWMY8fBkOh08Yu3jBIap4kBsWZ4cE2eJROCeUuXEhBdC7EIb6ExOcHYbnWUX6T+f8AwSIi6mLFwXF59BNFimMkpCEqI8CUStiYLehtExTyQ1irwRoUdsQhBuhGKKoiQi9jXwMS4mIhphoanCmKPh4xxjZ0VFG0RnyLpXopUi7S77xcfAfo7BqxuheC97RDZfZ5PthsTogksQhCeITmo8xusbuFX/T+f4xCEIQlcWrFrFlDUMHlgwVGz0GFfY2diF4G0hq+xIvYmw8s9HUSvyJF0xR5SnTOw0IEeuZIIJGpCEhIWNTwUbz3neNkxvopsTiYpWVx98OTpYXsTy9FOingeJJnYNUI8MQl2iid4WHfkXYggmI86pSiKJiOvZRCNvSO7e8QhCPAsWpbSlLiw7joGolilDUwgmPkFLoQh5N70MTf7C6TFmI2JMglBiFhwxNiQjxCYbMaEIQhCDQ1CDWOHgeLGiCRCDo2xsbG2emKFBOiRBcvAXo8ho4PUJi74FCw8FLgl2dCYhMrE6NixPEJ06hvof19sWlCaExIhEkF8gkI+SPkgSDQgggnhMwfClMxjyQaH0O8DjGN6GNPI0S6OleISISEDw0aYglHe/UfQ8yDoQaw0NZS8AbG8NlKPwIP5YgyDQ0OBoaGnR5QzwvGQSIQhGL0MaO8m+zqJ0Qu8JlExqlQUCFq8YmfYhPLl7KeQulB4j9xRKhERPsp4H5hSHJUJ16KhJ8n6AXmHEXkhEbuj7wSIITJJggOXsQ8nYVjQlCIaeRp8DQ6DV5/NDWYH8RloxhiTEGiIaQ1hj8DeKK8XaNobpCaPQw+4wyywwpw+AmJCXIO+MCAgn6E1con8luUTKUTTEiHvE6JzE7ifo/QTp3Y32PWkL3LCqGsYl9nkJjYux19DSMb1b2Jow/NQiiG2qjR+hw9tBHadiR1RBOILPIpHZRFBbQrOx2OxRNsQliMMNCIiGGxsY22NsdKMuNjHo9ePJxpSlxrIQNDQgakiBdFKJiQWofZY7BIJQWkzzi1OC7yiaL8Ys8i8ieplh82VeRbYncQ06D06jtg4aa8408fZH+BIJBJCIRCSEgxtlYqITZQmx2JsVKZQkQoi4aQYaDEoYbGxsbGx5NeMRDDDfZ4titGw1CajsjJjw8uUo8FqQQRYbueBMooE9TExMWJlExOCYmLJiQto8FOlFOxEhT3D2FSI00QhMhCEyb2KBhRmGRMQQSZGNCHYQwzwdBlhhlohiODeqGNpHYYYYbjFItEhJhhlhqNBoNEMsMOCEx4/scGPhTvi8TYsW5QmUbKP0WCxZ+vCidyibxFFsT8YTSUy9QSAheT3i7CT7HrtYiMWyINc5hlsURnZ2JMSYkxJiTO4ioRIThBI/gPYNsdHaRkZGRjQzuTKVlKN0qR5e9DQSNRoNBoSNRoNWNRlijaKiBhs+A3BjRHjsVjuODoSiXwSkIQaInZ22SjwXQmVifoURRcUJiF0ITExMTpRZROuDXqJEuhzKMbgRaPNCIcu8K1eVi1DYjJwtFRqhA0GhB+gX1EUW6EY6CYTiYTCcbFZCEGGGxRWj1NQo2MUbZWV4yibD10POyx0Nh+pRVPPY2zs/UnHz1rb9YUO/LJ0WFbL2JlFUeRM8ZPgZMonexMogmLExMTP1LBPSYmIrExVJ7huhOhX5EwS+GJ5M9eRB2jx4t+B4wc+hhmihYQJG0yJiQliUSiQaFtdEzsFlRRTossr4KLyblllFHbHcYSkNWjY8PoUYZcOilR9sNlHA4LR7WWnrPHKw8dixClOvREP4Hayo+yXyJC4Kjz0SeRVicKsUExMosTEJidxMTvjE5oSnoQm14Eo6DymJv6GgXmCV+Sg0YhnwDwuw0JFjdEaLBMhBZEV90UEEolECwpaEjVDQaD+A0Qamtw8dqPVSjY2HjHyabIR5fQ9TI9ZRvG0N4Yn6Oi9ETGoPoTLkQ0ysrOx6PJ9Zfg+TH8SJSd9brKJ4n8CcEgmITLiYmNxHVPsadHk7SExM84m14Ej5RKy+YbvB18MSUTLyJpiYSWWCdiGQaIQWJlyEIIQh9njG/jWG9fY/ONCVDfOPSXYlSDQ0th0P61vG9bHxTKNZBpDS8jOmPoJE6EJjjOg+sTgkqIvKF7iSRRHkbQlVTwLz2dehddjfR41jwVEhYhMT+SoQmxNFExOCZ4nySNL4IBr4GsQuC7K0xMhq84JGJwoyLLCwuiB9kIJYhCZCEITi2jpj6GtuPLlEIQ/wAYgj0LW/jbh8WVlLw+ijR4wmn2PrFTobKUVemR+SdCREjodAulOqf2EhAn9wheGJr5IfQp4FCgvNZ8L2UlSj7Eh3HGR1OzoXSFgR5ExUkxBy7EMdI9F2XZNPYTgi95MrQkCYX3P1CXCITETGIEhMiITE9iIh68oxujY3i3gyVIY2Psh48lKNjd1jdy54H9Yx7cvBnjwL1Y+lC69iI6xa8FCnsh0yBZ1aE3wNvgXfUILyNTwY321ejz7Oir2JryX6EK/gX0x2dZ2+qT7P0GKBqkdB8g+BCdafYJtr0wZ7ZCEgkRfI1+/IydIThYkYuyQSZBhrh65QZdjbLGIUxw6GyKHWJdiRM6IdouOmzZRi0bGNDDyGwvkQxssx550PpDZRvWUbx63j64Mb/A0Jx1Dr7CdEdmN30SlPQh2hp/IvYUfCk/I0IIRERfGUr+RNuesfmaEn8i69lLD1lmEyicKVlYmysrFfnmm0Ix8ksqZEMNPgkQhCCRwS+SEJCUe31kNVDDfUJBD8CgysbeKmdDINYhB9kHSjZ0Rj6EYt7CUWUb7GXhRu43jcG+so+Dp9BRXAQiCCIi+DpZ2V74y8GXPQujryQTKxwr9svZ5xdj6OpsITFCInCiKUWraJhK/IkYoxq9aEmQhGNtIxjfYviJ2htUO1rGuqNWrRpC/AjzoaGystPHePYIFKUq8D1og0NHaKQ66O0xuZ54t/ZV8kfI0IGj9FRfrHlY6fqeSoq+SH7yP4f7H2v2Z9z9j7n7H2P8f9F/Iv8ApT/0j6v8n1r9/wDw/R/n9j7kfSP4F/6X7/x/9P4Q/gh8rf4/4fe/5/Yj7/ye/wD+n8K/+jQu/wDbPYAzeEQklzYmLE+CFwWrU8nBZWhOJjsJSCoqE0dD7xsZIIkP9DtrpjWuhgY14E010UqpX6OhA3RjcH9H2PPgauEoS+xeIx63TyQaxxjTC+rGryJ9M/jIfx/6L+v8o+Zf5K/H7iZ8fz+x9yI+hXtT+cPnP3f8H3v8f8Pkb9xq8v8AyR//AEn8fuRfSI+EifhL9v8Awfxr9j6f8HwoifWfUNPg+iPoPrG30W8H8oW//C/yX7CTfJ7uz9TL5j+NLe0NXlD58P4B19sLro8vFwSPefR4IFwns8l1cFqYmXhZwpSsWayvv2H02sBV7Q4VCVVISU77I7BF7GvWP6E70xpov2N3so18FT6G4J1djjTE0dtv7jEkllbfct7jb7/uML8fflkITs68ns9A+cozvPSEfAz4bIvz+5Pv/k/V/wAn0Mt6ZT0y/ov5R9BcWRf2fOxXCb2yXvJfcTfhn04pBJ8D6BHwhfpGvpCX2JP5HfLZ38ij9nXyNeww1+WJ+yj9Dyj1/gRRvDrE3yJv3hBdl6h5JzWTjchS48urEnRKqPwQIeBXfI0vAZSiEPs6dDO3ZYxixw8dFUbg3TrwyteBCbwPWa/2N/G3Xx8IU+tI/QvydZBpMalhg69jBqIEU+jH50P6CV5F0BE8DR6ED/VyflgvXpP4ij7Q/UxsSMM/Qb+yymNiilDbfkceWNV7G3obPY2fn8lKeARa8jEngmwhRMrE7zXYuSWvbqfBt6E0fZPEessQ1danZ2NtDYn6PqN+y+y06G+xFjg4ifQ3t18ApIh+eJddBMVL+y3hk/JJA/gNsbDKhi/Ax/TIeGNveO3TYggheynlnwMj2P1Mp7Jez52fYNfsfyn2DmR8PNXKpUxr7Y0DT0P0HtBtv8b6KdsXhIb8i3lnpIQukOshCc6xMuXV0xd9l4JxhNfHyoxGvA3HY0eBie0Nl6PYvsaF0OQpCfQ4K7RNWxp9D1vV24eRdjd4Y3+z3mONn2kfZ9gl+z7T7hv9n3j+Q+4p7PuPuH84/lPvPtPvHgX85X8lssorKysr5pkiv8dIPoX4PCIY8ntM8Ej9DsXjsaDojWUvCE5Jv3wom0J08eC5eCE/BClGJ4Y0UfY0eR/A+jwKONHwDRLsSwOOxvaMQ2Jv0UUVg3ajx8GPh3y8lLjY3sRCfnpUX1p2xeIj0h7TEfIl+EfDj630pdDUNkvA3RGil1ERCEEuSxOCdFqYnBO61wE1jLGIUYxc8jqY2eRjQ3BhTpFsobFFDEx95Rtnevi7sKIyiiiMoooossssojKIyMjzsgkyMjGn6ysIl6xO9Cb6E/yfKxJ8sj6F4iEq8IgiEjoiGUuIpSnTGjGrEjDZDURlKXE8hCEJwScVhK6QhMmKMY8JiXwMbxsZRjEo0I8sDQhjXLqxYospjwm2xYq9j+wl+Rr8j+YSfJD2fYfcfefeX95JG79B9AzbvQ8H7Rg+AGhaiW+vgEj0JPo6esWWUWXlu40QdQ2ylYmNOnZSsTZS+MUpCR2IYy2RaI0S8VPZEREQkQU4IQhhMSpBoNTgy+svvKJ/I4MY3Rsp9hvGh9DY2djpCjHwNHgaHj+tYx9j6UIdo6CC+RRPHWvIXiyYsY0niMmJidLyuIaMYYeE2IgkPwJCGTaOETzDEMcDZFFkmdHQkshETHjExMhFBLE0NJkIQY+h95S5esY+sfFie+N71+SY9QfY+yZBjx8E8JswUpS6+uM5TGsISePwXgIZ0fXDNEIQhDwW9id1EIJcEsaDodjdDZbJUWEQ9C7YjKZCoiUVbNlGPKLGNcD61jEIS+MbKU9k49iC77GQaQ1RwNDRGNEyD0R7YpTzjRCDXOZGQmQgvkQYaJl1CELIMMwhCDRMgsWTsSPR0UXZ2y0EMZsvDP03W9QJEEhMIX0MeMm3uDaaGMYyj4NDQx94xl4t9lG6IaGqMhBqDRIQa9njINHa8C7HlMUuURFn0dD4Qgg8IQSoliWPyN+t86hYJ0WQeLD+GIQhDwJbT0UXxqeTgHi9skkeuGi8YJQmNnnWNlLHh4+i9a6ilxwauPxxueB98Hj1j6GeC9cGMhCldE+xO+BZeU1Cx74KJi4IQSITgmJC5dg0SEITHTEIQSO9RBbCJ8B2/AgSIQbylGy0Y8qE6WDHr+No3ceMeedfBkGhrWIeBkI9axkGhoYmJiZS9FLwmpzSlylngpR/kT4XKXIQRMNEKKGoTqiRPkmwgkJXkC+KLCcby43ENj4XHjGQqQ96x8LCn66xi14yDRCY+j9R69mWCwT/AKJZOa2UQkITncT4QhCEJxQhIRCEWKyYYX4Ax7Jngp5168T+B8H2Me3XyeM8a0NDQ+lkGiYwl8kJBYnlKeOcREQhCZ54QohRR9iZCDXUGiFEITPAk12LnCE4JHjKUp54QhMJEKN4+xvvHjYxjL1lQ3wf4P1H9cm+8fB8LjGeMhMhBonJZSid5oWUXZCTIQguaGJnkmIQg0QnNHnjBLsmLE9TKJkRBJiyFxjfCw8nWPLBu5crLyeN4zwdlL0Xh4LrLrx8JxaGicbtLwRTrELIQhBIhOaEJ8uiIn411zSmLFqzoXZBijKXKPG8bG4Pgyi43X8jPY+8T6Oso8e0bvOI8cJx8ZOVOilK99lKJ5cfYxcJrJixMpUeeC4rhLxmpcFlEutQusTKPW8e3Gy6xv4182NcGsY3WN551/hlPY1NhOUGoPhClKIvHsT7Eyly/mpSlFxnBdC4+cuLFiV4IpSiZS3bryjdHr17BcKWkHj8Z1wuvz+J9jJwnNohB7MontKdcL85Sl2/kpRM975ybdWonfC55cqXrLjZcu3k+Vx5ZyY/rGPg2efwM+xsv9BCZB5S5eylKJlTKUp9ngTKUTn4CkJyTE7+CixPFqxd5cbKUuXLw8DxvGN4x4y/PCseMYxjHy9j4XXSn2X8c2ZCDJMvGwvGlLlKLKUuWCdQsXKLjcpS5eK27YJ5SlRSlomXKWl263vrWPKXob+N74Pjcf4b3Bhr3/SzfOespRMu0pS5cuIRBcLtF+BPaJlE5qLeNLlxSlKWlLTxyuvfJdfkvC7dfXJ5eL7/AC+sT6PXF+eD5NEgnilKXKUvOwszzk287wuUTE7lYsW0rKXKy3KJ5YWlKeS+xsuXW4XqbfwXaW8Hry42W+B/mn9B9Z484tTKhcKJ4iiZdTxaVMvo6PA6eCl28LiFiLlLiKXLnsbLyZeHrLjGPwMfG4/OMfBnsfjhBj9ngfQu2Nfn9URP6GcF5z3l1O9beNPRSsTpT3woxZ7PZ4xYsQsQ8W//xAAqEQADAAICAgIABwEBAQEBAAAAAREQISAxQVEwYUBxgZGhsfDB4dFQ8f/aAAgBAgEBPxB0jEXY3WJo6ZmLBdDGGHmzQdsQ0DUQxMPbYuxsarQtYutJJCEIjRY+RbLXYkTyWdTvj3EPg8NiHxpvhCcWiYXPovyolCUMY4azsWal3jY+RjarxHUhLexCDKpR+ghDcQ3Bvs3f2F0Q0hUv+Br/AI/3EGhjWIihdDYg6CTmLo0IbDHzYjqJM84ro6YNhmpNEJRpwQ9gmxdYeQkehd0QkhNiaNGa4LF+WnXFjWFwNehZnwz4JZ70Dr4NpCg5vtY99ohfEEIJCWHldUTG4IZ3v0v7EoeoLf23/wAKZeHP2EEGh42NFo9j2xiXRU6PqHsNy2xg2SHoOHjcbiceBkxMx8GIIxHglSG47k3o3CJqJoa4rrFE6d3jISnkFiF2P7wXQ8XFw8N5fGE4tUhODQ1wnCcZxaZZtsmINwShzRad5v0NCg1RBIgkJecXeOxFihWuxs7PdiFdK/s//jGdN5EEGMeFoVH9Db0RXQ/riZe/QnovqQXRXod6GTor0fQN50X6G+hsHaUIh0GDvWLZMY2hoRju4Po8sqHdDkLQQR7QxUl2QSZ2E0LIubL894zM4Qa4Tg1wbBsTDYjsPWQTt1idEGGiCRBBLDHiNEaehrwdBTmj68Ehnknn+/kaEGIMawoZNjaFBSGTQZUEoWc7A1gm4eaeoWEwUJxpJRbvRV4fST7RMyacIskdSrZQfWmh1Kbg7oW3BShIxsV0XDy2N8H8NzCMaNlLxhOD1l8JhjYQagtBUiiFtqyEoGhoaGNUSEhIWhjomxISJo0GPzy//RNCVpCxfCr/AN+beRiDQyFBh3tjDSdkvFqJZK7NCSR3sdeRQzZnXYkQkFHQzQoW2OipsujZGiwXUUhBCdnwnLN4foGhbQvYShZE+hNjDiw/npR8bmBqcHmDWYNTm2FCUpG+httmKFENhryIINEIQgkQgwlvBIY47TtP2K/sQp3loRB41/3gEGh4rsR2tm2SkY08FhU87kphIJCZWUoorwbQ4haE1YJFojoUs8NFQ68jVlGNJiBFSYhjViYrg+b5snCExSlOyc4dZa+DhFs6EhIg0NZHsaEsJCPrEEqINDaHJaKfa/6E8l5+kf8A8/6Wzy6NDwYg8IpuDUYrGscEYtMI+F/ssbllFDZeCy/RZfkpEfk2jZWUJxFs2JCQlBZfO/gpnsmYUuIJ4afF9YeBNMJCDWh5GMhCEJwSomuDaOxKNCafr+xdF7+i/wB+o8Hg1g1hLCSEjQkhJMSCQSECTPOEEE54yrjliR43iaEkEDwsPISEoIWLzv4RkzMtUanC4ay8H2aiSCQkIJhoMg0QmIQmEEsbo49U+dYpXRKlHs/7/wCoY+hjwTKlE8dnQhYWEXCz1lbIJERENBBDIYwwxCYMVxB8GUEEXm8TDzCZmITMIQhONLwmZmXDxmxKoacAQUY8wZKQh0QRcXGohL1sUcu8tCmq9Il/v0g+8MYgxomEIQkyCwiCQkQhCEJSEEipDRDnAwXnhuMaGMazB5aIQhP/AMGE4TLKPZHl4eRISKwSGhB4saIJEEIQZKJC6HoYceuH7DiCXHVP/fwfdTwxoY+JCEIWOxCEsogkQ0NENcLcYSsDZsTothoMYxjXwz8bOcJmZmKNeszN9iRubETBMjRCCRCEzBIo2sWhKpJH72JaLB6i/wB+w+sMYxjXAhZQhcFmEQNUMtx/cSsDdm2N0OlWMbLy4bR4vB5fxX8BMTnCD4zLXF8qU2ahJ4aw9cEIQSINEGQSolBoeDzHuhIXoR7Qe/79kPDH2PI8IQhYSEIWEIcERlliEegbs2E48IR7PHjY97FdEimHi4eevw0IJE4T45wmZmcHilsmhai0GyQ1ig0MQglohCEJsSEhj5WvzK/sLolp5JZHS/8An/ODGMaGQWEIQhImEIbg5Y37jDZ9FbE7PCEB0SEiDR2EqJbmuEVG83LxeF4P4oQXKYhCQhMIya5Tg9YmXhnTBNlEEhITgmQhCEGEoMgkMorGrCw/SghBenf23/w/M6ftoo8MYxjIdCEISEsIQiYuMU0hOxx9EexC1BIIWtFmew2tnXFjLaeJh/G38sILlCE+OIi+A16LwmbENzZBcD7DQ1gkJEjwxoguh5UEFrezYhXRv8f0Vvb74MY0MfeELCEIQuC6EOp7CyEwiWKNoWFhqNo2VDibziuaUvG/JMTkifgZhJxaxcvGCixUS4polGoQSINEH9CWTwjBqMmoSsvP2/8AH+TL4yxjw8kIQhLCysJVkoMg4ONwbExMaMTKJiYwxqx6tG8YG8DZcX4nzhOS5IhMTExMMhOHZMvDV4PBCCRISg1rCsGJE4CDWErh4aYPo3exiaFrF+jV/wB+twncMY8PJYIQha4LKiikHDWdBxiFgmfmIQsDaEPQdYbyMW8bzvNYXBcoT54TMGsTDVGpwJGxCQkJwkrIJCRINE0NDQtqQaHwQaC/oL+xIsHlob9TX+/Wiwnl4Y+sIQhC4oRTdCiaLHhSMQj8hDFrwQuhaHKqN2zRxkBYXFL8y4TjCfhJmYaGsNUmD2LTdRITG949hiWyYRCDQ8RCRDHgyY0GC19v9CFV1I3/AL+Rrn8sYsXgx4QhCFhZWGNncU1ZJldM0F8HSExMu9YJi6FgxtgfQTxJ1D/GwhCYhCcITEJxmYMaw1lYpQ2Q1EUOIMTBISg1oYY1iPCxuMgtmwyamUYtzvT/AH74XRS8WM0ISyhHjEwsHk2GjOkbDoTgtMZNiGwWCESZuGzQYb4L8K5zhOEIQhCE+GcJmYaIdMJrGpikKxCbJBBLBoahsNUkGjYawaQ6r0aPsxYhdlN7e/8Av/cJ5THhjHhKZQuSGIPsQUSEiiGYlhHWxsiwTE6X0LqMOIvC/AuK4TKxCEzCEJyhOT5TDwSrLNIZhKC3kJsUDIJE0MY0SDWCRDwTVKsZKj1sSgxL5aRor0iWU+THiaEIQhCwuCCCEDzC08T0T8Hg6EJ4VgtiGNUyNRRBhO/gVwnwL4JiExMzjMzDQxoTZuGiQU9mj5CDIjcbITGfY0IQazvRMKbErP0Eqbb43/v4Pt18LwbnBLCEIQhCFiDQgghRCQkWWxbRCQQj0yLCYi+mHrgXlMoXBLmkL8JCEGpmZesPC1iSsdjQ3BP5ImyxE9JWLUQyeDQkNmPBHqRTF7e9iVZq/hRf79i4XCl5EISwhCFicBBBDyC08DxiMWHY6C3ktDEjR3UawhviXFc0hL5p8UITi8IJs0CZcQQdQn8mhD6Eo3axOCEzMTY1BXR23FiF2T2z3/fsjzhD4sY8QSEiYQhYnBoQQQshI9EDwYlwR0S3h4wmEIqhssGo4LguaXNcFwXJLwNQRIJQhCCRBonKExOKzRoi0pcoIIR0JCl7xtI+DpVkg0SkKbsEtCOOEOnw/wB/WFjxyY8ISIQSwhC5IJ4GihIaOHSOkJCCQkIJCEqQWCJGjQ6UfEXwLKJ8MJ8jxHxmZhiRoY6KUTLlogweIaOhIlEslpEMFoSEFrEs+hH5yX9j8xv65UuXlIRCEyhcJlBoQSiR03muISpBJiQmhCQtHnCEVVO9jUexa+FZnNc4IhPmmITg8s7CC+cvFhclyxOPCxNBb2btH0jUKBGzubjJsRM/AdfY1Bix0Pi8JCRCEIJCELE4NCQQUl0ND7yQ+hKCWxCQjwLCEJVDvQxsJ3gsLgua5pYnxz4WuDw+zsUlkx5bKUpcMsHp7IiWzhg1MSggUQ+1PP8AfyecPCeZh4eEhLExCZSJxYmCFNCRkHjQhC0IQhCNZqqiqpRhcFxXKCyspC1zhCc5zavBjEJuiREy8sZ0IJlo3nVmwqCWJLSxRgmhImxPqlX+v+eEMmLweHhIRCYmULj1hoQ0F84IPZvOgQSwkIJZSwh1o7yLJBcliZnCZWUIn4OcphjIdR0QhMNc6UTFiySIjVvTFCizRGokKE+X+/6XZ41ilQ3mlRcUpScIQhMQnB4eRKhBppm0g0Q94ghZQsIRNSqLPhQuMylwX4Z7w+DQxiPNxeGiDwxDEJCQkNdFxtdDV2Kc0O9glsTD8JSg3l4XJS8BGLUkRMwglhcHh4YlQg0KKSNxVQgsLKJlCVHjZN0m8TKWVxS5QnGfHB/A1xeF2SExMNXDGhohDRibEEhKHQw2GI0NxjEIR596/wB+43svwgNYPiJE4Tg8PXEglEOnhd2QUSEhCQhLCwsVRRYRoWEspcYJcl8cOh8oQnKEz0JTYSRCEwxrDGNYSwQSJoaDDFxZeE236G966xh8FYvAvG+OhLEITkx4Y8iwQkzzlNY8C1lC1lYSndSO1hLCyuK5rh2TnPinOPJMCWiEw0QY0QaIIorAkHg9GHikY9kGN55aQmPwmL5ASSNBhh5CEwgkQhCEITDy8mIILTNJOEghYQhfWULFVSqZZlYSvFZmVyXGfMycWh4THBog0QaINYlEhYEEhpgbGNDxRMP5l/v4K79lZS4MUVhRRtg2KxfG1hwcHhjwggghJkciEIQsrLVUIUjsScEuCFyQvwEIQhOcyxkLM0YQiGGhogyYITEOkUG4DRoigJXBOU+FB6YNEGhohBqEJhoaNiJlE5tlGMbKXGw1ggtM68gLQhCEhZWZqWRRlLisrjRYXxznCcXybCSJiDQ0NDQqEEhCEG9FBvIxoY0oatsor9bG/UdN2Y0QaGhoaINEGsNDWE0QSSJCMjEYNCBsbwNjZthRBoQSM8JRE+xLCWFiYWXvH5xIxC4rheK5LjMTCGsomZRYHA8sSvF1wgyEEEsEhaGxsYxjVJhlggsI8V8Ig0NDQ0QaxBoaINDWRsMKxTw744RdYMUsExMYxiDxlx/JIISz3icaIgLhoU8nXK4QuFLm4XBLiubEsnSLMPC74jzCEITDY2MYx4ZBnQiH5mdKBomDQ1hohBoYxjw2MxseC+YCwbiFHAysdGweybxZwgxOi5LjuLKjExonBcpilzc3CixSl41lFnsUVw1CYZNiVklnshMJj6wxhvDGxjGhoaw9foS2dD0ixhBjVGhoaJhjQxjwtlBoSIbwWCwUrgKJjQ1wNibExijEKERIhpvEnJUXFLhSiply44bohCE4BCMoog18C4wWNsGnmlOwuhJCS0d8MQmxIswSJCcGPsex6GMeIQaGjRDUElPotXt/0NWY0QaGNDGsQYxjHhhY3Y5IeMJCeBIaIQkEJkDroaxCfVGr8QTehgbob4vNex00JlQVKgm+ApSlKUaTECcTShMJ2JsiKzp8ArCskIQhETCOwqVFgy57mxqE0d2MbyFmCw8seGPYyExsaINFhkEib4QwPLGPDwxkHh5UhBUWiDGt4uEhbDRChMy5RWH1C2TSgT4CfAtKKDGUoLmmRXP++hCsGNlB2IUEISlEQJRKJg8UUP5T77UwnBYjZCEE3h1Op2HhMkEiHQ+x4eGNwY8wmWhognsNCVftj9feWNjHh4Y8MaGiCkgiHSJkmdxlExPyJ3FHkveJ0ixvFV77IDyErsUx7Ehs0siMLBMggkSkIarZI8A1Y1I4JDRYoESQaNQoNipiwkJEyLIYQehjZWdkIQguzuPo2O3AdFvC+xvxxYxjWIMZCDQx6qIavWr+41j0NMbFMaNhuRopjZFFDTKGw2ZWKiCCGjdgaGWCCC2KKuy2xKNo6ZRVMUWw3D68H2jHk8oa7Q37EPiUiQkhH5hrjCH9sGvsf2GGw2Zb84PsF7sFixQplZQmxNiYhFKJiFExMTy8GPE5dyBBFUMMdhOMkIIYlZfQ37PPJoZMtExBoZSZP9im8DUZB/Qs3KY3Q7wdGxJsOe0NDpZonA+AKFmMtR+pZGitCZEOxSX7NfIo6Y1+RiuG2wqXkfuIeSsOxZZZXgbDcbjKxN7wpiVyT4EUwuAJlExMQXIA3Ch88ppLosN0TFKE4xCkXsU7EpJBSnY0QZBkJjyNRj3sRb+FR7olsSgzwJ2JsSW0JqjbIJo0NJ2wzqxsIULsZTwXEisKUKfRYhlUGQ2xYstEaNlaExWDCCbYnBQWMGMbs2KJiYvgUFQqFCrIkEhUVFKXEITKeLmsrEGxh1keGho2MbY2E49DB6DN4SIhBkxClZaEwmRYgkEwSM1wmYR7xtvehjtBKUYvIX5E2NFTIha6Nir0VFTw2ssY5jAzGY2EbCQiww4JIcwY2XBt3QlrYiMSEhMxs0NQ1DYrCiYS4QSEwm1w+tipWILIvFFBCjwSLGigniDITkZ5hqLQmJjIM4QhCExMnwV0kJkJxKJgmeS0eKIGsReYmjGLsShMsKVlyJhMvwduiZLQoKXY6G8Jjgw2x5SIUFZQtgeuOhJJIhEhxhBNkQgsFEx0u8TaZSlKyigRRRYbbN4TFvCYmJlKUubiFDOxDNTXCQR5KdjDbsaIQhMGlg8wapM1iT0xd4SeGFSmzj7Eq0NRm0hMKCRiddYUjLRCwmUbKUogkxoiNIbGQ2hqSIvXC2ZSiaRAkSJRAp5PYGTg0Gpvg6KjIQhCEIJDTIRkYkyExCCwIJCTYmZRReFFm1igmKLFm0IJCcJFPZLGjwMRLwRjD2ijIY72NkNkTLTJiEIQaHk1cNGrr8DdQ2bF9h1Q1gy0YFUJsEGHtEjIyUpuUJhMUJiveLLDMiaeHQYaDE4JlFCeCf2WUUUrKwairhJ+PDUZ4aIggpikY/JYJCCUk1FEKEQ6NOiigkmP1GmmWCGjYY2ioaTKDRjfA1DInQilZAaMiNkNYQmYQg9jUw0NCp/pEtCoexGjZTTGjGW5tFEyEE/sTiPIkJEyZuFcI8URBl+w2XGjRSlKUWHXh7KEYoqadkkogkgiNLGicJhCEWEwhOESg/orNMV7CaDoaYVCRlo76GyE4JROjRjGwirNM3G66FV3gTpDeEtnoG4eYR7EvsQn1D0CQw1iYY0MeH2L94pYKCbw0hluOopC9ysS+huMRorQmRBIhDcwqYmaGh8FAowstlFFFCbfCEyuCR2Q+hspRRjfJE4plKXioPYRNaGmjoVEyH9MSQ9BIJtCCaY16DUVoTCRmmNX0NXgrCZdiRsrCR9FIVx4wQKOh7o64bNA9h22Ol9YBsuxkINDQyUaeohK7jsaGeCmmNGP1GxWWblYhBiCiZ0whMuy/GhkFrCEiCysLCY+zqNie4bockR6KJ5hCC4TguKY1GnCOxoxKQikR7EBqdRH5FoToQ+zxxsuhshVDUNdjQeobLwJhIj2INI7GQe9HYReBOoz0bR6dgsIgI8Bdkd8lWmJX1hMRmDmaQUieHiaJvFw0Y/QbonFUo2WieFiovK4mYTCFmb5IkygsLHiYceFyWV8PWU5gVv2OKiv0ImtoaXIThBPT6ELhBAm6ISe8ZDpSloJMa9IXkaHFJjTpiHSZTQX0I3gTIhOmQ/I4l2NPaD9Ukoa0xecY1CXqxg+jXsbQ4b8CcKLgr9GiEpROhl5In0RD7EFuGqxo8ZbGlBqxluiM0XGylKXE5vgsplysLihuhfOGIdjdGiEITCEswnFcJwQmTqHr0j20adG2O6ynQjzhVp2M03Qkuvsd4heVo17QaO0GxdGQGh7qNHkNPQauhr0p6EH9QwNxt5Dbuu/yVvDEn8CYJ+6IMYRMbrtP3X/ANGv+f8Ahv5f8iCTT/s/KXJUR7X/AH6jlW8JpFAPzP2I0hwYkEM9GJWipjU6NlYh2Mkw+stjecEiYQX2IYncSkIQ6xMJnsmLwhMUTxoubxaJ+JUQ1HuKelH4kHubEYku1Sez+S/hn1Cd6F70JxI+5ntbIeP5IeP7Iho6S/YUl3L9CrpjQft9jps/U8oSoR3jU7EgSsBjxj+ga+ht3D6CL0a9YaPOGhoasYbjYaayTTGhqDHhsbKLoZPIk8NDWDQjoqKUomKDS8EylMzCXC52KHiDVLE+GEwkLK4RsTvAndMJo+kX0jb6PchHn+AvK2fcxGkifgT0k/Yi9CfgrGMeZjyby9kbFFLQefM8ERLofoYsxxbMrMIQhCDROQ1hhoz0DZYoeBujaGx4E0VDSEEUbRHTE70NNFZdFuZkaJeiexN0O0KJiJicUWCYm0xCLMGiEE3gTekJ/TCcF6hN7gvO0L2C8wXkbI+xI8CR4CR0gkdJfsXzBspSCD7iHlH3n3n2H2H3CZieKn0T6H9T8hXoXqLCngjeBP8AH8CnX0ONLbF3HEIXbEkuvhSithBO8XwY8NUY2UceWQYasfqNhuMsI4MIyDQkxCEwtqhu2yCT8oSc36Py4SCpBP2hOaQWvA4tWAQDbshYRBohGQSEvWCHsOGw0ManqMvYh0keEv4J+MDX0bislE7oetF+kW6X8FfH8Ef3+wnez88XkZ938lfP8i9ovIz70X5C+4vIPtPtPuEnBHwIDd3BpO4VPRLyj7kehn2/wR1f4GvY/DWPQkd2h0i+OobQtGGmi5LxeHloYeCYmGswMuxoPHHnETJC+8Q2zrsT0JwiZPEXQz6Y1dEViDbbwIaF6BXBA84bu1hISOvXErpTUGv6Ek9mw57LuEO09C2JD9Ua7E29CZ4R9CPpX8Gvr+CHr+CflfwfS/gafA+uNHAoH0UafJ+aOWtjI4FvpI+lD8UH9Y/Yht2G1axG15Hqd/qN+mZ7GJnQ0b0v5Gg08I/KLbo8AJ9I7wKb2OmXO5pAyy4EhYaMYYazcly8XLyxMQa9DUy1BneehNMSGJ+8QSuhpp6NtFdDdoTa2daPSi2GUmOOinvQSu+xUsLPJCxBd0qxi7RVmgiQqY2J0ICZdf8AwTJjV5Zb/wAY/RsaBs/BtisS6EjwbBCT0dxnpZE8A0nkPGR6ZbtjT7pPc+4l7CY+5vioIpCJdooJELwDwY99i3eOXgfBSmg0QlExs9DQlRKQJFijSGGhSI+NKUpRv4WiIaGhqEy3os7KU7E/AlGPoaNmkkEjQlFskxvDNlKJumRuVp7I2Eo2kEuU8XZR7F2oZumjZiH0FEfohCDQkRCE9HakXgSPyYfkNPAmfgteC3gfgR7EfUX8H0l/GIvSJ/YXgFx6JTwL1DQrBsnGhrpHgDtBfbzxJXSOsXPRcdFSGiEISP2DvBXbKYl2JtCkQXBqijJehhhhojN8a8UrKKylKXDZRjRBWXEFU5hIQtpkhGw1UeMTglVKQYaEhsaRQednnYlnsSmHJUMdfkb2JsSDMYaXQ2+D6yvg+nnVOvUL0i9J9Z9YvSfST8C9IvUfQfUfWfSfSfQfUfQR6IvQkicG52CX4ErpEnw1EIaBKG+Bgz0PyheY30doaRoYJGJHhm/BQs1IQXhBhiEINEJxo3l8G8QaGUTpCEKzs6Z9oTYxuM2L6ZJsSbYTsxYWVRDTRn1n05dIdQsUTwuOhPjCEITMIT8BUVEIaCUNRj7Ro/Ij2NPY18jKSRQUhFtR0YdNoe41L2IQSsqNYryorOjCobREQg0TMQ0QeaUvAjohjkah2JVEJSQlJBCdxJGSYtvY6EcBJAkIIEhBBBBBBBJJAkIIyEuNR9hPKgJJJ5ABBJJ9hIYNXkah+4avI0eR+4fsGo2Z9w2eKyxsPuPuL9lCcsovF5QU7KC7yJ0MbeHipdisgqL8YgUfFjxSjy2VDjoRWGiQTgncQSxoJCcwKYk9BzsJ/gafDPziDVE/yxn8j/ZEepHoNJBI0ICUN14FehaDRyjJ+RkvIlIH6iPQn4GnwfRh+kfNjEkIFgl8Vk0O2eRs8jf5LeR1+SYQhCEzCEIT4oJHWEiNfkXYRv0xprorQpzUpSlF55IKilGxqlDWW86whpLYmLDVGvOCVIJYhBIWsC8gkQ7GjRS0YkPAzctGwXBaJtFRF2Ie0JEJliPAxDLRCZmIQnO4pcKUuITlMQglRIoZojJifDCwTdGMXeBWJhBWVisykYQQUuCdHdBBwbDdERopS4Lse1lCEpIJi2IkIQkIdCgssJ8JiIvA7LBBCgxoQ1DplEJnY8JRTFPogNkMQmIQhCfDOC4QhBI3IxcFCKcQgaYGhrE53jSBOhMhFIJWbFFyBQXAJD7E3vAU6FjGxseCCO8J5TDQTwuEENMK7LntEGJiDYmUgOhhBu4KcMDrBQOhso0eFwWWURkIQnxXEKKwlSsKJ+ycWaE4mQ2xYe46HxXJvFmLjvFE3gTLCilwobF5kwjDBBMhzLlGx8xOmhqjQsGwhZaEoJ4UELLy8kx1hSlGyiZCvFhCjeCYmNIRg2GGGphv5IENECRA1gylZRWUUUUpcUpeF4wes0hYXjShOJs6V4J0IKRFYjLDZTWJwQniUcCohhYhOLeBsJ5eSZS5UpSCxSlwxi7EIpoiGgoQTN5vKIYSSbDQYeKXDxeF4UpSlzcMaxSlLilKUuExMTERMg0bKykaccuCxDofYhaEQlGEJifgWilHwRASMTLijGJzDYYuKUTEx4uKJ3EELFLgz04Wx0TiuVKUpSl/DXO8tXCExS4XAgggjJBUPDRCZpRdYRNEOsITw14JBYomJ87gomuD1ijZS8Ux4pcUTWE8UpS4aPAy5GoT8O/gfyPMxB4X2UuKUvlhCC4hFkWE895UFso9kFwoncvCZcExYNEGuE4obGylELi4UtKUomXDSeRZhOVxRspfiZSmvhnF5pTQ8wnwtlEwikEylFgpRfAhMosTCxBFKUeKJiYmIdjGmQZOEJ5w1l6LBBMTKUrG6NiCZS8Ggyyy4zCfBef388zPjnwQmbCvOKFpSkJiUawiUghYuYJcdkITKYmIKYg1h5WN4hBqkzSl5UomUuIRDLLLca4TnS5p18VLxnxT5YThsrR9iokIQ6OyCEPihPjPOGLEw2J4TExM7GhonBTjCE4JlKUvCcFwdDLLLYa5Tih/PS8ZiYhCD+eEIQhBrMJnrE5UWadj4NDEJiYhCeHiHWViEGIhBrhSl4p3Cz3waTGWWX6jE50pcU7/AAc4se/kmfGJiEGiUnwaFhYQi47zDrh0LCKUonmEGufWYNExRPFzRMTxSlzcMdDDLdDgaITg4uxFHvo30VlKUpfwj+OZmYJEGiaJymJzTLjoWKLhBZpcJnYuDcLhMpS0TKMSGspwpSlwtCFLClLwTxKMOuxqMsQhBohCQ6zSlxS/P18kFiCWyTExCDQ0QmJ8qey46EPvEFzWaJ8GPiuT9rNK8KiwTLilZdlwpaUpcSjDQ6GXBMQmYQeUXN49fKycULCxCEIQSIMg1cPMKL4liYaJxXK4onhvDxcXgs9kx5HilKWcAnwpS5uWhsMMuRiEITE84nClzcIvGcoQhBEIJCQsTjBK4tEGhonBM3nRPKL8dwiiYmUuZwuEylLh7w+FxcJlE2JlLlMpSiy0MMNMWWoTEGiE+ClL8sITEEsTMzBBIg0MMMNDQ0ThPgRRdcHh/Cs0uE8z4euUJxuKUJ3CfClKWlzDYaHWhluTEIQhDZsguNLmnn4ITKxBIpkhETYkJXhBoeGRP4lhCFm58Zfwp4WFwfH6GMfB8lxomJ3HWVlMWUMaHrDSGkPDXHzmi4o8Y8fN24wRT//EACgQAQACAgEEAgEFAQEBAAAAAAEAESExQRBRYXGBkaEgscHR8OHxMP/aAAgBAQABPxBACx/aCGjDlJSea6AVg7ygqM33gaJ24gEEanzYuaOOLitK9R2nnxLnxOe4fCpTMUKl0Y/EeNzF8xRbfzMMswzOWHiOkxZORy7xLG9y+itfUxPpmV+IdcSlt03iJcNH5RE7xcZ4Vxv/AHaLVXgwy5as51cqicZ+oQEJV4K05uAmLFNjBExZKLQGIHvMBO8zuKviX/BHjFiLEe0WosYWNiXH4Rbi1FixY9TUVsToxYQ8Qwwq+83+g/QSpUIMINwh00hDcVVHNhSFtJGpMYjcrUmo7t3CakdhO9CokfMRYIlJxdlEHHlGyA8XEIKKL7Sy8x534g32irS68TPkzDSqrxK2VrtGwfnxFAzTllbsleKz/EFCxgSjXFNJ+P1DKY/5H7Mtl2ZeTPOSV7SyOGkUrKhweI0ZY55+oxN/cLV8aqMpWuCoSB94hVyfEIK/iPMDEMNt4+Zq41BVOehEwI7NzYIa2d+I1HiJsW33huiviVomaBqVYLxMin9oUriualA4g42269sypxf/AJMCbuymIwcv4j3uzjBDlm+bOe8IIZ+YSjS1nue4dL9E3L7YlVnOxz4IkHOZnGNhPQw4mJHF0Cy6ixR6C3GGMdxbixjqL0XqnQ3NyoMOpxCEIb1B/QMGbh0NwemC2EcMzo2Q1SYWW4DSVQAywTLGVB5n7kEcFfJqZ2x9IXBXxEU5+I498zBb8xbzxcE78zYvUoCn6gwlcShFZNVGTDTAE0O2BxJR72/8mI1KjiOfg/JHYwafGJdc8cpXo3xVv9R1n1CbGezNcYitw9sYBzzc2FG69wgWhXfEMwMO4qxuY3uIW5eYx3ieV3HYnCj25lwwl6rwHLCncSBc1DLKzmd/mXqHFSxyN8zeCnzEYh4wwo2YOYjK8XxMbCAe7rEIBvzuKndY1KRVZ5hXHl/Eeza4zLJq1y+4MpzZggU855jdzdxGvEwr8VAUFGmnMupz+8u9yZvxGo8XGC262VK/R1gqIsWMLFo6C3F6LF3jFixYx3Fl309QzKuVmV0O8H9A30vMHHUz0INwYMJcINS9zKbmIAZb28QkKJctQTZiKsSvMHAL9xNR8uI6QduJUK+yYN16irhtjQc/UuxL7BshXfEoe5gAy+paG8RCbszeIYhu4EBoFLwxa0OpamwV/H8MWboIYMo7sL+wB8S4t2zdN8sJgmd7Ry5l4SHmCcu7MrvJplNTXtG7dbj4jCXnk2sHgC63CHcrh7QTbf7y9XV1LwiWJCXMU1zqWD+cqFuPdlQcfMOhV9tRjzhrr+08tinVKKjrGi9Jf4ivqRyNjX++YSr03HpfKxl7BYd9wgMbPqIDhKmxgqWDvcMbFWzcuSs+5kyPFdnPiOMV+0HBn1KYHFV+ZRQt/v8AiZh/cMlCntOIVFmnxBHDcBSPMs7WJafUsq+kWKRbij9ItxYvQsWL0LFixY4jiL04lneGeldPnrVzJDpfU31NwMf3DHQb6aS1m7MCrZWFG5kGC3CDYSvMAcb9xFUDysU2G95Tjk/eBZz3hRq4xZcc9sw/8loLuCL2l2Grr6mi38xZF1L6uiyIPJ4iMgq4aiEUTi4Z6hUO9Z/NxjyDdlgfVpGGCgeDmD0Nk7kys7cpjBRXx+08pNYIe2ZicltMU7Iv3ArRTVzwznECVSBmVSzfeBKW3UzrokA0bKqFWSs+GU2UQBRv3zDj93+4GaavUYMWxz5l3MBbw8wKWE1NzRfErtmiINrkzmW0WXWZTXYNmohWjz5jQvXBdupz7iyWBv6ZVat9uY6OS5upQUbvbAvUxzLbNne5kZvywGqPAbg5dnIsEwqMHNQhfXL/ALtDmiEFhau0F2vmUQWzq4BG2cVtmddjcwxgM9AxcXoLFixYuYxaijGPR63R0I4Oh0CJMyodQ+egQh0MwHub/QPSKYEoWCOTEFwLhytahdM9RNd+UmjJ2sXINRWjNd4iD53Mfa9XBnzMTm8SsrcUx37xtY1KjfxKaUtvMyefMJN1CDdjmOpq4yjC5O0sMsPirZFcGGOMsBo5Mi/Y/MqgClT4/wDYcugaz0rSbI5ylyPiGpaPEMUxrFlQZWoypvfzFp+UXha4if038zBR5NkGo+Esbe7jUspehrUDSN+9VOa1+tQVFY84jI0vk1L7Ia0wDAc/EDT7pLeeW9+oMKy3rEPr31Dt3s3iCKVrtUcK16yQO3CGYu9jVkt7RCYDqKGgKMSiDhKQFRqvEwYIrZCqxqt1KzTu+SBAJjVMCSrXZmGoe2YGF58QgtmTUuIWLYbuO+ZVMLy+Jdt1faaxB7Tyn4iodCyFiowtxYxYsuLFixYvVMR1GPqcQ6fE3L+Zd9CDX6Lh+YdR6GYNQ/RT0YrY8XNMzyyzs9g5hcJeVzG2hPdD71m9y8pwsC+MylO9QfdalD/2Fz3+0UO6xOwS2xahe+YULxA9yAo/aUUVb7mc3NXXyRXRzUpSSkbTMy6wC+7j+YaK8RyFKquVn6B9INPgQd5hgmDHQyMf9qBrD4J/40Jx+p/ysBwvtUKsGnuQX38QHNfqLx+MBiROKS1G3xF/+Isb+WIteHOpTv8ACKN/tlLb+E0X+EdH7YscfYVHucvBHly9Yni/Vz/gkTGsHsVKX9IuFsjybgRw+tzg7ADh8kVaqviXi/Q0ztz43CE1NrtLOjejGJZOPMJWz8cf64ZFRMeE0mfeoKzumIcaiq/jHArntcQ+cqziFrMcyuphKSURUuLFiy5cWXFqLHcWX0WPQf04ldK6GZcP0X+ggw+peJfWpZl8N2NIRaDn/ZII3dm1zctzkzp7wz8PmJW6XhlnLmO6ax2ioO8xL3U725RXeKvtAcckS6tq4pktPaUTG/M29cytHW03LGkt3HoM50EaymqeZWWlDMuzQ75A/tfqIsNS6zISNu36D6nmhcO4eJZcGdTFddL0iq4dqls/jOcIbh9Quw+prj8QP/E7P0cJIXEmXnKuM8eKzw6yzhHsonFPqKcH1HtoprD1PHHDgj2InhHsJ4FQ/wDxH/ZHD+sLN/jLV/jE8/jCT+kvZH1E8PqW9PqXbpXqeea7TyT6jD+sIJ+2AcMeJYxK8SiAUSkO0oqadV6FxYsWLmLcWXLl9Xpcqc//ADzfiHUL6V99Rgwb6XmODKIrzcTbeJl28Jnm/wCpgS6Kmers73BqHGYlDnx0ngJfbuIMTajHM3mIpFHGPEU87iTTjiX8RcOWaFYhumd2TKyw6YYeYyCrLki8JEriUXcCvv8A9IuLUAfUIAJUJ2lYX7jqCyDKQQYgmzEVfMArEKtdUCB2IVQgBlPEBAckpKdp4I6Kngh2Jg1PBCrU8EB2ge0p2iO0R2lL1E8kpaj2IDiPaj2p2HVC7FOxRl4R/GL4/U16/U40UcdNrlUJfmXiLLiy4sVFlxbiy76al11uXmWfo5ho/SZ5mo9RqDCGYnS/uXXQOI6UypKc0tt2QrrZnvTDdY7zQvVFyyjdz48zY8Q21WO8WvxKGyF3deI0/qazOZ1O3g4jhxFQz8zcXTF48yie2ApRfCMcw3KSgF6/1yytAUHYJhSQPg2+ynzAADCoHz/MdTSYsMFeYJmZ6ReWKYMUyi6D8QnGYQ6ZlfcMdDHEMZgXBcqEAlB0RlRiE/WgIb1DIwe0Dogup4IZlEqZXUCiODxGF7Rbi0RYstjLlRO0ei4jHpuXXW/0B+m6g3AlRIldCkGXKms/osZkdpWCVC1b9DG7d4vmVYHKxLjhj1f3LqFq4S54nsiw8TCug5MMnkiquPEXGtxviCt1cSoiqYDd1jVzVTxcvXnMNDWRqKwD8l0fl/EOcSzWsVtwPr8kVgaoIx1BDB0juV4gagguoSCBmBBAh8IbgYlSpUroIIICURMVmGCcwB3NOeirbjBE3NoIx/ESJiCugjOIjtGaykoTXS4sWmLOIxdGoz8RixYvV9zi4SqnGpz+i+OhElSpTBag9alY/iag3Lmyal9K3Ma3KhXODg14/wASlV5O8dbF+M/7mHoHeVHbiW0cx2EW2ytyovUG7lcxajXmZc4ZhzB41CjUsdvcyb18xEOiKB5xucdZMxgRzjzMWCDPbb/EDeagKngA0D6Wnl5xGJ0jpHMG8RV9WkHQThNIVfQhqEIHQRqIIJD7wWAckDbMeY4jgAfMpSH5hVL1cuDMqeemnQwTsgjE5iQRJUeldV+oxcxzHXRbjFivU/dF+p6l/o4m+idTpzDodLmyJjpWINQbf0YMtz0IkuUKXkst95XGDL+P5lArHyRtnc4F1eYsAV4DiWB2mnbM47x2JxL/ABDODdDgiFy4hlng4mXl7XC9uNwbvxNN3XEF5zE5ILIsYqJd4gtjECRnhRFSKWvlx+AjkiyaoCJmer2bX5afmcRj0aQQQRM7hgucIIQJlBKhDUCoQolQ3mINTywe8qlN2wVRK8y5KfcqACZQTxcrtqu8Ss1kUBkc03r/AGpbV7CCA4ZYEuyLFFXUY9C5jEP/ACOr6alUS4uJcXoxZdMWotEWLLxFxFi1qXUu5dyuof8AwGG5c4/+Jhg9eJgTAymRKXERFsPo5g1TUAK+Mwig1f4jXvUS1UwDzFM8yx3Dp5zA3EtxiYatjPUD/wBlBidl7jLpnguoysD5lBnDBQWMrvKRgKHYKhci6lXaUPymJbAp/vARXox3BDDDHcp0moKgeYECEIENQiYHlDLz0W//AGaQIk2qCNEPcQQA3mHulnmdtgA2wBNKwkly0HhVMhlbykexQFvBLH3DmFsmsTOS43FFqLF6GMdzKMWcoxnMWMYy5dxYxY8xYvS3vGLiX9w6eYzcJXTXUl30upd9LlwlS2EvqQ+ulxxkjOgLUE/Md3bX9sqF2qW5a1G9PEC0+Z7HmcLNQw3BaLm39zREu8TFqZ6zcNPEB07xhlKVl5V+blrvc483DAx8Rp/KbXzqV6W6r4z+5G87YChQqtKYT3ZgYemfx/HVP0EQ5iLmc18zOG4EC4GIIFwgRaJWPEo0y1ajzOJY7sHcflMCXtjMgOLQs2OE/wB5moDQpjVp9xCNtqr3bggWhTu4Dro1GRqryG4uJBDAEBNfMIl7DncAA7l20yItRYsXvLixcy4vnoW4xZcYsdxeixaixYovS+nz0O/QanEuXXTfTHMvpuepeJvoPbpZOdzv0qa31JrMuJ1B0leiVtRqfs/vHBZhwu/ie/ELdjczRvBgoa9xYLn3KMYIW9TAePMpec1wRa3j7h6eYVwPqIVlRau78xzcS0bZzN97lbLp8SihxH8IAMLWq5jbYpfbn9vzFVMsE2NyC/5fSCngaPjELS5cUenSCnpgdIOgyQMVUCGYEIRUSoajFy8YHlAmkrzAIR3VLmNB0dcH8ypJ4ocTOC20/wB2jNcdoUrd89rmTWDiWF4TFHM7NnuYqmuGMjaX+8zUkENQrraHMJy1E1ytFFmXcWLGF56GjFuMuKxe8ei1Fi56PiLFi3GFj35mZi4tQYQm/wBB+sxNv6bg29CNyrgV0GpuJBHMyfKUqOwqw4P8n1MBb3b7xHATcvR5jaYvtBSf1LFHYqFDsQRqOPk8SjBb8wO/MADEt8uIq7ypBxebgcLGKTHqAobyvEvvutxkWblysvqMJ3RWe8OzdjvWA/Z+5QMteYZq0V5ifNH5l5a55gwY7izGZFRQ+JjriB8wEIahmHQuHENDLD+IjQljmOuNENRELTuKKZs3Z+IUWtDV0xbWZ2OZlXo+4g/jcyAwGiv9/iUB7h43/rgwG4sgt4/3qG1briNvM5KludQpeP8AMO0XBDYWMvCYIuYu4xcWLcctxbYxYwuIuJcWO4vT5iy4sXMYxal0S2GoEOlwZfS+lwl9Llw5/WNcy5eoMZXVcpWwalPonxrMpNodM1n+qlTV4TMVEzv8wcG8zFLP7jpHtpjoxuNAFX4h8AE2qcDUwKhlXPch5J3tjvKymb7Fm2/xFQfeYly1Gsy41+I1dpK7zE1Xp5rP5l9suUBwV8pApLoh0GMfhgyyXiZjuadHKXlf6oZodIgx/wBhAQhOIQ2S0fMGwm8S89hxErlL0ZloFieJnSm/JKled3G+ar1LxRZsySvZxEH33zKW4TOPcu0p34gG0MdmX+U4uPsFItlCYMda5Db4J4jXKJaShi9GRGkV6LLiy4suXFixixYuIvPRZc2S8effTMvxMsLqEq/MWXUan7xel1FgouYMGWg9NfovE1qC+oPRzHEqU1kzMz/PoR0tAe5dH4qdpnkhOGXvGKYvxcoCqz2m1VqIUvE2PuFMk+oXsONGY8nmKruTw17hUv8AaCpEBev3hZ4+ILOaJUi5dwAWphT8xbvcFu9l2oy/gZyMs605uB/ZfEZIpNHmv/YoMLFy4sehdOO0H6FpAh0CEGbJYMG9QUe5MIcOKmU1KrxMJ+QgFnI1KI/njEryIY3LgMj2lpTtmXbV/P7Stbr3cADXr3Dd4B6l1+7FTlj4h7C2LaxAtX1VsAhwnRKiWeI05lsvovNxhhWLFubPPReixhcRh9y8xtLmpzLz1xB6c9dTL1OpAl/cuHQLS/0sKeYPxLl9IzZaIB8JSi8A8ygylvdnHeSGuT/sa12jrdF6pi9IZjBXfF5jbBmXPglXO3xLMnzA/wDcpIw5gEbie3mNS3HuXJe+AgfRKVefuCp+bhvF5zy0f3MDGIqgADHJd/8AO4vY1P8AvazFFBmkv9ApcEH6IQ6BKhDEFnR9GghQyMNPXK6iOOfYI9HkL5jZbTBrsinkpLywqCxdTWBk2O4TeVa4JYqsKzUFDGTgYW1VjuwVq/jctbzjvqLoCB/mEUqVz4mxzXeZquWBFcvEWLLjCy4sWOOL6L0WLFiy8S8S+m+hllW9KlSpX4lSpUeh11+nicTiXLuDcIP6BhFz4lKzKXZYQewS+L/5ESgw7iUecxnD/MNU7MbhpJp4Y9OYCxbqAcsSl3ncyt1T6hQrtGswEVHZ0SgHaCN9u/MWSs8yvffeJoaimW7L3GXVWKjGsGOYV/a/Af2v1FjxD5MOtZL5MxWilke4ooMvqz7QZlEHiCGZs4QOh+Y9COuhdKXWyWuJUJe2eJdwVktg4U/FxUr78+I3zV2cENHAN9oLlWSZ7fflluQO0IMlXlN1AV4eUlFrn7mYeC6JUGh1TKuFVX2zOIlHMIlwzTmUouYvxGFuLFmR0XUHE0lxYw9AuYsY0x6GuhuVCVKlSsSpUS4nVmv1Y6cTUHEuDCr8y5fS6l8S9npBLVfEHLSxP96YjS3mrh2fb/vxDQ36xNRXn1CXD+NQcZxHLmAUFV5In0TgueYr4bmXiliCue7Ky+OIkq8MDWSDntAQXNdpXJdcSstzKZL7Tm9oqG8UBHly/vBUFxraBweWC4AJ8lY/f8QariL5jhAy4sWLMmW9vxBD+jENQITUJkR10jllDDEAraYcdeNR1ZqVLFNcf7/yWmTDNPaaIjXZnFqvepUNihzeLnYfMQGrWBKLJeeYacH7V6lBriM2hQ/Ma2qFDkxDzzC3NhqRrJYM3FuXGLixpFjA9LjDFoosWKLbL/SEOhmB0JxnqsWWRbItS/PQx+kZfUbl9OYsdRIMrxK88BpwFfiWSFf03+blIjezvDoYt1KAWKN5dzCU0sX6Qvd8Snw1EOV9x1tuvM+WWN1GlLWOW/iG7GJe1iompeYlUmokot8x9ipc02c+FbgrtqE7ZmJMBglQlprujX8zwu2DxV/yQY8RRQeJfRixHrBUHUNdDXUiTBBmbZomZRYPMNfZGcFOXtBHYWC7j20ld0lThwp8TGYA7jiKyuDv84j2ul78fUXQAXvHiAvF4rZLbPmnMo7Z1KGAbmBPzK/rYh3rDxcNNhOiIixZhLjC7lwhhY/CXF+osfqKLL8S/mavqQhCDDv04lxY6j7lxm4zXTsxBl/qvpcu5c44l4lGUuXiYX5nZ4j7hOKtVjLmZOWDi+IKXQ8kzHXGWXYsziCjIPmJJU2DMb4HcRtPe5dQL3zLJc7lxBqtTBn8zDdyxExmJq/iGrPffiLN57Q+7e46bDjvVRcYL8tfySoBjFti29l8CxBVOiD8/wBRQYNxVzC0uXLii4mO3QGoM9QhuV1Cv0XbMlVDcG/EuD9ko1+4tE+QqU5A8/7/AGpk86u7iDyUXzMbd+19SwKvda34hUrOdx3y8kyuQLsr/eIKUfjOYcFZAzcBVUEYMUKHJcsV2hpgshpiOVmXhmLEXMWmW9Fpb0MXLixVl1zGHEZfUnMGEu4MGoeWJfQt6zLl4ix/RzL6kuoNYg3+q8wZcWpWszFnMD8TUQXxr+YVJ6Qm6Gu0WBTHEUZ0Ls7xxE33uZSw7MBiuO8qLfuEFZ7RU0nLBSXi83KQGiA7wyi6Ki+9eZkNUxLqLOOSPLh9QxON4mKLjwQS2nKRBWOXwZ/kiXeyG/tk0bfvfknYV09cfiLEGLJBsg1LgxYseI6l+WDoH3MoEIQ6B0qDExSlemKCTBj+IVkfUQIzf87maBZxxLaPKq+JgZDugVXdmyADVeJUXzi3c7hY1X+8wlzXlgCnTUK23Xj4mV9sLfWuxHNTJlADndcwFlglSF9AWosehcGWS5cuK4ly+JcWPecz8yqh0viXi5cGDCCPaMMXF6eZc313XS5fT+Z7Qi8dNS4dFy1zERla33ltkC2/95ld3aBnhLmH4uKYDEW6UXfEagt4iqs2cQLo5NkAS6cQRFlEwl1c5mx/EsVe3TDl/wCRF5xUsmcniAGPuOg1XeVa7RwF7mpNgwpN6DHeXH+TmvxUsNG3gC38So6lG9pPuFgIoQMVQYMIWLGWQQO8CCoQhuErpU3Kl0tJ3ekCge8BGkLQHUvorLl7SjFUG7F1TT/v+yoGj4TNdK1AyJk/uNttjcmXyYmgWHZzAGXAZIbQOXUIt0RH0FqELmHMd1mjmZguUIsWXUdJc7pcY7jF1GLGd5xGiV+hiy+l9Ay5dkuXL/VcvoS5d9b8wei4svMWLhlb99B8Bl55tolg4q5beJarX1O/AI1Fl4Z7+4ZoUkXPeO3eLMh77RUpV45lNmztczlj5mC2+IY8+5STh3C4CV2hwRoFuvMItrMrRQiIMDVCd4uRDB3ZrHZ9BUPCr+VQT8n4hk5QV3A3/jcWejhlBuKoN/oW8SnaCBcCVmcQhCECV0I4dQvJUonCc3Y0S6QpqxiLa60eSFMPGZhAuMev9/MPNOMpj/VC2WzL7hHBxDbW+ccyw1j1M3CzNRbXtuBRrN8VC5Vl4rEaCWuzxcAKVRqMZhOJfWcQwRRi57S5fxL9y5cvEuXL6VCHQLl1HLGXqanPS5f/AML8/puX0vrx1X0uVOXtSxZ3mFUDvilg0Sy8S3ZMziXFJiJlqVr9JnOTe4wUfcTQu3EpxsLu0lgytvgg5X6ZiN+0AWM94KraMVoee0paauMDnvTHvqMaxLQ5Tv3i34ZL4TP8R1UYBYF5Cs/BfMBMEEx31/EdwYqg2RQhpLiy4vTwQQIEOm0N/oqVM1MfQ+WbsQlBdzA4dQQ06/H7RXll4fMyXIY/37xqXxeypXS8d7j0b+dQlcZ7zeXT7mFd+HMIxSPuKc3nEqML+IuNeCW5rZmPzuDMuO2d5TRe5btGFSMadFy44S/MzW+l9GHQZUqLbKlPSv0czv01+i4Tmc//ADvpcGUzlU3Z4lcNOHzDthvDuDpmGuoLeNzKxuaiN/lL1gimD2mU4IRhcg6jx/wikcJyQFJV2rcfKC3HMAi7zeNQXkXiCGG7cypusRrSjj7jqtI8Sk4UPALcK0IqVphkKtDy2nwnwqejH8R9B0Bnt0GXqX0u5RCKqVK6BUNw+4Ss/oSyUdC8SVyrhukPd47agMKOcGJj+GIi9gd1mPdd4CCrMRhF92Gvm/7marcnaYt0OIaYL7xlp12jssx4mQGsbiWRMH6wywVF4zUuDMqVFXUuDcuXcPcuDPz+gMSty4ypUqcTUvrueJdzXnp26eprpeZfS/0X0uXL+elwcShTlNssdyg3eX/fEo6lO8xqJLnxCtL9kJuDTUNHHxKhxicg44gVgHxOIRBKU1jZiWyaWUEXtgVNxZgggBZvXiFSxXDLBDlzCLUvlhS8X81M2QLFY1+4/cIqtFOWPe7PTY9wK+ZairMr3Y4MUGoOIUYN9Ll/MVS0CDpUqBUCBA5hz+o3NkEzwGV8wr5+YSw/8l4pFVo2xsKrOfM1ii1/vzMa39DAbsxzn3MFm3VbliqVWszPjOMwl7dm5kvNok8nf+ZpvvncTqrw3Km5NWXG9ewlaZlwZjYltS4RcbQfuHaXLldNwnxCK9Ho467l4j110vovmLOJcHXS+lwly5cvHQ6rUrlWPSmFNmEWO0HaIv8AU849/wC8+WE/MG9XCFV8yxbMzYNpcLU32b/EIGlbxOY25y8wNAr1GAyBzfaYBYPBdMAplqhhUzgzd/cu957Hn/XFFWgjMGupN8M/mUTYvPdPxkKoprTnv+fzhiKLvBuDcIoQQsYW+gWQJUqV0VAgQK/ULJgh3LiCgXBQwV6i3cs+Q5DERuQbHUxVVPJLsge9S3N4fklKIFma4m98JRxRUzXqOgVntF13fPuVBvRFa8R8OyC5ROYcwg3NZlyZhURZcGaQzBqG4szB6E8dWVH9PjovQ/QGbm4zc+ZadFy8dOJcupfW4yqNKZWZYbwDFdyrFw9N5h8x+3mOXaHjLoEG8NQkcV6njgtpFCv2QA0yyutVW4AjMaSPl+xxBDh41/tQiTBeHN1LVXF4TUW5EJ4Et+r+mII+AMj2m35K9hC09L8v/WDUIdAYNzUUIuYly8S/MDEDEIkSViVAuBXWmPVhuVTdN0BQXiYRt4i77DXNRwFW4JaoNnmM2+fzCzjJuENU+hhgNAfiIEy1xLV2M+4s3faLQuvcytDm+PcQDKvErXoKEuIawCWoXMDMuly4Ny4M1LqGXcMTcOIQlwly/wBB0Ga6n8S+nnrc4lz3LuXmHiDXQoy9RcwZd+Zcvpey1TtLfzm/Zqo7LyxSqbbqfaXZu5tlQO87jLYDIgBxC7QBFZg1MHqViHBvxHexnsxeZXqGyPKMpTbpxOzjJdtx2VINnYH7L6msQUQR+V+B/EIQ4H4b/dYGA7wMCQJA94HvKQF7ntKd5WV7fmBAlQLgEolSutSom/0GDoXEv1KfWKmxvUcVXBqyoRGR2zWu+JW4zLa8NYxFD+MxFZY/64kCiqaMS673qYP4qLW6j/eMGJejYlEc6OYJuVEsDOJSdL6B3DUvMNQYQIdL6EufPTn9DFxNy5zL7dLJfbpeZqcwZc3L6eZZN6jjpeM7iy5W5hVzMpaFtsxy8RXP8I3feLCY4g+ZZKGflLLhBdoLX8wgYeF58QqLqp5Q0Su0ofExItYVLanXaCN0i1Xn/ZlDLU1TNbmZ4LfyvomD+Y9V2t7ALT3VfLFRWl/Lc8kD3gPcBA9oGBg94DvK1uUOYl2ykGGVAgV0qVj9Bqa6PTaLBcsIJjgVC2X1WlbKYcwlAW3mMJXH4lCF+Nwm8/xLl1ohHqNsyYd3FW+fMcGtvmOS1Fw/6+g0tOQxKGgoc5ILFAYud9fqGDMGlwYMPMvEO0IMDMJxqM46aZ+3TXvo9HyRj14630uM9y5f6LnEuXUsXotepdS5W8w6SW26uKt8RRLSXFSrWKg6CniGHQuIv9cuv1L4gBz9wkdPzASUH3K+m4hAAbmcaVKzBjzO+U6YgUBYO7WPmDSUfYf9A+I7CpbkJA8upVADWVbZP90St1ZiHl+YEgARwzNtypzDPlKEGHKeiV7kelBAldAhFY8yiVcqVDo76CMyhshS5QkRg1Ioe2axCdk9czbccu0Ftq/3j3eb9RbyV45nZZ5lG/3izXMWDiLfxLPOYb/7KD/MAMFPuHi7sgKDRDQzK3if64PaDUNw/EIPECEvofnqMMdEiXuMWa6XUZccS45gzbO0W+hWYNS7g3LlzHRYNx6KhlKJ3niVa+IqrzHnMdanYg9vzKm86gXAjxMVrMQSkmoIg1CLyS6zgRHBxvHMBVlusygwErdQiuzCFJY+Jqhzi4ZrgW1v/wAhQ0APRgirq+VV78VLY4EO6HP0fEBbmBOYZ7hnue8POV7wowec94+UfONOZ7whKgQIECVUSJNMvo9MHoWGZoKubMQCVohvEfEL22KlSVWhMrx5j+2MktJYeJ31nUvy/wB/qjM38zPZrcdD37wKao6TcUBhGyNapQJcIyUTvLAbxuXVmUEMS4ZgwuYNw8dbh0P0VGLiL0uPRZfS+vM1FlzmXHEWoMuual9b5lyqMrDtLmOkfMuV7bjfn8xGIh6THS8RpbqZeb5hjZjzBDkgDhKbtXKkHanN1c1KzDGrtAJjEA0uY6ruXh3irL4nWMP2kwP9iZVncAHtWj5ZcNZnZf7IyIPFDznvL3B5kP8AT0WXP5g+8H3l+8X3loDEFCkIBm3WVFREuVicMd9MnoYuYLmKWjiboKKsp4h0ebVRcjDWSKcouLOC3VzJXWPqaKGfUNXVbxKcE0l43F3O+sTeXUW8QhE6VjMJOSBLaz0xYEIQIeoNQhD9OuneJmc9L6s56EXxLly+l9Llz4nzNMvJFx4l4ly5cGG5crYzxlS2e0yj9YXXMvNwF6S3QTQDluFshC6Slt+Uuta2bWq+0NSw2JWIfiK4rKQplzAcCAWCeA2n4PqFO0ckeFs3nqz4TAzs74NfgEvmDvS1uEFcy1uEEiy3mCnaj3JaCEDpUNwej0SJHmJi6jD2gjGZkwPStGCSrlggO8ykrmNlzLJmIutP3HRh+5Wlmo6rMZeydtnzEd/mPbU3TUyXoUw5I7uMCXAClusXCKLDklgbjzSswm0IT10uX4l9LxLly8R8pfmPW5cXHefiXLly/E3LzMxZeYvXUvMW5eoPHTm5crhQ4Dj9ygNNx8phth8ZbWY0LvLIDmCdkILKVrAhCTt/Ri1WqrzL+28YBAMTtDBjUAAq4DdYlbM0ChbxLqK+5NH4CLLT2OALYaBLNXLXqGoODBMOYRazyzHuPe5d0BzCMG+hb6DEJVdbhluXUzLj0uo2iyiCxlE4QYlCy08TZ2gLDNQ7Sm4aUr8Q05TF4neVM5hucy51knN/EdGXPLNopmLtGObmswoJGZYCnHMHGKHeXBmVnPQzCDjxLgy5cvMuX56XF3Li3LqXLl3LvUuXcZcuLFly5fPS5cf0XUu4MvE3FQytSpFzY+ImWVSW3cF3ncdy5MzVmX6YJWYUABjERSQ5mUdS77QGncJ2s8Yh2D5qGjExMT8cUy5jVMfEyEEPa5pBBXNFS9GmdsCT0X0UdcCbmm1+wi/mPQs635OpKGIMG+lwe8BKXAbhFale8TEykYT7jjFJVMcp6LGEoJdctuXi95QE3BCO7ZAy90oREXEoEWMROaqosNYfqVcZiizl/EW4vyRjGBumV1qC4W0vzAUG8wmsyszKQEqkr3lJTvKVHCbzA6MpvuX/AEBj2l1zGFxmLGXNxf0XmXNS5cuL3jCu/QMuVuUYmFzMkVYxrW7XMvu6ly5lstC5hIQM6h8oA3KBIouvylI5fcq455g0dym9r+pSMeJQCvxMLEEEcqlJEbpzKUZA8ZfxBj+ocsDHak/ZWLl2HSqnzkxlgpm9IGZPPSxRZnzTImvUGDMeZY5ijmOXeB7wJyQHeUrcfKe8fOPnHOBhbmAu4D0MEybl0B/qOSc5pKziEmrajXeWahnZWlzAbZs9z6JVtFSxcRw/3GPiP3HM3ia4ldimmYyriECKQgqXiEhmVOZXvAd4HuSuJQ5nvCc4iVuUgYBh+ntY9fizGV7yoxEQa+pXvASkBMmXL5nBKuaiSG447dC5XNWTUzeE7mH9ojmCsSMRgFsAb+oJy5lXP3ApCGrKOnPfUotvZGQpzBDRCqtPiADEKyYUoeI77lvjuz4oanImOFf6GOFQqzwt2Z4sX6JR+MvFB18qZY5wy6WTxFmP3FHNo4Uxly+ivuOMU5nfw7vyhY39R5WRJzEVuB/6lVge8Am8woLoxdTMlDVwO8ZzLzct6DEYDeGCYrGMU8zF1faozV2RUOY5MRCiJ5iVuJBjErxKjBIdRq5TllzMKM3LTK73EBse4Z3TjMaFG2PJA94DvA9yA7wJv8zvJKVXEckGjMD3gN3A7gHmBeYefT7zBuIQEvL8RfeL5ZUxhMkBAkMwMruEXqpRc0IY7l4zLxLgLKFgKuOCuVWU/EVmEhyA4hjdTC3lMaU+5mc18xgg5ee8M2YuWYPmoYMalVaQwFZglASkLNzHEoPfzECks1+Iil08S1inN8GD+ZXwS7djxeyl9LBWTj9o2v3HKEkfXCehqjj89Cgxei/UC9xquWxBMRviPmsv5/Mc6GWytI1rcTOpgzuWtRSKtGSYGYiqjrzmErYK1i+Ylg/cpGYlNw13DqrledTKPyhJHNlxSpiHIuQuWr1CLYxaIAJeol6l84i3ioqK+ZeKIheYtajXMIF7eZalr4uGFZvmGDlKGyMNkBz0AWA0w3QRcJBr8oaqAY4zEhFMRmD6IoJzFdoj2iVqOnQFcYIt8zOYAFGBGFkl4ZlTD04plZi7MuyZiZgUlCm8S73uLl4wP99xnD+0AXMCtQuUMvMpGLL0JsDLxK3lXtDV5lmkKtlwqFQhglEMlhrjtCkpi/yhR5xKjyA/MN3Ijfhb+VlW+W3UkGKNvbI1n8fdmdmZmErvHGiZk3HNYswdTyQbi10J2i3iKuopYj7am5K+IjGZVisRBXj5leJfcf8AmSnVqCpZAGhAG6waivbwCIcB1iP6MLnoGpO13+0JSKjYH/bIfQb8wkoe7jxH3LbGAbbEYWDxFhScQiMjvLGwjgCuIpb1BFWuONIy7zhWVUdhVHkhfJ8xux8fETKpcFaYIQ9qgksRd3cd4H+RBOfzApNp77hVNfcvTg9Ichp9wgbv6lh/KG8nzHKbQgbgwioqUTHMeoYQpDCBqdhAzFXia6iWxUYIWB24lmvmWK4oMxUJapUOI6Ua8ZzhKq2o4VMFe4WfgGC/bWBFwgweX5lBzDB2lNhDp29RCq4iLLWAf0SsAIQHebRNUD0gT+um8OZWF+czsupUywbfF5/EMWmq1BUlXlnCHprGTxGNNUfb9UybzLLBmZvpMxQQa6DEqCBiJiY8w4DtczYghxKTiZmJqxSEDiYMSgI0rZAAZhNOL5lyFeJdZxETX5g3KfMDo3XKIYCzRzfniMMPI2HFpSX7Zp7nxZZZfcsR5siEodSw0CntX8Je4cG5SWXplWIOmDChcE3mcRaotTc96mxjBQFMOsU3BpirKxuaf7xHH9RZgp7kzVVucMU3m+omvOOv82Ib+8OVPxB9z2qVUIcACc1FDNIVLX5gAH8pUUbTFzKyj4jhdkeo++ZcZXwRdplFFCuYVZ7YWZpe8s4YBYahglw3HKntmJWZQYB3qVFd0wNTEAyjMIEUMdYY9tOWVxT9RQKYwGOkZw6Z9FKkcSwSoGVzVtqfvkq5kn5QMl2wLtQrH0neQ6Y76ntGTTdw6pvU5UAh/ZDKr0lKYlIG5igU7xrGyfvNEbTJ8wUIuoyW/Lj+WPccswe0pT39CCcIynA1c+Fns6Rt/wCQXNU8X6I2CLvplnUIKbj0Sk9YQKS1S5YiiQlRuKJuJqZO0tQNd4wZX9oduEQcKjjue420Md4RCyrMCVa7SWna9Zd/tEiKm1qHnLh5oJiJkNzBWD7UvqoWhNUHBAPfK/EcMAxzz5GHdY9JUsT95WvzcwQYxCfUsXnUoE+oVAUXKLEMinuLRjG5jSy0fZFcR7f7IaTfiJ/yJXfwg2K81UrSvqFSi+IFDjIqb1c5lWVUwQs+JwMeInS0MBc+Iis/qNAL6iBw9R276j05dpX3P1K5Dg0inEGceoUdZjXiENSjxFMMQ5le8C4c92G1CYV8WRWe2T1BpUVzHRxJjlrjiU1qGUhUh/BBI7xN71UHcmpe0TiVHMo8kZplMetWWAXb2mp0hg0lNWZhFVBDUIZl0xiNpSuYupHPMpdS63GlVmadpfVEy/RAGX3LHUaaEBrsVb7ijbX55v5Q7EFxL2swjkgOPxBVZO0j9uh4onaCg+0XVVLs9H6XRUFplaiUuWCOYB7xHT6mPCy2Bf1GIMrdQBjJu+I1e2xYJhplhGyVpueBxfohJdJsKN37WvBmBwhBTVZ5WA6/7iVzQs4Kx4Frvb4gqyvFQo5eoru54IIWNESLFc0Sgn+Y5jDmWPgZjtGneOVKXtFtXPCbmBt8ER2JDTPmmWNZXZz/AJyDPDxB4pXqWj+EQ3X6i6s/Uog5eoo7u9Sj/XoRo4fUW1+E459TWA+JqB9QPT6lP+s5GfqMf0mYnXy6YnrINYlSI7QIRcovtK3UKN1PSCGJNLMVISrAV1OIhu8AxASiYGOktWlxHW4ECPFnCx3c5ZY3YDtC4Mdq38QmX7CMpJM0Seo3hPqVpkfEADEI4qV8S+mPUwjzBoo1H2YvMwJsz4OjT5gNWmIQR0tQCylqecv2Q+IWeT+4JCY4AvV188H4Iz8KEHCGtwh5AhpXZykLww+IQpL4l2/VKqXhLrWu6SnF3clQF/qLsrfJNGOeJTAW1fM1YnkhzuN7TwfaUxpHziCbJYMvJ2MsrhxQqN2FmDRcS+K+I11quZiUW2otBtqXNQvxuD1CO2QDuboRqCKzw/gIMB9yU4vggn8IIg/JqXjQd0MGJcigJIiNl949y58QEfxlr+G4YNX1KnslasIpNW1hGD9QfD6hYLxgUqTCDAwBMLEFlNaxKY7BEcE5KGBcHQPQTvQPZEnEHtKOIVhlDtMou4jcV4uXktYntEHSBtECgDqCwxNkAnmiGDQW4QqXs4yGlygXxGUO8owZsyA6Q9wVgzjmaN+o80L9Qd7fBL9R8R039RLLlW1+pUU0+IHBconmKdoswNROTvtNneNNSxu5aaUczJaM+WH/AEwPBggLsXc4X9m+0ldCs7Vw/wA4hKGuDyg1APcYm+nmPsvvWYhsDxBdlp3hJwWypmg5Y3KMFmcS75FlwEoXMEhZeQ0xK6qagFijQ3Bp8iKBs9qJFFuYJBOYRzBDcFvM25iVbiNfEFlnE8JL+PqeK5lWNeCL4/UItj6iq6fUptEfUvwkwQiWIVqEsrT0ZiVZezBdah7TkVPfmH4Yjy1hMACj7iegiXEpqoJxOGY9QH/ErNTDo+oeJ9QI1Uo2QLpCH2mLiN4qV4f2jOyOTMZyRVVLDiH+WEbDcFiGFpjiVcU7w8pemCMRBMB6NWansmKIEExBgohLEeBhb1A3MRWlnEE8x3yoajh7i4x0FHHfRS81E8iGZkDQ+Ia0JxtiLD4hjVIS1IWlkHtE8D4m4E9x1dcRuE3/ADD8yri6guma3tWvzUtnxzGj8qywTUOVYAewwR0rQ4O0z94b7y8I0lSkGBY3Cu0BqoV2JRYlZSWSxlYlxz0WIHMRzO8ncYj8yx3KtZgMqiF6lfUsZqFepQcSowDqC2yUPEK3AKC3vD3RHK32h3xLHWJtExzCcspVJiCkKKylMzMe/wAwQ0j4b2l2VtdAB1USvEp3uWu+mdFwYrMCX+UZCT5i6U+5USB/aFV/ul6v3xWoZFWAQD3MW5h3DvcDe53Ep/6gDmU7xC7/ADEJBeYKQu8dOwuFeCLjMp3iM2kERkMyhuE5gBuW8xj3csYNRC0NRpewzqXOJJyRR6hbmOMcH955yguEqWprGGzbcaeMQxomXGGN5PEFaJ6lwphdRQm9dchEgx5IKc49wbfmW7ko93L+xBvmFWWteAz9KyrBhg/aNYDLHOrn8TBF+6BEospzdkF2wuWYNQi0G3McTXnosly3tHziJVhCErUs4lHEA7wSikG8wXmUuCm4qNsIQC6jd6jo6Qz3Rd2PuOwIy0zmWE3eEs8Yg8vcyP0Mzi/zDAp6lBSoVK/Kd5+5SYZScTnWJriO5YiTlhQIEKUMMYRK4jFfgl4YIeXf3D138wDFpL22ENzy/uWNsDamGTSCGkKty1xLcpLOT1BD+0qWbHs9wnhArSJDIgcCC8IBzmW8wgyk4MDbIIbgi/yirqQMwYezUMGIZJ7ZwP7h+B9xUwfmK9hl6pcTAG2/U8Qm1SaYfcq8wELcy3ziWI0XcYL6JYl3Cj3g8xZdEw4QSD53xKc+KIlVijWIae//AERTAo6K4/YJYOAKxwVLv7qsurm3EQmm1PBFbOARQsmMx4pAmLHsZoB8RDViFyzlYeUql+5UrpvoxhpzAHcHvCBwQe4ix3k7rpqG4VQr3jJHeYzhCFXRGGrJYVcstisGashMKrT9S+Cdpki63zG8/hLWQzxzFiy37xEfYgRhndO8eOxBCEmZS51K+ZrXB1c0QK5iTQhzIQ3HCS62xW8JfoSnaXlXMD9QLiBGoFtQLpncSDx22sW3NZs1IUXAF2vkjg15JjWfczmkWNpR/vKjwm0ddxjWI/MLW35YotPuWTRhRN/mEKN8RCS4R3j3Dn7DCeW+zGbFPmKB/Zmxp7YSv844NiHIpFlkPsXaMDb4ibQ/E5Sx0GkbyLmkygm8RIuc8qa6GsBTv177hlK3LUItUOgYftO9ODspasC/SxxXdh7KCDVxmlA7lUdrcqlcKjmgSE8QOiBC3kgIFbzAzBfcMKjBu5vIw3AMrLHrkC5RxFWo7MUf9ncQaF1kxxzBMCwpVxiBF/mDWofjo37EZE8IvhO3X6lhC2n1FVyfE/8ACjRp8kHw9QHh9Ss0gNRhg7RK+CWuNTs0HpIE6Sl1+5VxlPb8wKC2LYZlDDLkw7RIdkHxPSYhV4lJWWLgXicFQFYm3IzFWZhHAgYGFS8pBGKPUusoOr0ShL8kyV2eGN0oHmCSqnuUh+EAZox2UHhnDibQ+J5J9QCn8cTXqe4bpM0IS8GjceBt5ZVsDxcAYCeGAFIHmAOXgSvGwTpe9RGq/cLKueGCLV9Qt5K7RVQniNbhGqLUqgB8wIQ/cFsDfaU+JnDEJbi5nczIUh5ynP7S1ncGt9JzMotDgAXttKLQ5cQILgTQWQHuSmG5AhOAZZIp8SyfMleWahYJS2DEC6tLyG9oFJXuRWjOKQYxIAd40azPOlfbDw/iD5+4dukNl3CoLhVsIdw2FO+i/JGjc7gI5MIB0g+6JRW39zkopuF/9iVGJwanCZbdJHaLyL6nGtmnuN8zMxzz4nKLE+IRwsxcJ7iuIrdTtPxL3n1E955UWtMzah4wrwgHYZnLlnBEwtRxQSujOK0KqWh/F1L1l6ILt+IV614qWmRFDSnuOsKG7ZiFmPFA9qiPFeIzhv4YgGMYTL9EEbBfOJarIkZfpmcbTuRWW+xNYBdZilRlixR2IyDDjUVN6IQhvbDVFfmJgrdoobeZGyKmc5gxIdyUuO8wAg+5jAWWW8zNQekslj1LFjIOMTvExAfMrKN+WHFL8y2HPULheOgSZTPEvN+Iu0eDvDg1jPZr5BARGX8SUvqUgizMgI+qMrX0mLKS9kshaqh2kyyIpeUO7JzrF8tzvxfZPHhg9yu5Qcxk7CIMNh1iA4l14mfWe5CAcVBzM7fQG9464zXRu5jfuPfjCLHKJ5VGcvuJNvUawrErlY+UPaNhFRuWRc9DKyFYQYus94WsEsFYjNzs4OsVAGoIgHCBcEAME8BD0nmPuU8kr5j354H6huBjXSM3tOFIZSXEVGeIsO5SRq+YFQX2mLZE4aY6qlBQF8xskDiNNj3YNItebjEbE2JarfGY+xWVWoTKV2yw6JJZlA9oTUSX6yeGIyFOLMqjJzCWBeKlkR6Qxz9oNGtYaiTtxqWxkbq4XJ1zZLqqrzAgWOtw1iGMDe6igSNFtxCFaKiwdW1DQlDN2eYM0GztEFdeIgTDzGqSOalK5+YeoD7gRaPMw417h5zJG+neYjxFbcFa+IDCJSAqhPcoPgjbYyfMUMxijA3NC4ZuKcgzhsw0wXLFBPUuYWckhoE7MRuiL8NzjWe2Y6LB38wpXpLHMEQdTIl3HZEuF2jI0dDYjC9AtjLKKzAEF7RtzHHUsnpG0DtPWU8S1xM+CXZMwhszLyEqZ8Y7A+/coc/axAxjqaIatEPF9S/n8QflPMg3az3h9pTnqXLgI4YntY+RFVMNsSog3G6FyzMM1HpgdpUCap8xaq/uOhpO2Ykoj3I4QAae0aChPJKyNNNI1iohWgR5lOWvmpWPzCGUG4uCnDHcls3G7G4Ai27kTJd8bjZIbxYi8dsqXUYOLWRVzjQDeMulXYCv7S1pHswQysYK/mOtzcj4j9rdWu5lHzg1cNAqOEzEwhTepQXWDzHakqgunMqvmJQ+ola5i8ELIp29pQF2xFb3gUfpj7B8VBsi5CpxABDgIpdxwCjdKWVrvCUyvxmYYeWSIEvUUoA1FrUsFbd9ie4TwYFPv90WioVhhkCqMMBuCOZd9BQgZYkUJTIEcyTSRDiXldBGyLlLiLjBnUzO8OiJAv8AaW7SU+ZT0JrcpGVTUW493S9kWLGDvMe56x0YY1ETt9wHLBw4nrGnEo5lM4IEBwlDM1xDtUu6m/RohhJiUmPxBs7SgLPxMR+0oJbqeSZcw84ecD3qUvcq5lL3Ad5T/kAwEcenxlkqWioAwShmZqKQQ9ybrnThL1w0WBLcLtQRm2btZMcdatZTNV3cOOHIMxWCxbVxILNW67pRBOMAfVwAyN+IkCutYP6lmSc0q/xNaDQcNwUuW88h8zGj8Kl91B3GFmfUrVScwb9nt/8AEUQa0F84YMvQJPjl/MwtoADb96l3kF0R2oE+4vMClLJrDSLY/EGgombfJFGCKkoGG9j6hsJTJAFuXHEUeyDY/iVuU7j9iOI2O5fxEuiiAPZl7aOBeQvUfKoAYeKc5u5tnP8AESuIGg5+2CSuFbiWNu6Ys5NO8E4TwJkxfIpVYXyZrC7I/dKTJgdGi4LKyhbdvfjcZ81F1e12CLjEVWhVkqB7/FQU+BDDdV8jMKNYfModgdBDcSVggblsW8RFkjPaA6jhDntKhcuxqWxWMHxMsH7Ipn5xsYnBL1GHPQsRMXcYI8uZkwU5nfTJCRYxEGIhESKNRHMXt+YgdobFWMOw77QANEQ5l6xcUMfiI9RqXBY/MM7x76kiw7kG5gq3ExmDkgwuaT6m4ZgQWob3NXCWkLXOoHj5gqEpsuaEKdSZu8QaXKsSqM0p+iMv7tDtUJWXupv5mAFXoy4+CvR3zUJo/SzZ5JjTZVo4blFBDZTSxU0ElAYr9x/Nuv8A0lC23mfww8xjETlJlXmYYqIFcqtD8xjUS9yzSF45mujwXGiqZ7C9m/sgmxOyfyxNWO2yD+Is18X/AAmBxYAZfZZ/LHKZ3bf2QVqEQfiOUjKvlswx55/KIoE3nD9ofFUWZEzqLkNfSYijXDRwrzQBlUAvEU7JP2BR+ZmNdqK+L/EA3wx3+2NCloUThVRfJLOfO2ifJGChDgIVO4vhjQKYftLqJrAWLhV3u8RwVQ/4xY9ah/MUcqqS7KiYRLeU/EFZ6jqUtGjgr0sruAWQhrfiIoB7k2uDhi9fPBKv3m0WeoZrI8xKYgepvpxHcWcE7suXRK9oZrpANywM0sBitkLqsQ3crdQmjuAFkQzx3gnvFHMvWSUF6g9mA6jd6WFdo4ba8k3MFSvMIsMXhgryfMAZi8P7RFeJx2EGgYgpGGSAaOE4rdxq3COcwCoBX8SvUumO2YIOJS5vEtDcGpt3hCDCHQmnpXeEKMVLwpBzEL9zbdvwy2lUXtdxj1iskrSgMeoq6LQsfcEMjQL3mI1ysSKLQBaO0BwqJgs8Iw33U1uyJUrp+4kavBcYMr/hmPCz+f7I9VBhsf2UVbSbwX4II+yIz4LgJ/MO/SMD4Pr1fgI5YseT/MKdE80X7sTd/Ct+826B4P8AEB6rbW6PIWtZa4muqT2+AggYoKw71VfxMrRqU+Df1FAacP8AysPxlgVqVZUvlLZpjKBhIBsUfSGfE0i31MeCaoIebDmougN4ZZ7O0aEJCSYcwGj6oijgeI3MEHwSwYD6js/xFYLFYieIecQ7Lg2pawV6j1GM2HuQdhYTSM56RV0xKjYxdHtG0thEkZGtk/5YnOKDZLtpLJwirUqXqWyMQQihq4UjFJhqNtzeY64jf3CVVlyqaM4CdAVShcc7uWG4BQhKOYrtBzJ0IKxHGtB9XGijxUyGJZKalRRRCFdkuaIBUwDea7QOJZMl8dCaPEGDAJMoMIQcQO8IanuZazOZnuM0t+G/iMV8wv3JrQ5/voEHw38BYRg6vKa+iJbaKzhznaPqw7P+4Ten4/5IOhmMqkisVb3H9SlsnP8AVSbIf6sscat92/chxzdH9EsDxpE2Foe4yLcy7kYMRbc49o6234viNIUvzHC8HOLgs4qrahTi6zoncKIdo4q8PEyEd6sgZs5/ER5/KXRZfbEQNVdg8CVQY8szijT7gpYfAmgAgAAAOxNkzAgQVXKqwtwMsqlhSJaxcBrbMi8zxlDorMqFd5TlCwSNGqhTzKSD3IVKIliOsynUFrMLuYYzHxlViGCsaczSVlOXLHDZMYbZmfyhkN1L8CPEYqTshIZ7jFgVdm4UsX5lVUF85jVFLjQ3iWcljwylpjqCHeJruAbIPPHmKWTwwC1ThqHQD2Y/N+4DnEe8j3AbXn3BVBgTOGI1+6GaJ5mUuNeSMsKdmF7vJC0VoTQ2qIUhXO4XT8A+IA50agch4lG/2TByswLxFVFXaE6ZeEfsQMU+y/clLRf4bhZlftIW2Ha/i41H0VaMC08796gK+mI+Kf8A3Zh2Y5pH4jl//vhjOPeP8xy+6/7Is2O7/oymeEEPxLuTHqZRuokcp9w5ovIkGsHX9kxN+NlQ4o0L2Qcrgz/XXMMv/Bgi6v8Ah2JQ2ffTixl/7oh7cGsF/RGV8II/Il7e+E/fAPQ/wywBmHFKccw7JJusJq6pWIunYvoRvC/70JRROqyEgXyVu/djk289l9jLleQHN68qrzZK+gftOMKVo+8Ewpio/CECGtZlafh/EAi85a32wCgBBzNxw9AnMIOJdy30qZU5p8wE4gUww8pZ3ly+lz8R9wy3BngPUcMZ8kXRmW6QbmoeG+g6cgmg/MLsMQyWQWsxDlMDuLdCwB5l9kJXjDMyCk2d4w2DyQ0sp5lGGtycywpXZlSFjXMEwmwVEVYYTLDITheMR4pcmJrh2hZLGFuWGebq6iirnhkj1Ap3U0Zl7Lu6grkC74mU9DJCwJzQrGJ07TwId8AafuH2n7f3TxnLdcFgiwEWA7wzYQXWPJmVdsO4x9zDKUQvBvsb+ou4ezuOz1twOF4skJBuNhf21AGFNTs5E8R0WG/yAjSJ3Ev95u1rNftBlgyeFgs9TB/EMaAu8dvzDX0Af3neT4B+8KM5xV/ZYJQbh/bHE2dkX/Ykylax/VJhheRfwJsnd/7ItI4+TUuM/wCB2oNkh2/7qPORYBogsf8AjCWFh4H1TL/O9l/ZiKPyw/eKOa2X/uz89g/dilfKsfyzIsef+kcP4Q/mIGf6SOu57p/EPaZ5UIMfIn+Ja0jzb+Ye1/kd5o/WX9xdw+o4Q9CKLv8AbH5h4UVgCPoCCqGPZSKD6TIogut4haaWk4C38XPRyMUcNAiXsRVfce0lBV/mVS+U9f8AjmVceVsnp/NwgG0fsghqi+Jbcu9NQbvuTcviXFubemoJlIspcIsZW6mNTZAcktNwfBhbmX5lyyXUuXLOYQMRsRWsRekQgy7mcBDrVxHaLeIDGLpfzAoBHdkgKjTzLgT4xMijNgrr+YCWs7EiMkc9mFQNOVV7OIaDaWOh8d5ZazDbZ4fUreeRMPhIxNSYdj9RrFS6SLUrQMzSi/mJZoBq8DuMba/A6P5/eajlcpTmTGYeEiDS0KT97NxlPIXQZRw/G4tvJQWA3VXx/txBpau5e+g3wxFBg3NmSjJyVFeC8GR8gqHSBpU3wbP8e0xSurw/xBeYHcfzB7/z+ZT18n9kJa13b/vNMKxd7+04Nmq2zOXK4N89p4OQvwYH6XuayFZGgeTNfhKCL7jxYfhSMeSVwXtsn1D2FVMtHiz7BKWAIofcP5IELEjeB8YSIIGTQ0/TGtS+2SZ8xexED/3HszU+FV/MWbUe/wD2hxfD/ZKFX+P+0dz5h/MNB8k/mFC7PhSzIeT9VLC7jdjfxL7EbNf4m6Oi/wDwzwQAh+0qhZoVj32gFzzcKe9RLwbWfXLCAJ6IXzVx7YD7YVh8ynCYVVPkJBimLGw+4XShpX9lmFOmHNAcAEUiE8jhSP5Q9FXRaPu0SmGtT21IaIGCofNBAFfsCn+SBj3yx6vH++XbrNXdx820/FalyQi7jJowXFl/lTB8pPjO5XGg6Xa9A/VyyPFi/J/JMC1KW3zWCvuUwahwfkLP4lcLhQcfNFRKGnyPhSfUtaLDle9Q0XlrH6/7DEMZup7rfzDr4cZB6FXiY4+I7Jc1rEGeXRcuUOZt4PwwBwIdRHvxAiJbV5lwFECe8RxOxmV9yF+UT5lu8IAwv0EdyYc9OfQZbljsiKxuUwZXUOidU6Ad5e7mGjL5Sx7Yb3R2QIW+nUC43XC/mMIHkBVQy6DsG4lgrNbEKplriPaOAoxXmUFwvnvMp5mr8MuSYWlwj1ppVe4DnNuiAfukaLIliiPmVK6lgZjVmOQbF9RAW1kbslebtrvDtfD9S0fEGArG9eIHEsW1F4POPpmrwMRfY+TjOji9wB5hVNTTf1LO89zPeBtNBFvPiHHvjdfXY8DEOgLQovpz+8A+QcYIfDg+Agyo1Ry9FyfCepjfqcDtmvyYKiq71H2H8jEFXMI/dMftHaBOL34E5rQgg+NCWRZkUHsFI0RHIBfISJSTIKPgU/Uy6qYE/WH1KYlkXD6Jqi8iP2hwKYRUXCIpQs9+34hFQrkoHxCAobwz+IQLumz9rI6ym+Q+YXl3ELD9Twdgo+ncchc3lTMODlQPmojtXLgl+9/mLtqVpf4j8V2ujT8fyQQFSc7PrX4is+uQfmf/ACJ5zYAP97ikEjAmf6+JXrYLafmAJ2QwjFLsd6MQVVq7Kkc8cMKHHuBhV7oy58+tTcFnZBlA/wBd4gDJewZgUbBWCKfY0wL9JRY8v+EsTMDOxv8AEQFvFgE/Mbuou6xcKKsOQQHBgBqLXTX6UHMIzArh9xEgeWF0fJbBeY04Qk1TPJb3YDh9TlA9E/iVTDx6ivlHDDmOmS4HBNoo8zJKPEGphCCDokBAwt66LvoK9L61coejqV5jySc3KSsRxySxU0whuV9zuIGjAkApFcjpiob49pW1rUE0aG1wgJfCxy1l3hkOJaLKO3iBmysjuFjmbZRAQW+8acy0eZWBrkqCaEwYqbME24/3eWnzs5loxAusP7IyU6hrwOD95d5ih/2AljqGYlwVttO6A8ll2t8efMcTu77XYf7uA2rrEt86mg2p2/75gHclWYGhTs3IaoBpqd1O+/8As1Tfcc/75maUdz/fxMQk7mIcB5aZVEB2FzhnttDGn7u5VkV8yuA+4Olad46YLzUynhZbhf3D5B9TIDHvUKilPENy71H+SAVB9dwQwtHtUDaE7hF83skt5Pepb+TIZH0wwROyQVs+ZNnJ3SNxXcuPZIe7AWSRYBDNRZdl04hS3fiZjB8xgVHiGRKvMCL+aaQTUM1r8S5cdypdS5cQc/mbkQPSX1g+YE1/G0ake+iBlkQFaIhzWBF7LJiL4YprUU8QgzCTcB3IB0zI4YpBTTKOwYNyksazEPiPYQoSq5ip2gzaLFVLjCCy+EEkX6SLJcWPVwbLVmFVRQAHvFeefBlO3bLtq93AbpQ49y4jl2mjRriYVwlMLQ4yzum+/ECQBr1LpDR7QlqxXEaSgC88wNOTUWJ8MWMoMp2Ymb/cdXiCksWDJ5jSZQz0IMBK64IlMtaLl3Y4GWd2HJ6wqjFXllvm4LwKxj3KqbqNql+Itmnuo4al+Iv/AEIhm30iF0viPO/CazpEgGZAun1AuP1DQH1PF+oSA9QI1lHGoGawHhM2kf8AgQHBP/M6QV4+qFn8UD/qgevqgXD6lHY+oUmuiUR15iCi3x0JpRA0A/8AhQiW0m/H3A9fuAqL5gbYdsmHBTuYPzHx7RbFqRyXR+Jl5v2ymvozAdGqWITzSKgQ8sqEEdUzWiNG420z5IdwTDu5fuA1zAgWAlz7wDxKOHEFJRmIwI5Rw7Te7h2IMNRFgMMbmMLYgGXBz0WXLlksrnaRhyPZ2TFG/cY9p6ed3AoMPMpK47zGJj3AO33Diy8wr2b2RL0+mFa9mac+Iwbjm4CEsdxxd3vzGAbgYIoAjCsx894a8FAEWIsG5fmWy1Il2kqlSKwdiDuCUcSAIAMlQgAwdNIBlIpL63BmpzLCWRbZXxLMW8y7w+pY5/EKG/xMuYv/AMT/AFUw5leZqXiEuDOJrpqXiDmViXXMQcxLYixwmwM4/wDc/uKBql8y2wncxAMncfzAET5VH8xFGXZb+YiobxYPol8sX7hop8zn30kB39Ihao4vKPueZLj45NRIT7bbB9MwTPtTU8lOPuGqhBTPeYMIX0w5ZA2krxK4MLiXO8EaZitZ3kBAsqwaUeY1EzaNuI/CGvcAwMrGVQ2QW4M9QrnrVgGZguMMM6xsNQsF5yTCXbtGCL9x2AlP3MRoeY5xpizkPcypR+zNFbHEKhZdXY8VGCFJFEMv3gEFLnMZALuyACyDUX1BeBzD7i+SHfiOH3FyKms8eZoA2J4ks4QbjLuH3EuH3PEnaJVw+494IdlEcIExA/M80r5JZzM24DknkIhyR7pM+ydlGeeF+yCcn3BuSXcwbk+55Ijz+ZRz+ZRyfc80E5lJdxCPhnNEHyfcJwIHh9wRxPmf2JDn50YNr1lD8wJsO0GkF+WJlY8GoNteF1BtpfmANH1B7CIPD4ndE20oYDCrWL6r8QO2IXMMG+L8QlDfZGg1vYi6/BxFAyctMFARoD+8RAsOTMRg3niAiUmlgwlj5j1jHuYm9wBhhYsYhKHMW9wQSwlkDAPMr3gO8WGWM3cIKqFAF5ie8VzAP/sClT7TLxBqClu0ZDbvcWtVcaW3PEtWb7MJ5r5g5N9zmEfmaC/Mds/mXxbSZnN+5mgVAvvzFNbJ21/EsZZVa1ABDzaXezzbv95smp5Q/vAD8OyY9u+5EX3m7n8RAfT/AOJSVg2gxVg+7P5gx4llLADMIV/g/uUpt9P7g2PoP7iyvr/6jQAdjJtiV+fX+cy9j4s/uIWr4TdIef8AicVe1/UJjS+X9RjoPJ/xHQE8/wDiCmCMgB9xUwCe48Q+5xQPmJ5H3BNXqW7HwQ4X6nZk9QUzR9Ts31A9F+JwjpJ5Qy9sIrknARheCdkjMcOIbpHhYvdZwx6ie35rozK7F5UecuEuE8K5YafUW0ZYwJyB8Qb4srwKnYEFxGjRD4SnIRHFECM7kslnZiV1UCv5gCb4qAeLgVAwJFX4llWHFCfJE1K+WJfEBwtSpMZt2RGrtwYsnzBMH3LAF+Zu0sxAk3HlDmD4ZiS2ZYQN5iicjBSY3MUvMc9w+8O9xUd1K0uWTESo0zmUquYrlA3cx3pMe6YbC2+pnxMwT7g09BTS/mZmM+ovT9kdvmV8EvBL8xRN4ClhmJIlORlY7cwDRqptYIBbgJvO52lqsiRcrMMVlainBnipV8hlFkgHcCMIOBzGo6+ZSd4HJLT+4WF15lhvMyZI+IwXRMZZHywMuRjJ8zssodRB/qAj3nNwZ5h3wzwwRss7y3I48wHqC51KIy+JVAcMdViZu0yVEDeIINwp5mfEMNzHuK4l+mMLiN8xX16hV3Cu4tv95fzH8wV7YfwlDAPad5ZLdMo8wGO5kgH9SziWN4jYwxK1iJUTMIO2X7hmQ/EIcTyEzl14qKm7+SLD5aN4aHZlINx3ZTVh5lUo+4SAWFaRGCpTApuXdoC+nITyzDACCjdRUMfLDjTL3is1+4JgfMoC4tIrBoTJle8Dvl8S/wCCexHsSo3L3ciCmoA4uB3z3iXnczpnxELdYhilx0aqk3/7Aa3cHTkiJvEUisO7jpi1P2uClw44OIWxDDczgIGVGeSLwhrTCCpuVtUnMEI2O8UCFEDUcQRdXiGhguATLFfMXojtrtMQ5lrDLnIR3H4jOSvcMLyo5QDu5oWchKPIyq8kv6j4Z/4sG8YlmHcogNjO5AG4v3iFICzFQ7SoJZzU33O/BswY25qNucR8sQKqCH+5UYaqIWGIuOYs5n/qSvlirBil1KXDcRSu4xy/mLw3OduHlHwl3EhuEDeViE4ieIV4lhFTDDYrhN8I5ZR8Tfh6SahHmDQXzOShkEwAGsqn90qIaZygTJ9wSBapnlgVuUczcNxWF8zGNA5h23HeWTLmpmINwDbmESNU7iunSZf3h8PRFaO4cgYio8cRwluNpaVeY0+uIjxFPctWZYvXxLNvRtM39Q4bm72mtb4lNiyCKftzHFbEdbuAvntDK/aYRxBNLvmpYpqoGWmPkmwxTOPmGOJgtEJDgQmFRH2d5g3nzGBnnMt6ibLmDUZ1pApBsIzLFSG1eUIZ/M7riFuYd78yqUcjLSVHmaxhSYMMu3BEsKjzGXLQiuFloAZhG4kdx7uAr1CU7qao0TvpiDmZKWVlleGKcsPjMIazK3VMUwzADEQtUI4hDmYSlmCRYxBkM4NSspeGKMGIh2mDopvUC8QHeYDxU3Q+Ym6KnbzG7p9JyLGsrK+7BDlApwl7K4FlwRuENy67le72iVkxwRmElbQIyD8wmDRKXc7jLLzPeGFfmDYVoOI2fqMAdcS8iqKhjW5fhL+YgilvcLIz8RcliXvEF2xu3n3HepfhqOi7e0yb5mJolKde5cYqXcmojQYlaMREWtd4DEUd5y7PEKZD1KCizvAVY4hh38wHQTSMeIX4qK7qETj30VdXKHWYNWTAEAcUw4ITHK4kbIaMjtGKukvxcxECm7+YK5Q9YcMyxMDWICQLyYjxmFj4gU8wKhkmO4RHrEHdQ7lguZQVcKVFpmmazuVkE4YjieKU1U7RAVVSl6lTZBGAiARJhkh3oh8MwslgczLiV8RalpYhzuIcwHmUWIN4H5iRB0VSNNwJdceWbOgg4Im6RyyzpRAtNQENPmXD+UACO7uc6O6/eOdn3E1GjgmGbZjfmNncBRKy6+5dK5lKHEr51BwgMvncQd2Rnh3BPaLFjcI7lK1ESht9RW0N5qCELYTPnEHVyw4jKSMLCd+E/djmLjWfUU8wMiH1DJd1KuMkXZrsxvx+IAJcXGJoshAP8QtVmAvJmFwjAgBMs1LeJUcR9IYmDoKDpbiqcPEo0wyL1AbYjIyjL6llZgpuFEq6m0MswpAo7MoTzKqPCPElnEHhLDiF6nJzLnEA6xGOIpU70qnmjbUu45ajVy/Fw0xO+R6WiNcVKumNaiWKGyJwTuRbdympaEoN1M3iAeZR5lxhm3HQc86huoikm+4wEWkqo6itFI7mEKv8oVF4TdM5YhrMrIfiMDsQ0Qw3OwwfeGJmYeont2gHOPUp3Vyqsk8ZqYsmUSnzzUVbgCufMtcfiWbuGVynjIJfMWDvDeHcGY+okPELLUvK+4tmrPUVuLJWcXNoKJapMQuPuZu5ATE4A2RL1KkYBxl6QdpQyrHCWMajXeoLNxw8RpuPl0OTI6SGptH4xyjTOo6MSwGDBBbC7wnmAm4TAERKLqVTNRqWETwlE0S6J4CCIwDtKSpA4cSqKFIPfEocyk30QRp9wJnxAQm6iVEPmVfEbeZxQWNdbik8MA8RoufiE0aYTXvEvMymjzKu5fzKMLRKS8ZSaPqJTHvB3IHZZX2/mcwjkMf9ReYFwwbDOX3Kf+oNwWIdDN8xuWQY24Y8jPioXyLEaW9o4Y8zUrMbjcbMAtfMIbuEckN8zkI0+Y0ck2uGPmCLcTFC+Q/EoP7lxdY7SjIP1NdXHvsiqk/EcEdHPiVri4VMfiD95zCGufUBN5h8iWCOpuVZmmULRpM+CZyhLD/MtiXG70ax+kpdy2kpEYhtqCDNQqZueSWl3LapmUozCQWNGVNTCE49FDsdCGZEe8FOiMBUr8SnvpygYmpj8xDLCY84TzDtZ9YEYCMAgup4Z90aMwlphAQpC8LFx48TKDqKNQxCDiVME4RDcfE3h0xbOQYJ30gCQGedGTM4yUL2SolDkhhzKvUKrWOmjfQAPnhli7jYPE7iAehVeoADvKsaEfCmLxEd4ZriAcw84mbUddfMo4hdgqKifdPJBGZZ6ium4UKcwDU8l9DuspRmUuKB3jeNfc7ZnEXdRVS5DwjDbcbOphcpuyCyop+JQZhIXTzDO8QO88sB2w+UDiJihiorMcFxs4j5dbGIZZMMHMICSnERKwEYolt5lo5jW8xeuiLmJLeZlzcU6YwWeIgLfRZRc3xMkGFJkkRBthSNosyxB4QnicBOEgdoC8XOIxFtiGiXM7iNDbGv/Ut/6iBvHaJkglNNwyzLxVz2iAlzZf1LXE01j9pybgrmChEEjliAYjSdvNQyVmHze8DDTzUu9RE11ArGo0sqEywoTBiHxL0N2/8AJ2ICRcXuFmmGnmWYlnuLcyTLoCyZ/EadLHrKI3gO3QyiPdMZ/CWhzBLbOUxMQY78GNy3meeWcy43AYJEOOi2vmfaYSnmNkPCFoKepYI1mMcovhliXjfmeaJejm/MRO8sYn3jaX4YLW5gzDzhBeoRdxBxGmZdQkgKgw6G5UMomMkqD5htx14iOLnHZ6HCRxbshETRfSpRnHiKPuKJuv4gx3WYWXvLlyje4ylxD/yZ5VPjU3sRzFR3KPMYBcokVQCXufVxVL+Yy0yztHOemqMRUXIobjPiFSOjOCUMy/uIPSEffeYL0jdyyEjcw9FfES4h4jS5xN9He41EOhzRxLlVKHMfnUXFhlnMKmYTzBeZ5MQDmAwRCIRvGxMsZSjUUTn0NuIxhz0I7xMSxpGWRs9Om40seZY4gpkwx3EM7hfc94L4hFGWRfmXMoJDPMzMwZdQ+RAEUTcUXQyRbiGAtFGhm4CCKYjio6C7P3LHNkqMdo2t/Eocy33LuFnxExl7x2Xmo4U5lealFxHpUe44ZviDJTFqF9ILBrPiIPhgv/Es3M25lywXsZnxDXmdlFNbgTGYGCMTSP5isK30B9xE7wUcXA4uEVSmBDvhaK+82eZTKiRIkwjcdTCZDFxB4j5jHjiaBqIUuIw5ZnywWsw2sy3mWu5YbhbmWQpIDGEdKmNI5QZ1BUSS45dBNZRggzLIDlg3PBmKOLlJEiE9TTELxLJd1Gn/ACNL6Tyh855IPRVkD/MSW5jHqBL6CjSN3xFcdMIykzhcCiWNwC3mYazmHHaAczbLplUtxGnxBWYtw1rtANxEqrzEQJSYeYt+JllixdATzLz4m1yhFGMs4qJwjB5iLK3MLbcMOINxCoK6HghGE+eUJ3mGYJSep2iRKiZYIKYIISKOokbIQNwcKek4Vswbg4zA75gu37hvMrdwFTOBqAgEHjfQ7nQplC9JJSS3mMFqNI4SxIkdGJAYntEQgJCKRHDEHUe6BYVcKVKHoF4F6LJZFF7wQ0dFXQWJs6FBLCLKReI+L3FvJ8Qu4K7+pni2F+WZOYYQv6TL6mWNEwZuFDFhm5S0qyJ4hnEwSK4Mf/EWJApjEo7y+Y58sz9S+eCoon+Yizcwc3FDCuIum46TPabQL4lYogv/AJCbbqWJRay1lsIG5USJnoeVRvx0dkUgjbMfjGKRgwLO0CtTQ/UYrMU5jCWsuDOZfUBylkt5gDnqA6A9v0KUTLzMMaxyx0WfLLa6ERpicSqYQfmYfB0Yosf7jbi5X1B94vMGFIReD79LfGowsyRznuOLEzZHhiqLRcp2/ExjNeZed5Y8RUUxWjEK7zwlHqCJLLx0C24519TM4hybhZ/mHhEqDR/EPSGHRWPErxDPiBdfvHUZg/vEdiMNrKlcwj43GziXGZQTCCMs4uc4OoU8yldoUhXjMa6hTMWoYeZi9zKZRjwTHjoyj0ydpTPWOcbYxGcZpMo0lppiqUczLvEo5h1SxXMHC4Z46NkLTfcsamURMsZ+2U8SydiVcSiKk+8RcXJMTtDu1AMKY38RbZ8HuP4jKYNVO70DCHygjXENdGFjR8THO3Cz0EXMo5iiqNp8IOM3LslzLV1MXaXZ/c4kM4VHg45gfmYS8w3BzB+J4Ef5mnEMIwZxBbqD/iHcx4OI1UtqKoMwYLCee8W9x8JlcrxG8sOIKi5YNR5xEVmBE7fid0CV8Sn3DLUrsfcCWjhlCmUtiMMfeOEsj4RtxPBHu1MSOHNzeY+IUlXERjORGhzmeeead5+4DzNeYbWYXeeSFoYx27yxjZ6nD26GNxmU8Qp2I14jZLm5WBZe7hhEMQiRRPEIO6EiImDmDi+hhylsvJ8kXQwxi6HeYpe6nkue8TB17llW3BFXUrZlq7gJT8Qw8xxKq4XKSAlRzY9sTOnMS9fcZdTDKm0Kst6I4/7BiHTFnxDvcHichz7gK7yr4zG/EsZme4Htcq+4IrHQHQMNSsQJUDtK/wCSqMzeobhhgrlEVk3MoxjM7jtG0acdQavRXUfjGsqo9AxVRD1HMWxDmaoVmZ5SWcxcQ+8NMP5hecoXhYYi4hHLpiztxJg6rRKSDLJVy6lzczg3CBeIfUJX3KV5jrEoIymM2yyNlcrIFy2L0K56ExXMdShOPE021K++Jj8dCqFm+YPaGPUI7nyg2xD7jX1PSNJQxsHL0tCzMS/MBDSLE0gyxDEGtahnzCMEzxG9dBlNNS1wKhFFi4QgXK+JWJUDGSVAviBMIlQtCBKj0PTpaS2NNRrGscJZxHDovQYaSvqWIhEOY/ec9zyTRmeeC1mWG4cKQjOZ4jkhGHnExSpY06DfxGeWIuLJ7I0yoWRTuQHEwlRYjmKQTDqo3uvM2dI2VcynoUzZKypFqpba5wLR+0pQ1mW0/aVYgjieieaCcwai8S8cEd9pfeLE7E37PRLOYleIdpXiVPKD9xZIv8xPE7lEweZlK/E2/qWQkuwp0HKMqNRU7Zhl2mPt0EMIRtKqceYBN+4E7yoKZhBNkG4hcDM+IFk/CMsNpRUe+6nw/oV/y4yueCJUac4g4bUqd1K83jzL4xpmXcJNkHlDCFoZzOZZYRYLjyj84zbcaxoaj4QXiVrUWcVDG4V8wJhxCU1EJxFqVMpzKOjZHUctz8MP48S5emXMK9ypZTsigz/qgub5iiZu4ivccteYtvUY86iMyIlD5iF5lrlgYv7Qbgp6ETnOj+4molq1no2QD81ODzE/epSrho9EFFzBpPMTcHJG0u5ODDpVuZlAxOalERaRKImXxN4EWGcsTMWYKVLpJuo4+4aZeIKC2RGCRgwcsFjYIcQa+puvMOYODzHDUFESUJcEBAFJV37qIQ8wFRECoMQYgsYaL5qOEy5zKQ3MLM77iY2Z30NWYirjfiLLVcVFcxiK6i4Z2iEuAticREAIAHn+o2v1EU+Ilg3NRZhmIWiZO0XbFiaijwxOUSETnPEbHpzEYrFzGXnRKS3x0P/Z",
        "alt": "Picture",
        "width": 643,
        "height": 428,
        "align": "float-left"
      }
    },
    {
      "type": "paragraph"
    }
  ]
},
  },
  {
    id: 'post-3',
    title: 'Central Banks Signal Cautious Approach to Rate Cuts',
    author: 'Marta Kowalski',
    publishedAt: '2026-07-21',
    category: 'Rates',
    html: post3Html,
    json: {
  "type": "doc",
  "content": [
    {
      "type": "heading",
      "attrs": {
        "level": 2
      },
      "content": [
        {
          "type": "text",
          "text": "Combined Summary"
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "This summary was generated from 1 selected article:"
        }
      ]
    },
    {
      "type": "bulletList",
      "content": [
        {
          "type": "listItem",
          "content": [
            {
              "type": "paragraph",
              "content": [
                {
                  "type": "text",
                  "text": "Consumer Spending Resilient Despite Rate Pressures"
                },
                {
                  "type": "text",
                  "text": " — Retail sales data points to steady discretionary spending, though savings rates continue to decline from pandemic-era highs."
                }
              ]
            }
          ]
        }
      ]
    },
    {
      "type": "heading",
      "attrs": {
        "level": 1,
        "textAlign": "center"
      },
      "content": [
        {
          "type": "text",
          "text": "Taken together, these pieces point to a market environment balancing near-term caution with structural tailwinds worth monitoring closely."
        },
        {
          "type": "text",
          "text": "Text Formatting",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#365F91"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "heading",
      "attrs": {
        "level": 2
      },
      "content": [
        {
          "type": "text",
          "text": "Inline formatting",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#4F81BD"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Here, we demonstrate various types of inline text formatting and the use of embedded fonts."
        },
        {
          "type": "text",
          "text": "cc",
          "marks": [
            {
              "type": "subscript"
            }
          ]
        },
        {
          "type": "text",
          "text": "sdd",
          "marks": [
            {
              "type": "superscript"
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Here is some "
        },
        {
          "type": "text",
          "text": "bold, "
        },
        {
          "type": "text",
          "text": "italic, "
        },
        {
          "type": "text",
          "text": "bold-italic, ",
          "marks": [
            {
              "type": "bold"
            }
          ]
        },
        {
          "type": "text",
          "text": "underlined ",
          "marks": [
            {
              "type": "underline"
            }
          ]
        },
        {
          "type": "text",
          "text": "and "
        },
        {
          "type": "text",
          "text": "struck out ",
          "marks": [
            {
              "type": "strike"
            }
          ]
        },
        {
          "type": "text",
          "text": " text. Then, we have a superscript and a subscript. Now we see some "
        },
        {
          "type": "text",
          "text": "red",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#FF0000"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": ", "
        },
        {
          "type": "text",
          "text": "green",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#92D050"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": " and "
        },
        {
          "type": "text",
          "text": "blue",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#0070C0"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": " text. Some text with a yellow highlight. Some text in a box. Some text in "
        },
        {
          "type": "text",
          "text": "inverse video",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#FFFFFF"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": "."
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "A paragraph with styled text: "
        },
        {
          "type": "text",
          "text": "subtle emphasis  ",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#808080"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": "followed by "
        },
        {
          "type": "text",
          "text": "strong text "
        },
        {
          "type": "text",
          "text": "and "
        },
        {
          "type": "text",
          "text": "intense emphasis",
          "marks": [
            {
              "type": "bold"
            },
            {
              "type": "textStyle",
              "attrs": {
                "color": "#4F81BD"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": ". This paragraph uses document wide styles for styling rather than inline text properties as demonstrated in the previous paragraph — calibre can handle both with equal ease."
        }
      ]
    },
    {
      "type": "heading",
      "attrs": {
        "level": 2
      },
      "content": [
        {
          "type": "text",
          "text": "Fun with fonts",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#4F81BD"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "This document has embedded the Ubuntu font family. The body text is in the Ubuntu typeface, here is some text in the Ubuntu Mono typeface, notice how every letter has the same width, even i and m. Every embedded font will automatically be embedded in the output ebook during conversion."
        }
      ]
    },
    {
      "type": "heading",
      "attrs": {
        "level": 2
      },
      "content": [
        {
          "type": "text",
          "text": "Paragraph level formatting",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#4F81BD"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "paragraph",
      "attrs": {
        "textAlign": "right"
      },
      "content": [
        {
          "type": "text",
          "text": "You can do crazy things with paragraphs, if the urge strikes you. For instance this paragraph is right aligned and has a right border. It has also been given a light gray background.",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#000000"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "For the lovers of poetry amongst you, paragraphs with hanging indents, like this often come in handy. You can use hanging indents to ensure that a line of poetry retains its individual identity as a line even when the screen is  too narrow to display it as a single line. Not only does this paragraph have a hanging indent, it is also has an extra top margin, setting it apart from the preceding paragraph."
        }
      ]
    },
    {
      "type": "heading",
      "attrs": {
        "level": 1
      },
      "content": [
        {
          "type": "text",
          "text": "Tables",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#365F91"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "table",
      "content": [
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "ITEM",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#FFFFFF"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "NEEDED",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#FFFFFF"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Books"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "1"
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Pens"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "3"
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Pencils"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "2"
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Highlighter"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "2 colors"
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Scissors"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "1 pair"
                    }
                  ]
                }
              ]
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Tables in Word can vary from the extremely simple to the extremely complex. calibre tries to do its best when converting tables. While you may run into trouble with the occasional table, the vast majority of common cases should be converted very well, as demonstrated in this section. Note that for optimum results, when creating tables in Word, you should set their widths using percentages, rather than absolute units.  To the left of this paragraph is a floating two column table with a nice green border and header row."
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Now let’s look at a fancier table—one with alternating row colors and partial borders. This table is stretched out to take 100% of the available width."
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "table",
      "content": [
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "City or Town",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "Point A",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "Point B",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "Point C",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "Point D",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "Point E",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Point A",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "—",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " "
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " "
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " "
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " "
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Point B",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "87"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "—"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " "
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " "
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " "
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Point C",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "64",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "56",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "—",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " "
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " "
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Point D",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "37"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "32"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "91"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "—"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " "
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Point E",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "93",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "35",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "54",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "43",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "center"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "—",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Next, we see a table with special formatting in various locations. Notice how the formatting for the header row and sub header rows is preserved."
        }
      ]
    },
    {
      "type": "table",
      "content": [
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "College",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#FFFFFF"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "New students",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#FFFFFF"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Graduating Students",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#FFFFFF"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Change",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#FFFFFF"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": " "
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Undergraduate",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#808080"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": " "
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": " "
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Cedar University"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "110"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "103"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "+7"
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Oak Institute"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "202"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "210"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "-8"
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": " "
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Graduate",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#808080"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": " "
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": " "
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Cedar University"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "24"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "20"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "+4"
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Elm College"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "43"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "53"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "-10"
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Total",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "998",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "908",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "90",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#000000"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Source:",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#808080"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": " Fictitious data, for illustration purposes only"
        }
      ]
    },
    {
      "type": "orderedList",
      "content": [
        {
          "type": "listItem",
          "content": [
            {
              "type": "paragraph",
              "content": [
                {
                  "type": "text",
                  "text": "Next, we have something a little more complex, a nested table, i.e. a table inside another table. Additionally, the inner table has some of its cells merged. The table is displayed horizontally centered."
                }
              ]
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "table",
      "content": [
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": []
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "One"
                    }
                  ]
                },
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Three"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Two"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": []
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Four"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "To the left is a table inside a table, with some cells merged."
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "One"
                    }
                  ]
                },
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Three"
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Two"
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": []
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "Four"
                    }
                  ]
                }
              ]
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "We end with a fancy calendar, note how much of the original formatting is preserved. Note that this table will only display correctly on relatively wide screens. In general, very wide tables or tables whose cells have fixed width requirements don’t fare well in ebooks."
        }
      ]
    },
    {
      "type": "table",
      "content": [
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "attrs": {
                "colspan": 13
              },
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "December 2007",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": " "
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "Sun",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "Mon",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "Tue",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "Wed",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "Thu",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "Fri",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "attrs": {
                "colspan": 2
              },
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "Sat",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "attrs": {
                "colspan": 2
              },
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "1",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "attrs": {
                "colspan": 2
              },
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "2",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "3",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "4",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "5",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "6",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "7",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "attrs": {
                "colspan": 2
              },
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "8",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "attrs": {
                "colspan": 2
              },
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "9",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "10",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "11",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "12",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "13",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "14",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "attrs": {
                "colspan": 2
              },
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "15",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "attrs": {
                "colspan": 2
              },
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "16",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "17",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "18",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "19",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "20",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "21",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "attrs": {
                "colspan": 2
              },
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "22",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "attrs": {
                "colspan": 2
              },
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "23",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "24",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "25",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "26",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "27",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "28",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "attrs": {
                "colspan": 2
              },
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "29",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "attrs": {
                "colspan": 2
              },
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "30",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": "31",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#7F7F7F"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            },
            {
              "type": "tableCell",
              "attrs": {
                "colspan": 2
              },
              "content": [
                {
                  "type": "paragraph",
                  "attrs": {
                    "textAlign": "right"
                  },
                  "content": [
                    {
                      "type": "text",
                      "text": " ",
                      "marks": [
                        {
                          "type": "textStyle",
                          "attrs": {
                            "color": "#365F91"
                          }
                        }
                      ]
                    }
                  ]
                }
              ]
            }
          ]
        },
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": []
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": []
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": []
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": []
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": []
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": []
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": []
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": []
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": []
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": []
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": []
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": []
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": []
                }
              ]
            },
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": []
                }
              ]
            }
          ]
        }
      ]
    },
    {
      "type": "heading",
      "attrs": {
        "level": 1,
        "textAlign": "center"
      },
      "content": [
        {
          "type": "text",
          "text": "Structural Elements",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#365F91"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Miscellaneous structural elements you can add to your document, like footnotes, endnotes, dropcaps and the like."
        }
      ]
    },
    {
      "type": "heading",
      "attrs": {
        "level": 2
      },
      "content": [
        {
          "type": "text",
          "text": "Footnotes & Endnotes",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#4F81BD"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Footnotes"
        },
        {
          "type": "text",
          "text": "[1]"
        },
        {
          "type": "text",
          "text": " and endnotes"
        },
        {
          "type": "text",
          "text": "[i]"
        },
        {
          "type": "text",
          "text": " are automatically recognized and both are converted to endnotes, with backlinks for maximum ease of use in ebook devices."
        }
      ]
    },
    {
      "type": "heading",
      "attrs": {
        "level": 2
      },
      "content": [
        {
          "type": "text",
          "text": "Dropcaps",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#4F81BD"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "table",
      "content": [
        {
          "type": "tableRow",
          "content": [
            {
              "type": "tableCell",
              "content": [
                {
                  "type": "paragraph",
                  "content": [
                    {
                      "type": "text",
                      "text": "D"
                    }
                  ]
                }
              ]
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "rop caps are used to emphasize the leading paragraph at the start of a section. In Word it is possible to specify how many lines of text a drop-cap should use. Because of limitations in ebook technology, this is not possible when converting.  Instead, the converted drop cap will use font size and line height to simulate the effect as well as possible. While not as good as the original, the result is usually tolerable. This paragraph has a “D” dropcap set to occupy three lines of text with a font size of 58.5 pts. Depending on the screen width and capabilities of the device you view the book on, this dropcap can look anything from perfect to ugly."
        }
      ]
    },
    {
      "type": "heading",
      "attrs": {
        "level": 2
      },
      "content": [
        {
          "type": "text",
          "text": "Links",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#4F81BD"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Two kinds of links are possible, those that refer to an external website and those that refer to locations inside the document itself. Both are supported by calibre. For example, here is a link pointing to the "
        },
        {
          "type": "text",
          "text": "calibre download page",
          "marks": [
            {
              "type": "link",
              "attrs": {
                "href": "http://calibre-ebook.com/download"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": ". Then we have a link that points back to the section on "
        },
        {
          "type": "text",
          "text": "paragraph level formatting"
        },
        {
          "type": "text",
          "text": " in this document."
        }
      ]
    },
    {
      "type": "heading",
      "attrs": {
        "level": 2
      },
      "content": [
        {
          "type": "text",
          "text": "Table of Contents",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#4F81BD"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "There are two approaches that calibre takes when generating a Table of Contents. The first is if the Word document has a Table of Contents itself. Provided that the Table of Contents uses hyperlinks, calibre will automatically use it. The levels of the Table of Contents are identified by their left indent, so if you want the ebook to have a multi-level Table of Contents, make sure you create a properly indented Table of Contents in Word."
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "If no Table of Contents is found in the document, then a table of contents is automatically generated from the headings in the document. A heading is identified as something that has the Heading 1 or Heading 2, etc. style applied to it. These headings are turned into a Table of Contents with Heading 1 being the topmost level, Heading 2 the second level and so on."
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " You can see the Table of Contents created by calibre by clicking the Table of Contents button in whatever viewer you are using to view the converted ebook."
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Demonstration of DOCX support in calibre",
          "marks": [
            {
              "type": "underline"
            },
            {
              "type": "textStyle",
              "attrs": {
                "color": "#0000FF"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": ". 1",
          "marks": [
            {
              "type": "underline"
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Text Formatting",
          "marks": [
            {
              "type": "underline"
            },
            {
              "type": "textStyle",
              "attrs": {
                "color": "#0000FF"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": ".. 2",
          "marks": [
            {
              "type": "underline"
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Inline formatting",
          "marks": [
            {
              "type": "underline"
            },
            {
              "type": "textStyle",
              "attrs": {
                "color": "#0000FF"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": ". 2",
          "marks": [
            {
              "type": "underline"
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Fun with fonts",
          "marks": [
            {
              "type": "underline"
            },
            {
              "type": "textStyle",
              "attrs": {
                "color": "#0000FF"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": ". 2",
          "marks": [
            {
              "type": "underline"
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Paragraph level formatting",
          "marks": [
            {
              "type": "underline"
            },
            {
              "type": "textStyle",
              "attrs": {
                "color": "#0000FF"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": ". 2",
          "marks": [
            {
              "type": "underline"
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Tables",
          "marks": [
            {
              "type": "underline"
            },
            {
              "type": "textStyle",
              "attrs": {
                "color": "#0000FF"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": ". 3",
          "marks": [
            {
              "type": "underline"
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Structural Elements",
          "marks": [
            {
              "type": "underline"
            },
            {
              "type": "textStyle",
              "attrs": {
                "color": "#0000FF"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": ". 5",
          "marks": [
            {
              "type": "underline"
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Footnotes & Endnotes",
          "marks": [
            {
              "type": "underline"
            },
            {
              "type": "textStyle",
              "attrs": {
                "color": "#0000FF"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": ". 5",
          "marks": [
            {
              "type": "underline"
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Dropcaps",
          "marks": [
            {
              "type": "underline"
            },
            {
              "type": "textStyle",
              "attrs": {
                "color": "#0000FF"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": ". 5",
          "marks": [
            {
              "type": "underline"
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Links",
          "marks": [
            {
              "type": "underline"
            },
            {
              "type": "textStyle",
              "attrs": {
                "color": "#0000FF"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": ". 5",
          "marks": [
            {
              "type": "underline"
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Table of Contents",
          "marks": [
            {
              "type": "underline"
            },
            {
              "type": "textStyle",
              "attrs": {
                "color": "#0000FF"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": ". 5",
          "marks": [
            {
              "type": "underline"
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Images",
          "marks": [
            {
              "type": "underline"
            },
            {
              "type": "textStyle",
              "attrs": {
                "color": "#0000FF"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": ". 7",
          "marks": [
            {
              "type": "underline"
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Lists",
          "marks": [
            {
              "type": "underline"
            },
            {
              "type": "textStyle",
              "attrs": {
                "color": "#0000FF"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": ". 8",
          "marks": [
            {
              "type": "underline"
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Bulleted List",
          "marks": [
            {
              "type": "underline"
            },
            {
              "type": "textStyle",
              "attrs": {
                "color": "#0000FF"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": " 8",
          "marks": [
            {
              "type": "underline"
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Numbered List",
          "marks": [
            {
              "type": "underline"
            },
            {
              "type": "textStyle",
              "attrs": {
                "color": "#0000FF"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": " 8",
          "marks": [
            {
              "type": "underline"
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Multi-level Lists",
          "marks": [
            {
              "type": "underline"
            },
            {
              "type": "textStyle",
              "attrs": {
                "color": "#0000FF"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": ". 8",
          "marks": [
            {
              "type": "underline"
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Continued Lists",
          "marks": [
            {
              "type": "underline"
            },
            {
              "type": "textStyle",
              "attrs": {
                "color": "#0000FF"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": ". 8",
          "marks": [
            {
              "type": "underline"
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "heading",
      "attrs": {
        "level": 1,
        "textAlign": "center"
      },
      "content": [
        {
          "type": "text",
          "text": "Images",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#365F91"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Images can be of three main types. Inline images are images that are part of the normal text flow, like this image of a green dot"
        }
      ]
    },
    {
      "type": "image",
      "attrs": {
        "src": "data:image/png;base64,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",
        "alt": "Picture 4",
        "width": 14,
        "height": 14
      }
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": ". Inline images do not cause breaks in the text and are usually small in size."
        }
      ]
    },
    {
      "type": "paragraph"
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "The next category of image is a floating image, one that “floats “ on the page and is surrounded by text. Word supports more types of floating images than are possible with current ebook technology, so the conversion maps floating images to simple left and right floats, as you can see with the left and right arrow images on the sides of this paragraph."
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "The final type of image is a “block” image, one that becomes a paragraph on its own and has no text on either side. Below is a centered green dot."
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "paragraph"
    },
    {
      "type": "paragraph"
    },
    {
      "type": "paragraph"
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "Centered images like this are useful for large pictures that should be a focus of attention. Generally, it is not possible to translate the exact positioning of images from a Word document to an ebook. That is because in Word, image positioning is specified in absolute units from the page boundaries.  There is no analogous technology in ebooks, so the conversion will usually end up placing the image either centered or floating close to the point in the text where it was "
        },
        {
          "type": "text",
          "text": "inserted",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#808080"
              }
            }
          ]
        },
        {
          "type": "text",
          "text": ", not necessarily where it appears on the page in Word."
        }
      ]
    },
    {
      "type": "heading",
      "attrs": {
        "level": 1,
        "textAlign": "center"
      },
      "content": [
        {
          "type": "text",
          "text": "Lists",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#365F91"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "All types of lists are supported by the conversion, with the exception of lists that use fancy bullets, these get converted to regular bullets."
        }
      ]
    },
    {
      "type": "heading",
      "attrs": {
        "level": 2
      },
      "content": [
        {
          "type": "text",
          "text": "Bulleted List",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#4F81BD"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "·"
        },
        {
          "type": "text",
          "text": "        "
        },
        {
          "type": "text",
          "text": " "
        },
        {
          "type": "text",
          "text": "One"
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "·"
        },
        {
          "type": "text",
          "text": "        "
        },
        {
          "type": "text",
          "text": " "
        },
        {
          "type": "text",
          "text": "Two"
        }
      ]
    },
    {
      "type": "heading",
      "attrs": {
        "level": 2
      },
      "content": [
        {
          "type": "text",
          "text": "Numbered List",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#4F81BD"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "1."
        },
        {
          "type": "text",
          "text": "    "
        },
        {
          "type": "text",
          "text": "One, with a very long line to demonstrate that the hanging indent for the list is working correctly"
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "2."
        },
        {
          "type": "text",
          "text": "    "
        },
        {
          "type": "text",
          "text": "Two"
        }
      ]
    },
    {
      "type": "heading",
      "attrs": {
        "level": 2
      },
      "content": [
        {
          "type": "text",
          "text": "Multi-level Lists",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#4F81BD"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "1."
        },
        {
          "type": "text",
          "text": "    "
        },
        {
          "type": "text",
          "text": "One"
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "1.1."
        },
        {
          "type": "text",
          "text": "  "
        },
        {
          "type": "text",
          "text": "Two"
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "1.1.1."
        },
        {
          "type": "text",
          "text": "            "
        },
        {
          "type": "text",
          "text": "Three"
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "1.1.2."
        },
        {
          "type": "text",
          "text": "            "
        },
        {
          "type": "text",
          "text": "Four with a very long line to demonstrate that the hanging indent for the list is working correctly."
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "1.1.3."
        },
        {
          "type": "text",
          "text": "            "
        },
        {
          "type": "text",
          "text": "Five"
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "2."
        },
        {
          "type": "text",
          "text": "    "
        },
        {
          "type": "text",
          "text": "Six"
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "A Multi-level list with bullets:"
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "§"
        },
        {
          "type": "text",
          "text": " "
        },
        {
          "type": "text",
          "text": " "
        },
        {
          "type": "text",
          "text": "One"
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "·"
        },
        {
          "type": "text",
          "text": "        "
        },
        {
          "type": "text",
          "text": " "
        },
        {
          "type": "text",
          "text": "Two"
        }
      ]
    },
    {
      "type": "paragraph"
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "     "
        },
        {
          "type": "text",
          "text": " "
        },
        {
          "type": "text",
          "text": "This bullet uses an image as the bullet item"
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "o"
        },
        {
          "type": "text",
          "text": "   "
        },
        {
          "type": "text",
          "text": "Four"
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "§"
        },
        {
          "type": "text",
          "text": " "
        },
        {
          "type": "text",
          "text": " "
        },
        {
          "type": "text",
          "text": "Five"
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "heading",
      "attrs": {
        "level": 2
      },
      "content": [
        {
          "type": "text",
          "text": "Continued Lists",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#4F81BD"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "               "
        },
        {
          "type": "text",
          "text": "i."
        },
        {
          "type": "text",
          "text": "         "
        },
        {
          "type": "text",
          "text": "One"
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "              "
        },
        {
          "type": "text",
          "text": "ii."
        },
        {
          "type": "text",
          "text": "        "
        },
        {
          "type": "text",
          "text": "Two"
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "An interruption in our regularly scheduled listing, for this essential and very relevant public service announcement."
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "             "
        },
        {
          "type": "text",
          "text": "iii."
        },
        {
          "type": "text",
          "text": "        "
        },
        {
          "type": "text",
          "text": "We now resume our normal programming"
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "             "
        },
        {
          "type": "text",
          "text": "iv."
        },
        {
          "type": "text",
          "text": "        "
        },
        {
          "type": "text",
          "text": "Four"
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " ",
          "marks": [
            {
              "type": "textStyle",
              "attrs": {
                "color": "#808080"
              }
            }
          ]
        }
      ]
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": " "
        }
      ]
    },
    {
      "type": "paragraph"
    },
    {
      "type": "paragraph"
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "[1]"
        },
        {
          "type": "text",
          "text": " In paged media, footnotes are usually displayed at the bottom of the text. However, in ebooks, a better paradigm is to make them clickable endnotes that the user can browse at her pleasure. This conversion is handled automatically by calibre."
        }
      ]
    },
    {
      "type": "paragraph"
    },
    {
      "type": "paragraph"
    },
    {
      "type": "paragraph",
      "content": [
        {
          "type": "text",
          "text": "[i]"
        },
        {
          "type": "text",
          "text": " Endnotes are typically used for longer notes, they remain endnotes when converted into ebook form, except that they have an additional backlink to make it easy to return to the current position after reading the note."
        }
      ]
    },
    {
      "type": "paragraph"
    }
  ]
},
  },
  {
    id: 'post-4',
    title: 'Custom Made Word Document to test varioud features',
    author: 'Kenji Watanabe',
    publishedAt: '2026-07-18',
    category: 'Policy',
    html: post4Html,
    json: htmlToTiptapJson(post4Html),
  },
  {
    id: 'post-5',
    title: 'Semiconductor Demand Rebounds on AI Infrastructure Buildout',
    author: 'Sofia Bianchi',
    publishedAt: '2026-07-14',
    category: 'Technology',
    html: post5Html,
    json: htmlToTiptapJson(post5Html),
  },
];
