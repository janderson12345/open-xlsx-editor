// The HTML pasted into Excel by excel-clip-paste.ps1 (written to clip-html-cases.json): made-up tables in the forms
// Excel, Google Sheets, a browser and Word put on the clipboard (each program's HTML flavour, as seen in their own
// copies: Excel's from excel-clip-html.ps1), with the plain text beside it. `read` the cells read back from B2 on.
// usage: node tools/excel-compare/clip-html-cases.mjs
import { writeFileSync } from 'node:fs';

const excelHead = `<html xmlns:o="urn:schemas-microsoft-com:office:office"
xmlns:x="urn:schemas-microsoft-com:office:excel"
xmlns="http://www.w3.org/TR/REC-html40">
<head>
<meta http-equiv=Content-Type content="text/html; charset=utf-8">
<meta name=ProgId content=Excel.Sheet>
<meta name=Generator content="Microsoft Excel 15">
<style>
<!--table
	{mso-displayed-decimal-separator:"\\.";
	mso-displayed-thousand-separator:"\\,";}
br
	{mso-data-placement:same-cell;}
td
	{padding-top:1px;
	padding-right:1px;
	padding-left:1px;
	mso-ignore:padding;
	color:black;
	font-size:11.0pt;
	font-weight:400;
	font-style:normal;
	text-decoration:none;
	font-family:"Aptos Narrow", sans-serif;
	mso-font-charset:0;
	mso-number-format:General;
	text-align:general;
	vertical-align:bottom;
	border:none;
	mso-background-source:auto;
	mso-pattern:auto;
	mso-protection:locked visible;
	white-space:nowrap;
	mso-rotate:0;}
`;

const cases = [
  {
    name: 'excel',
    // (Excel's own copy of a formatted range: classes in the style block, widths on col, merges by colspan/rowspan.)
    html: `${excelHead}.xl63
	{font-weight:700;
	background:#FFC000;
	mso-pattern:black none;}
.xl64
	{mso-number-format:"\\@";}
.xl65
	{mso-number-format:"\\0022$\\0022\\#\\,\\#\\#0\\.00";}
.xl66
	{mso-number-format:"0\\;\\[Red\\]\\\\\\(0\\\\\\)";}
.xl67
	{text-align:left;
	padding-left:18px;
	mso-char-indent-count:2;}
.xl68
	{color:red;
	font-size:14.0pt;
	font-style:italic;
	font-family:Georgia, serif;
	mso-font-charset:0;
	border-top:.5pt solid windowtext;
	border-right:none;
	border-bottom:none;
	border-left:.5pt solid windowtext;}
.xl69
	{vertical-align:top;
	border-top:none;
	border-right:none;
	border-bottom:.5pt solid windowtext;
	border-left:.5pt solid windowtext;
	white-space:normal;}
.xl70
	{text-decoration:underline;
	text-underline-style:single;
	mso-number-format:Standard;
	text-align:center;
	border-top:.5pt solid windowtext;
	border-right:.5pt solid windowtext;
	border-bottom:none;
	border-left:none;}
.xl71
	{mso-number-format:"yyyy\\\\-mm\\\\-dd";
	text-align:center;
	border-top:none;
	border-right:.5pt solid windowtext;
	border-bottom:.5pt solid windowtext;
	border-left:none;}
.xl72
	{mso-number-format:"0\\.0%";
	border-top:none;
	border-right:none;
	border-bottom:2.0pt double blue;
	border-left:none;}
-->
</style>
</head>
<body link="#467886" vlink="#96607D">
<table border=0 cellpadding=0 cellspacing=0 width=371 style='border-collapse:
 collapse;width:279pt'>
<!--StartFragment-->
 <col width=145 style='mso-width-source:userset;mso-width-alt:5302;width:109pt'>
 <col width=73 style='mso-width-source:userset;mso-width-alt:2669;width:55pt'>
 <col width=89 style='mso-width-source:userset;mso-width-alt:3254;width:67pt'>
 <col width=64 style='width:48pt'>
 <tr height=20 style='height:15.0pt'>
  <td height=20 class=xl63 width=145 style='height:15.0pt;width:109pt'>Item</td>
  <td class=xl63 width=73 style='width:55pt'>Qty</td>
  <td class=xl63 width=89 style='width:67pt'>Price</td>
  <td class=xl63 width=64 style='width:48pt'>Share</td>
 </tr>
 <tr height=25 style='height:18.75pt'>
  <td height=25 class=xl68 style='height:18.75pt'>Pears</td>
  <td class=xl70>1,234.50</td>
  <td class=xl65 align=right><s>$3.25</s></td>
  <td class=xl72 align=right>12.5%</td>
 </tr>
 <tr height=40 style='mso-height-source:userset;height:30.0pt'>
  <td height=40 class=xl69 width=145 style='height:30.0pt;width:109pt'>Plums,
  long text</td>
  <td class=xl71>2023-03-15</td>
  <td align=right>4012.125</td>
  <td class=xl64>007</td>
 </tr>
 <tr height=20 style='height:15.0pt'>
  <td colspan=2 height=20 style='height:15.0pt'>Merged across</td>
  <td rowspan=2>Down</td>
  <td class=xl67>indented</td>
 </tr>
 <tr height=20 style='height:15.0pt'>
  <td height=20 class=xl66 align=right style='height:15.0pt'><font
  color="#FF0000" style='mso-ignore:color'>(7)</font></td>
  <td align=center>TRUE</td>
  <td></td>
 </tr>
<!--EndFragment-->
</table>
</body>
</html>`,
    text: 'Item\tQty\tPrice\tShare\r\nPears\t1,234.50\t$3.25\t12.5%\r\n"Plums, long text"\t2023-03-15\t4012.125\t007\r\nMerged across\t\tDown\tindented\r\n(7)\tTRUE\t\t\r\n',
    read: 'B2:E6',
  },
  {
    name: 'excel-xnum',
    // (Excel's Save as Web Page and older copies: the value itself in x:num, x:str, a line break in a cell.)
    html: `${excelHead}.xl65
	{mso-number-format:"0\\.000";}
-->
</style>
</head>
<body>
<table border=0 cellpadding=0 cellspacing=0>
<!--StartFragment-->
 <tr>
  <td class=xl65 align=right x:num="2.71828">2.718</td>
  <td x:str="'0042">0042</td>
  <td>first<br>second</td>
  <td align=right x:num="45000" style='mso-number-format:"Short Date"'>3/15/2023</td>
 </tr>
<!--EndFragment-->
</table>
</body>
</html>`,
    text: '2.718\t0042\t"first\nsecond"\t3/15/2023\r\n',
    read: 'B2:E2',
  },
  {
    name: 'sheets',
    // (Google Sheets: inline styles, data-sheets-value and data-sheets-numberformat as JSON.)
    html: `<meta charset="utf-8"><google-sheets-html-origin><style type="text/css"><!--td {border: 1px solid #cccccc;}br {mso-data-placement:same-cell;}--></style><table xmlns="http://www.w3.org/1999/xhtml" cellspacing="0" cellpadding="0" dir="ltr" border="1" style="table-layout:fixed;font-size:10pt;font-family:Arial;width:0px;border-collapse:collapse;border:none" data-sheets-root="1"><colgroup><col width="100"/><col width="161"/><col width="100"/></colgroup><tbody><tr style="height:21px;"><td style="overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;background-color:#cfe2f3;font-weight:bold;" data-sheets-value="{&quot;1&quot;:2,&quot;2&quot;:&quot;Name&quot;}">Name</td><td style="overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;background-color:#cfe2f3;font-weight:bold;text-align:center;" data-sheets-value="{&quot;1&quot;:2,&quot;2&quot;:&quot;Amount&quot;}">Amount</td><td style="overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;background-color:#cfe2f3;font-weight:bold;" data-sheets-value="{&quot;1&quot;:2,&quot;2&quot;:&quot;Ok&quot;}">Ok</td></tr><tr style="height:21px;"><td style="overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;font-style:italic;color:#990000;" data-sheets-value="{&quot;1&quot;:2,&quot;2&quot;:&quot;Ann&quot;}">Ann</td><td style="overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;text-align:right;" data-sheets-value="{&quot;1&quot;:3,&quot;3&quot;:1234.5}" data-sheets-numberformat="{&quot;1&quot;:2,&quot;2&quot;:&quot;#,##0.00&quot;,&quot;3&quot;:1}">1,234.50</td><td style="overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;text-align:center;" data-sheets-value="{&quot;1&quot;:4,&quot;4&quot;:1}">TRUE</td></tr><tr style="height:21px;"><td style="overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;font-family:Georgia;font-size:14pt;text-decoration:underline line-through;" data-sheets-value="{&quot;1&quot;:2,&quot;2&quot;:&quot;Bo&quot;}">Bo</td><td style="overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;text-align:right;" data-sheets-value="{&quot;1&quot;:3,&quot;3&quot;:0.25}" data-sheets-numberformat="{&quot;1&quot;:3,&quot;2&quot;:&quot;0%&quot;,&quot;3&quot;:1}">25%</td><td style="border-bottom:2px solid #000000;overflow:hidden;padding:2px 3px 2px 3px;vertical-align:bottom;" colspan="1"></td></tr></tbody></table>`,
    text: 'Name\tAmount\tOk\r\nAnn\t1,234.50\tTRUE\r\nBo\t25%\t\r\n',
    read: 'B2:D4',
  },
  {
    name: 'browser',
    // (A web page's table, as a browser copies it: th, inline styles, b and i elements, rgb() colours.)
    html: `<html><body><!--StartFragment--><table style="border-collapse: collapse;"><thead><tr><th style="border: 1px solid rgb(0, 0, 0); background-color: rgb(221, 235, 247);">City</th><th style="border: 1px solid rgb(0, 0, 0);">Pop.</th></tr></thead><tbody><tr><td style="border: 1px solid rgb(0, 0, 0);"><b>Avon</b></td><td style="text-align: right; border: 1px solid rgb(0, 0, 0);">12,400</td></tr><tr><td style="color: rgb(0, 128, 0); font-size: 18px;"><i>Brill</i></td><td style="text-align: center;">3.5</td></tr><tr><td colspan="2" style="font-family: 'Courier New', monospace;">Both <span style="font-weight: 700">columns</span></td></tr></tbody></table><!--EndFragment--></body></html>`,
    text: 'City\tPop.\r\nAvon\t12,400\r\nBrill\t3.5\r\nBoth columns\r\n',
    read: 'B2:C5',
  },
  {
    name: 'word',
    // (A Word table: paragraphs and spans in its cells, the borders and widths on each td.)
    html: `<html xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:w="urn:schemas-microsoft-com:office:word" xmlns="http://www.w3.org/TR/REC-html40"><head><meta http-equiv=Content-Type content="text/html; charset=utf-8"><meta name=ProgId content=Word.Document><style><!--p.MsoNormal{margin:0in;font-size:12.0pt;font-family:"Aptos",sans-serif;}--></style></head><body lang=EN-US><!--StartFragment--><table class=MsoTableGrid border=1 cellspacing=0 cellpadding=0 style='border-collapse:collapse;border:none'><tr><td width=208 valign=top style='width:155.8pt;border:solid windowtext 1.0pt;padding:0in 5.4pt 0in 5.4pt'><p class=MsoNormal><b><span style='font-size:14.0pt'>Task<o:p></o:p></span></b></p></td><td width=208 valign=top style='width:155.8pt;border:solid windowtext 1.0pt;border-left:none;padding:0in 5.4pt 0in 5.4pt'><p class=MsoNormal align=right style='text-align:right'><b>Hours<o:p></o:p></b></p></td></tr><tr><td width=208 valign=top style='width:155.8pt;border:solid windowtext 1.0pt;border-top:none;background:yellow;padding:0in 5.4pt 0in 5.4pt'><p class=MsoNormal><span style='color:#C00000'>Draft<o:p></o:p></span></p></td><td width=208 valign=top style='width:155.8pt;border-top:none;border-left:none;border-bottom:solid windowtext 1.0pt;border-right:solid windowtext 1.0pt;padding:0in 5.4pt 0in 5.4pt'><p class=MsoNormal align=right style='text-align:right'>2.5<o:p></o:p></p></td></tr></table><!--EndFragment--></body></html>`,
    text: 'Task\tHours\r\nDraft\t2.5\r\n',
    read: 'B2:C3',
  },
];

// Paste Values (Paste Special as Text, the same HTML on the clipboard): the text alone, typed in.
cases.push({ ...cases[0], name: 'excel-values', values: true });
cases.push({ ...cases[2], name: 'sheets-values', values: true });

writeFileSync(new URL('./clip-html-cases.json', import.meta.url), `${JSON.stringify(cases, null, 2)}\n`);
console.log(`${cases.length} cases`);
