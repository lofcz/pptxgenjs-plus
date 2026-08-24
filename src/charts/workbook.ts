/** Embedded Excel workbook generation for charts. */

import { JSZip } from '@node-projects/jszip'
import { CHART_STYLE, CHART_TYPE, isChartexType } from '../core-enums'
import { ISlideRelChart } from '../core-interfaces'
import { encodeXmlEntities } from '../gen-utils'
import { makeXmlChartEx } from './chartex'
import { chartColorsPartName, chartStylePartName, makeXmlChartColors, makeXmlChartStyle, wantsChartStyleParts } from './style'
import { makeXmlCharts } from './xml'
import { countErrorrateSeries, getExcelColName, seriesHasErrorrate } from './utils'

function getDataTableStyles (chartObject: ISlideRelChart): { customFormats: string[]; dataStyleIds: number[] } {
	const dataTableFormats: (string | undefined)[] = Array.isArray(chartObject.opts._type)
		? chartObject.opts._type.flatMap(type => Array<string | undefined>(type.data.length).fill(type.options?.dataTableFormatCode ?? chartObject.opts.dataTableFormatCode))
		: chartObject.data.map(() => chartObject.opts.dataTableFormatCode)
	const customFormats = [...new Set(dataTableFormats.filter((format): format is string => !!format))]
	return { customFormats, dataStyleIds: dataTableFormats.map(format => format ? customFormats.indexOf(format) + 1 : 0) }
}

function makeNumberStyles (formats: string[]): string {
	const numFmts = formats.map((format, idx) => `<numFmt numFmtId="${164 + idx}" formatCode="${encodeXmlEntities(format)}"/>`).join('')
	const cellXfs = formats.map((_format, idx) => `<xf numFmtId="${164 + idx}" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>`).join('')
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="${formats.length}">${numFmts}</numFmts><fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="${formats.length + 1}"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>${cellXfs}</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>`
}

/**
 * Based on passed data, creates Excel Worksheet that is used as a data source for a chart.
 * @param {ISlideRelChart} chartObject - chart object
 * @param {JSZip} zip - file that the resulting XLSX should be added to
 * @return {Promise} promise of generating the XLSX file
 */
export async function createExcelWorksheet (chartObject: ISlideRelChart, zip: JSZip): Promise<string> {
	const zipExcel = new JSZip()
	const dataTableStyles = getDataTableStyles(chartObject)
	addWorkbookFolders(zipExcel)
	addCoreWorkbookFiles(zipExcel, dataTableStyles.customFormats)
	addSharedStringsFile(chartObject, zipExcel)
	addTableFile(chartObject, zipExcel)
	addWorksheetFile(chartObject, zipExcel, dataTableStyles.dataStyleIds)
	return await addWorkbookToPresentation(chartObject, zipExcel, zip)
}

/** Create the archive folders required by the embedded XLSX package. */
function addWorkbookFolders (zipExcel: JSZip): void {
	zipExcel.folder('_rels')
	zipExcel.folder('docProps')
	zipExcel.folder('xl/_rels')
	zipExcel.folder('xl/tables')
	zipExcel.folder('xl/theme')
	zipExcel.folder('xl/worksheets')
	zipExcel.folder('xl/worksheets/_rels')
}

/**
 * Write the fixed XLSX package parts: content types, root/workbook relationships, metadata, styles, theme, and workbook.
 * These parts are independent of chart data and must exist before the data-specific files are added.
 */
function addCoreWorkbookFiles (zipExcel: JSZip, customFormats: string[]): void {
	// B: Add core contents
	{
		zipExcel.file(
			'[Content_Types].xml',
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
			'  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
			'  <Default Extension="xml" ContentType="application/xml"/>' +
			'  <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
			'  <Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>' +
			'  <Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>' +
			'  <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
			'  <Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' +
			'  <Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"/>' +
			'  <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>' +
			'  <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>' +
			'</Types>\n'
		)
		zipExcel.file(
			'_rels/.rels',
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
			'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>' +
			'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>' +
			'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
			'</Relationships>\n'
		)
		zipExcel.file(
			'docProps/app.xml',
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">' +
			'<Application>Microsoft Macintosh Excel</Application>' +
			'<DocSecurity>0</DocSecurity>' +
			'<ScaleCrop>false</ScaleCrop>' +
			'<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>1</vt:i4></vt:variant></vt:vector></HeadingPairs>' +
			'<TitlesOfParts><vt:vector size="1" baseType="lpstr"><vt:lpstr>Sheet1</vt:lpstr></vt:vector></TitlesOfParts>' +
			'<Company></Company><LinksUpToDate>false</LinksUpToDate><SharedDoc>false</SharedDoc><HyperlinksChanged>false</HyperlinksChanged><AppVersion>16.0300</AppVersion>' +
			'</Properties>\n'
		)
		zipExcel.file(
			'docProps/core.xml',
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
			'<dc:creator>PptxGenJS</dc:creator>' +
			'<cp:lastModifiedBy>PptxGenJS</cp:lastModifiedBy>' +
			'<dcterms:created xsi:type="dcterms:W3CDTF">' +
			new Date().toISOString() +
			'</dcterms:created>' +
			'<dcterms:modified xsi:type="dcterms:W3CDTF">' +
			new Date().toISOString() +
			'</dcterms:modified>' +
			'</cp:coreProperties>'
		)
		zipExcel.file(
			'xl/_rels/workbook.xml.rels',
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
			'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
			'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
			'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>' +
			'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>' +
			'<Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>' +
			'</Relationships>'
		)
		zipExcel.file(
			'xl/styles.xml',
			customFormats.length > 0 ? makeNumberStyles(customFormats) : '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><numFmts count="1"><numFmt numFmtId="0" formatCode="General"/></numFmts><fonts count="4"><font><sz val="9"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="9"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="10"/><color indexed="8"/><name val="Geneva"/></font><font><sz val="18"/><color indexed="8"/>' +
			'<name val="Arial"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><dxfs count="0"/><tableStyles count="0"/><colors><indexedColors><rgbColor rgb="ff000000"/><rgbColor rgb="ffffffff"/><rgbColor rgb="ffff0000"/><rgbColor rgb="ff00ff00"/><rgbColor rgb="ff0000ff"/>' +
			'<rgbColor rgb="ffffff00"/><rgbColor rgb="ffff00ff"/><rgbColor rgb="ff00ffff"/><rgbColor rgb="ff000000"/><rgbColor rgb="ffffffff"/><rgbColor rgb="ff878787"/><rgbColor rgb="fff9f9f9"/></indexedColors></colors></styleSheet>\n'
		)
		zipExcel.file(
			'xl/theme/theme1.xml',
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF"/></a:lt1><a:dk2><a:srgbClr val="44546A"/></a:dk2><a:lt2><a:srgbClr val="E7E6E6"/></a:lt2><a:accent1><a:srgbClr val="4472C4"/></a:accent1><a:accent2><a:srgbClr val="ED7D31"/></a:accent2><a:accent3><a:srgbClr val="A5A5A5"/></a:accent3><a:accent4><a:srgbClr val="FFC000"/></a:accent4><a:accent5><a:srgbClr val="5B9BD5"/></a:accent5><a:accent6><a:srgbClr val="70AD47"/></a:accent6><a:hlink><a:srgbClr val="0563C1"/></a:hlink><a:folHlink><a:srgbClr val="954F72"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Calibri Light" panose="020F0302020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="DengXian Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:majorFont><a:minorFont><a:latin typeface="Calibri" panose="020F0502020204030204"/><a:ea typeface=""/><a:cs typeface=""/><a:font script="Jpan" typeface="Yu Gothic"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="DengXian"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Tahoma"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}" vid="{4A3C46E8-61CC-4603-A589-7422A47A8E4A}"/></a:ext></a:extLst></a:theme>'
		)
		zipExcel.file(
			'xl/workbook.xml',
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
			'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x15" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main">' +
			'<fileVersion appName="xl" lastEdited="7" lowestEdited="6" rupBuild="10507"/>' +
			'<workbookPr/>' +
			'<bookViews><workbookView xWindow="0" yWindow="500" windowWidth="20960" windowHeight="15960"/></bookViews>' +
			'<sheets><sheet name="Sheet1" sheetId="1" r:id="rId1"/></sheets>' +
			'<calcPr calcId="0" concurrentCalc="0"/>' +
			'</workbook>\n'
		)
		zipExcel.file(
			'xl/worksheets/_rels/sheet1.xml.rels',
			'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
			'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
			'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="../tables/table1.xml"/>' +
			'</Relationships>\n'
		)
	}
}

/**
 * Write `xl/sharedStrings.xml` in the index order consumed by the worksheet writer.
 * Category charts, scatter charts, bubble charts, and multi-level categories intentionally use different layouts.
 */
function addSharedStringsFile (chartObject: ISlideRelChart, zipExcel: JSZip): void {
	const data = chartObject.data
	const intBubbleCols = (data.length - 1) * 2 + 1
	const IS_MULTI_CAT_AXES = (data[0]?.labels?.length ?? 0) > 1
	const firstDataLabels = data[0]?.labels ?? []

	// sharedStrings.xml
	{
		// A: Start XML
		let strSharedStrings = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		if (chartObject.opts._type === CHART_TYPE.BUBBLE || chartObject.opts._type === CHART_TYPE.BUBBLE3D) {
			strSharedStrings += `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${intBubbleCols}" uniqueCount="${intBubbleCols}">`
		} else if (chartObject.opts._type === CHART_TYPE.SCATTER) {
			strSharedStrings += `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${data.length}" uniqueCount="${data.length}">`
		} else if (IS_MULTI_CAT_AXES) {
			// series names + every non-blank label + the leading blank string (emitted below)
			let totCount = data.length + 1
			firstDataLabels.forEach(arrLabel => (totCount += arrLabel.filter(label => label && label !== '').length))
			strSharedStrings += `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${totCount}" uniqueCount="${totCount}">`
			strSharedStrings += '<si><t/></si>'
		} else {
			// series names + all labels of one series + number of label groups (data.labels.length) of one series (i.e. how many times the blank string is used)
			const totCount = data.length + firstDataLabels.length * firstDataLabels[0].length + firstDataLabels.length
			// series names + labels of one series + blank string (same for all label groups)
			const unqCount = data.length + firstDataLabels.length * firstDataLabels[0].length + 1
			// start `sst`
			strSharedStrings += `<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="${totCount}" uniqueCount="${unqCount}">`
			// B: Add 'blank' for A1, B1, ..., of every label group inside data[n].labels
			strSharedStrings += '<si><t xml:space="preserve"></t></si>'
		}

		// C: Add `name`/Series
		if (chartObject.opts._type === CHART_TYPE.BUBBLE || chartObject.opts._type === CHART_TYPE.BUBBLE3D) {
			data.forEach((objData, idx) => {
				if (idx === 0) strSharedStrings += '<si><t>X-Axis</t></si>'
				else {
					strSharedStrings += `<si><t>${encodeXmlEntities(objData.name || `Y-Axis${idx}`)}</t></si>`
					strSharedStrings += `<si><t>${encodeXmlEntities(`Size${idx}`)}</t></si>`
				}
			})
		} else {
			data.forEach(objData => {
				strSharedStrings += `<si><t>${encodeXmlEntities((objData.name || ' ').replace('X-Axis', 'X-Values'))}</t></si>`
			})
		}

		// D: Add `labels`/Categories
		if (chartObject.opts._type !== CHART_TYPE.BUBBLE && chartObject.opts._type !== CHART_TYPE.BUBBLE3D && chartObject.opts._type !== CHART_TYPE.SCATTER) {
			// Use forEach backwards & check for '' to support multi-cat axes
			firstDataLabels
				.slice()
				.reverse()
				.forEach(labelsGroup => {
					labelsGroup
						.filter(label => label && label !== '')
						.forEach(label => {
							strSharedStrings += `<si><t>${encodeXmlEntities(label)}</t></si>`
						})
				})
		}

		// DONE:
		strSharedStrings += '</sst>\n'
		zipExcel.file('xl/sharedStrings.xml', strSharedStrings)
	}
}

/** Write table metadata and the calculated cell range for the embedded workbook. */
function addTableFile (chartObject: ISlideRelChart, zipExcel: JSZip): void {
	const data = chartObject.data
	const intBubbleCols = (data.length - 1) * 2 + 1
	const firstDataLabels = data[0]?.labels ?? []
	const firstDataValues = data[0]?.values ?? []
	const errSeriesCount = countErrorrateSeries(data)
	const barSheetCols = firstDataLabels.length + data.length + errSeriesCount

	// tables/table1.xml
	{
		let strTableXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		if (chartObject.opts._type === CHART_TYPE.BUBBLE || chartObject.opts._type === CHART_TYPE.BUBBLE3D) {
			strTableXml += `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:${getExcelColName(intBubbleCols)}${intBubbleCols}" totalsRowShown="0">`
			strTableXml += `<tableColumns count="${intBubbleCols}">`
			let idxColLtr = 1
			data.forEach((obj, idx) => {
				if (idx === 0) {
					strTableXml += `<tableColumn id="${idx + 1}" name="X-Values"/>`
				} else {
					strTableXml += `<tableColumn id="${idx + idxColLtr}" name="${encodeXmlEntities(obj.name)}"/>`
					idxColLtr++
					strTableXml += `<tableColumn id="${idx + idxColLtr}" name="Size${idx}"/>`
				}
			})
		} else if (chartObject.opts._type === CHART_TYPE.SCATTER) {
			strTableXml += `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:${getExcelColName(data.length)}${firstDataValues.length + 1}" totalsRowShown="0">`
			strTableXml += `<tableColumns count="${data.length}">`
			data.forEach((_obj, idx) => {
				strTableXml += `<tableColumn id="${idx + 1}" name="${idx === 0 ? 'X-Values' : 'Y-Value '}${idx}"/>`
			})
		} else {
			strTableXml += `<table xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" id="1" name="Table1" displayName="Table1" ref="A1:${getExcelColName(barSheetCols)}${firstDataLabels[0].length + 1}" totalsRowShown="0">`
			strTableXml += `<tableColumns count="${barSheetCols}">`
			firstDataLabels.forEach((_labelsGroup, idx) => {
				strTableXml += `<tableColumn id="${idx + 1}" name="Column${idx + 1}"/>`
			})
			data.forEach((obj, idx) => {
				strTableXml += `<tableColumn id="${idx + firstDataLabels.length + 1}" name="${encodeXmlEntities(obj.name)}"/>`
			})
			// Packed errorrate columns (only series that provide non-empty errorrate)
			{
				let errColId = firstDataLabels.length + data.length + 1
				data.forEach(obj => {
					if (!seriesHasErrorrate(obj)) return
					strTableXml += `<tableColumn id="${errColId}" name="${encodeXmlEntities(obj.name || 'Series')}_errorrate"/>`
					errColId++
				})
			}
		}
		strTableXml += '</tableColumns>'
		strTableXml += '<tableStyleInfo showFirstColumn="0" showLastColumn="0" showRowStripes="1" showColumnStripes="0"/>'
		strTableXml += '</table>'
		zipExcel.file('xl/tables/table1.xml', strTableXml)
	}
}

/**
 * Write `xl/worksheets/sheet1.xml`, preserving zero values and multi-level-category merge/shared-string indexes.
 */
function addWorksheetFile (chartObject: ISlideRelChart, zipExcel: JSZip, dataStyleIds: number[]): void {
	const data = chartObject.data
	const intBubbleCols = (data.length - 1) * 2 + 1
	const IS_MULTI_CAT_AXES = (data[0]?.labels?.length ?? 0) > 1
	const firstDataLabels = data[0]?.labels ?? []
	const firstDataValues = data[0]?.values ?? []
	const errSeriesCount = countErrorrateSeries(data)
	const barSheetCols = firstDataLabels.length + data.length + errSeriesCount

	// worksheets/sheet1.xml
	{
		let strSheetXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
		const arrMergeRefs: string[] = [] // multi-cat axes only: vertical merges for the outer label columns
		strSheetXml +=
			'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'

		if (chartObject.opts._type === CHART_TYPE.BUBBLE || chartObject.opts._type === CHART_TYPE.BUBBLE3D) {
			strSheetXml += `<dimension ref="A1:${getExcelColName(intBubbleCols)}${firstDataValues.length + 1}"/>`
		} else if (chartObject.opts._type === CHART_TYPE.SCATTER) {
			strSheetXml += `<dimension ref="A1:${getExcelColName(data.length)}${firstDataValues.length + 1}"/>`
		} else {
			strSheetXml += `<dimension ref="A1:${getExcelColName(barSheetCols)}${(firstDataLabels[0]?.length ?? firstDataValues.length) + 1}"/>`
		}

		strSheetXml += '<sheetViews><sheetView tabSelected="1" workbookViewId="0"><selection activeCell="B1" sqref="B1"/></sheetView></sheetViews>'
		strSheetXml += '<sheetFormatPr baseColWidth="10" defaultRowHeight="16"/>'
		if (chartObject.opts._type === CHART_TYPE.BUBBLE || chartObject.opts._type === CHART_TYPE.BUBBLE3D) {
			// UNUSED: strSheetXml += `<cols><col min="1" max="${data.length}" width="11" customWidth="1" /></cols>`

			/* EX: INPUT: `data`
			[
				{ name:'X-Axis'  , values:[10,11,12,13,14,15,16,17,18,19,20] },
				{ name:'Y-Axis 1', values:[ 1, 6, 7, 8, 9], sizes:[ 4, 5, 6, 7, 8] },
				{ name:'Y-Axis 2', values:[33,32,42,53,63], sizes:[11,12,13,14,15] }
			];
			*/
			/* EX: OUTPUT: bubbleChart Worksheet:
				-|----A-----|------B-----|------C-----|------D-----|------E-----|
				1| X-Values | Y-Values 1 | Y-Sizes 1  | Y-Values 2 | Y-Sizes 2  |
				2|    11    |     22     |      4     |     33     |      8     |
				-|----------|------------|------------|------------|------------|
			*/
			strSheetXml += '<sheetData>'

			// A: Create header row first (NOTE: Start at index=1 as headers cols start with 'B')
			strSheetXml += `<row r="1" spans="1:${intBubbleCols}">`
			strSheetXml += '<c r="A1" t="s"><v>0</v></c>'
			for (let idx = 1; idx < intBubbleCols; idx++) {
				strSheetXml += `<c r="${getExcelColName(idx + 1)}1" t="s"><v>${idx}</v></c>` // NOTE: add `t="s"` for label cols!
			}
			strSheetXml += '</row>'

			// B: Add row for each X-Axis value (Y-Axis* value is optional)
			firstDataValues.forEach((val, idx) => {
				// Leading col is reserved for the 'X-Axis' value, so hard-code it, then loop over col values
				strSheetXml += `<row r="${idx + 2}" spans="1:${intBubbleCols}">`
				strSheetXml += `<c r="A${idx + 2}"><v>${val}</v></c>`
				// Add Y-Axis 1->N (idy=0 = Xaxis)
				let idxColLtr = 2
				for (let idy = 1; idy < data.length; idy++) {
					// NOTE: Preserve 0 (a valid value); `|| ''` would blank it out in the embedded worksheet (issue #1430)
					const yVal = (data[idy].values ?? [])[idx]
					const ySize = (data[idy].sizes ?? [])[idx]
					// y-value
					strSheetXml += `<c r="${getExcelColName(idxColLtr)}${idx + 2}"><v>${typeof yVal === 'number' ? yVal : ''}</v></c>`
					idxColLtr++
					// y-size
					strSheetXml += `<c r="${getExcelColName(idxColLtr)}${idx + 2}"><v>${typeof ySize === 'number' ? ySize : ''}</v></c>`
					idxColLtr++
				}
				strSheetXml += '</row>'
			})
		} else if (chartObject.opts._type === CHART_TYPE.SCATTER) {
			/* UNUSED:
				strSheetXml += '<cols>'
				strSheetXml += '<col min="1" max="' + data.length + '" width="11" customWidth="1" />'
				//data.forEach((obj,idx)=>{ strSheetXml += '<col min="'+(idx+1)+'" max="'+(idx+1)+'" width="11" customWidth="1" />' });
				strSheetXml += '</cols>'
			*/
			/* EX: INPUT: `data`
				[
					{ name:'X-AxisA', values:[ 1, 2, 3, 4, 5] },
					{ name:'Y-AxisB', values:[ 2,22,42,52,62] },
					{ name:'Y-AxisC', values:[ 3,33,43,53,63] }
				];
			*/
			/* EX: OUTPUT: sheet1.xml:
				-|----A----|----B----|----C----|
				1| X-AxisA | Y-AxisB | Y-AxisC |
				2|    1    |    2    |    3    |
				-|---------|---------|---------|
			*/
			strSheetXml += '<sheetData>'

			// A: Create header row first (every `name` row provided)
			strSheetXml += `<row r="1" spans="1:${data.length}">`
			for (let idx = 0; idx < data.length; idx++) {
				strSheetXml += `<c r="${getExcelColName(idx + 1)}1" t="s"><v>${idx}</v></c>` // NOTE: add `t="s"` for label cols!
			}
			strSheetXml += '</row>'

			// B: Add row for each X-Axis value (Y-Axis* value is optional)
			firstDataValues.forEach((val, idx) => {
				// Leading col is reserved for the 'X-Axis' value, so hard-code it, then loop over col values
				strSheetXml += `<row r="${idx + 2}" spans="1:${data.length}">`
				strSheetXml += `<c r="A${idx + 2}"><v>${val}</v></c>`
				// Add Y-Axis 1->N
				for (let idy = 1; idy < data.length; idy++) {
					const idyValues = data[idy].values ?? []
					strSheetXml += `<c r="${getExcelColName(idy + 1)}${idx + 2}"><v>${idyValues[idx] || idyValues[idx] === 0 ? idyValues[idx] : ''
					}</v></c>`
				}
				strSheetXml += '</row>'
			})
		} else {
			// strSheetXml += '<cols><col min="1" max="1" width="11" customWidth="1" /></cols>'
			strSheetXml += '<sheetData>'

			/* EX: INPUT: `data`
				[
					{ name:'Red', labels:['Jan..May-17'], values:[11,13,14,15,16] },
					{ name:'Amb', labels:['Jan..May-17'], values:[22, 6, 7, 8, 9] },
					{ name:'Grn', labels:['Jan..May-17'], values:[33,32,42,53,63] }
				];
			*/
			/* EX: OUTPUT: lineChart Worksheet:
				-|---A---|--B--|--C--|--D--|
				1|       | Red | Amb | Grn |
				2|Jan-17 |   11|   22|   33|
				3|Feb-17 |   55|   43|   70|
				4|Mar-17 |   56|  143|   99|
				5|Apr-17 |   65|    3|  120|
				6|May-17 |   75|   93|  170|
				-|-------|-----|-----|-----|
			*/

			if (!IS_MULTI_CAT_AXES) {
				// A: Create header row first (series names only; errorrate cols are values-only to keep shared-string indices stable)
				strSheetXml += `<row r="1" spans="1:${barSheetCols}">`
				firstDataLabels.forEach((_labelsGroup, idx) => {
					strSheetXml += `<c r="${getExcelColName(idx + 1)}1" t="s"><v>0</v></c>`
				})
				for (let idx = 0; idx < data.length; idx++) {
					strSheetXml += `<c r="${getExcelColName(idx + 1 + firstDataLabels.length)}1" t="s"><v>${idx + 1}</v></c>` // NOTE: use `t="s"` for label cols!
				}
				strSheetXml += '</row>'

				// B: Add data row(s) for each category
				firstDataLabels[0].forEach((_cat, idx) => {
					strSheetXml += `<row r="${idx + 2}" spans="1:${barSheetCols}">`
					// Leading cols are reserved for the label groups
					for (let idx2 = firstDataLabels.length - 1; idx2 >= 0; idx2--) {
						strSheetXml += `<c r="${getExcelColName(firstDataLabels.length - idx2)}${idx + 2}" t="s">`
						strSheetXml += `<v>${data.length + idx + 1}</v>`
						strSheetXml += '</c>'
					}
					for (let idy = 0; idy < data.length; idy++) {
						// Preserve 0 (a valid value); `|| ''` would blank it out in the embedded worksheet (issue #1430)
						const cellVal = (data[idy].values ?? [])[idx]
						const style = dataStyleIds[idy] ? ` s="${dataStyleIds[idy]}"` : ''
						strSheetXml += `<c r="${getExcelColName(firstDataLabels.length + idy + 1)}${idx + 2}"${style}><v>${cellVal || cellVal === 0 ? cellVal : ''}</v></c>`
					}
					// Packed errorrate columns after all series values
					{
						let errPacked = 0
						for (let idy = 0; idy < data.length; idy++) {
							if (!seriesHasErrorrate(data[idy])) continue
							const errVal = (data[idy].errorrate ?? [])[idx]
							const errCol = firstDataLabels.length + data.length + 1 + errPacked
							strSheetXml += `<c r="${getExcelColName(errCol)}${idx + 2}"><v>${errVal || errVal === 0 ? errVal : ''}</v></c>`
							errPacked++
						}
					}
					strSheetXml += '</row>'
				})
			} else {
				// A: create header row: one blank cell per label level, then one cell per series
				strSheetXml += `<row r="1" spans="1:${barSheetCols}">`
				for (let idx = 0; idx < firstDataLabels.length; idx++) {
					strSheetXml += `<c r="${getExcelColName(idx + 1)}1" t="s"><v>0</v></c>`
				}
				for (let idx = 0; idx < data.length; idx++) {
					strSheetXml += `<c r="${getExcelColName(firstDataLabels.length + idx + 1)}1" t="s"><v>${idx + 1}</v></c>` // NOTE: use `t="s"` for label cols!
				}
				strSheetXml += '</row>'

				/**
				 * @example INPUT
				 * const LABELS = [
				 *   ["Gear", "Berg", "Motr", "Swch", "Plug", "Cord", "Pump", "Leak", "Seal"],
				 *   ["Mech", "", "", "Elec", "", "", "Hydr", "", ""],
				 * ];
				 * const arrDataRegions = [
				 *   { name: "West", labels: LABELS, values: [11, 8, 3, 0, 11, 3, 0, 0, 0] },
				 *   { name: "Ctrl", labels: LABELS, values: [0, 11, 6, 19, 12, 5, 0, 0, 0] },
				 *   { name: "East", labels: LABELS, values: [0, 3, 2, 0, 0, 0, 4, 3, 1] },
				 * ];
				 */
				/**
				 * @example OUTPUT EXCEL SHEET
				 * |/|---A--|---B--|---C--|---D--|---E--|
				 * |1|      |      | West | Ctrl | East |
				 * |2| Mech | Gear |  ##  |  ##  |  ##  |
				 * |3|      | Brng |  ##  |  ##  |  ##  |
				 * |4|      | Motr |  ##  |  ##  |  ##  |
				 * |5| Elec | Swch |  ##  |  ##  |  ##  |
				 * |6|      | Plug |  ##  |  ##  |  ##  |
				 * |7|      | Cord |  ##  |  ##  |  ##  |
				 * |8| Hydr | Pump |  ##  |  ##  |  ##  |
				 * |9|      | Leak |  ##  |  ##  |  ##  |
				 *|10|      | Seal |  ##  |  ##  |  ##  |
				 */
				/**
				 * @example OUTPUT EXCEL SHEET XML
				 * <row r="1" spans="1:5">
				 *   <c r="A1" t="s"><v>0</v></c>
				 *   <c r="B1" t="s"><v>0</v></c>
				 *   <c r="C1" t="s"><v>1</v></c>
				 *   <c r="D1" t="s"><v>2</v></c>
				 *   <c r="E1" t="s"><v>3</v></c>
				 * </row>
				 * <row r="2" spans="1:5">
				 *   <c r="A2" t="s"><v>4</v></c>
				 *   <c r="B2" t="s"><v>7</v></c>
				 *   <c r="C2"      ><v>###</v></c>
				 * </row>
				 * <row r="3" spans="1:5">
				 *   <c r="A3" />
				 *   <c r="B3" t="s"><v>8</v></c>
				 *   <c r="C3"      ><v>###</v></c>
				 * </row>
				 */
				/**
				 * @example SHARED-STRINGS
				 * 1=West, 2=Ctrl, 3=East, 4=Mech, 5=Elec, 6=Mydr, 7=Gear, 8=Brng, [...], 15=Seal
				 */

				// B: Add data row(s) for each category
				/**
				 * const LABELS = [
				 *   ["Gear", "Berg", "Motr", "Swch", "Plug", "Cord", "Pump", "Leak", "Seal"],
				 *   ["Mech",     "",     "", "Elec",     "",     "", "Hydr",     "",     ""],
				 *   ["2010",     "",     "",     "",     "",     "",     "",     "",     ""],
				 * ];
				 */
				const TOT_SER = data.length
				const TOT_CAT = firstDataLabels[0].length
				const TOT_LVL = firstDataLabels.length
				/**
				 * Label groups outermost-first, matching both the column order (col 1 = outermost level)
				 * and the shared-strings order written above (blank, series names, then labels in this order).
				 * const LABELS_REVERSED = [
				 *   ["Mech",     "",     "", "Elec",     "",     "", "Hydr",     "",     ""],
				 *   ["Gear", "Berg", "Motr", "Swch", "Plug", "Cord", "Pump", "Leak", "Seal"],
				 * ]
				 */
				const revLabelGroups = firstDataLabels.slice().reverse()

				// A: Map each label cell to its shared-string index (blank=0, series names=1..TOT_SER, labels follow)
				let nextStrIdx = TOT_SER + 1
				const arrLabelStrIdx = revLabelGroups.map(labelsGroup => {
					const grpIdx = new Map<number, number>()
					labelsGroup.forEach((label, idx) => {
						if (label) grpIdx.set(idx, nextStrIdx++)
					})
					return grpIdx
				})

				// B: Collect vertical merges - an outer label spans until the next non-blank label in its group
				revLabelGroups.forEach((labelsGroup, idy) => {
					if (idy === TOT_LVL - 1) return // innermost level: one label per row, never merged
					for (let idx = 0; idx < TOT_CAT; idx++) {
						if (!labelsGroup[idx]) continue
						let idxEnd = idx + 1
						while (idxEnd < TOT_CAT && !labelsGroup[idxEnd]) idxEnd++
						if (idxEnd - idx > 1) {
							const colLtr = getExcelColName(idy + 1)
							arrMergeRefs.push(`${colLtr}${idx + 2}:${colLtr}${idxEnd + 1}`)
						}
						idx = idxEnd - 1
					}
				})

				// C: Iterate across labels/cats as these are the <row>'s
				for (let idx = 0; idx < TOT_CAT; idx++) {
					strSheetXml += `<row r="${idx + 2}" spans="1:${barSheetCols}">`

					// C-1: add a col for each label level (blank labels are covered by the merge above)
					arrLabelStrIdx.forEach((grpIdx, idy) => {
						const strIdx = grpIdx.get(idx)
						if (strIdx !== undefined) strSheetXml += `<c r="${getExcelColName(idy + 1)}${idx + 2}" t="s"><v>${strIdx}</v></c>`
					})

					// C-2: add a col for each data value
					for (let idy = 0; idy < TOT_SER; idy++) {
						// NOTE: Preserve 0 (a valid value); `|| ''` would blank it out in the embedded worksheet (issue #1430)
						const cellVal = (data[idy].values ?? [])[idx]
						const style = dataStyleIds[idy] ? ` s="${dataStyleIds[idy]}"` : ''
						strSheetXml += `<c r="${getExcelColName(TOT_LVL + idy + 1)}${idx + 2}"${style}><v>${typeof cellVal === 'number' ? cellVal : ''}</v></c>`
					}

					// C-3: packed errorrate columns after all series values
					{
						let errPacked = 0
						for (let idy = 0; idy < TOT_SER; idy++) {
							if (!seriesHasErrorrate(data[idy])) continue
							const errVal = (data[idy].errorrate ?? [])[idx]
							const errCol = TOT_LVL + TOT_SER + 1 + errPacked
							strSheetXml += `<c r="${getExcelColName(errCol)}${idx + 2}"><v>${typeof errVal === 'number' ? errVal : ''}</v></c>`
							errPacked++
						}
					}

					strSheetXml += '</row>'
				}
			}
		}
		strSheetXml += '</sheetData>'

		if (arrMergeRefs.length > 0) {
			strSheetXml += `<mergeCells count="${arrMergeRefs.length}">`
			arrMergeRefs.forEach(ref => (strSheetXml += `<mergeCell ref="${ref}"/>`))
			strSheetXml += '</mergeCells>'
		}

		strSheetXml += '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
		// Link the `table1.xml` file to define an actual Table in Excel
		// NOTE: This only works with scatter charts - all others give a "cannot find linked file" error
		// ....: Since we dont need the table anyway (chart data can be edited/range selected, etc.), just dont use this
		// ....: Leaving this so nobody foolishly attempts to add this in the future
		// strSheetXml += '<tableParts count="1"><tablePart r:id="rId1"/></tableParts>'
		strSheetXml += '</worksheet>\n'
		zipExcel.file('xl/worksheets/sheet1.xml', strSheetXml)
	}
}

/**
 * Finish the nested XLSX archive, add it to the presentation, and write the matching chart relationship and chart XML.
 * The generated chart XML references the relationship as `rId1`, so both parts must be emitted together.
 */
function addWorkbookToPresentation (chartObject: ISlideRelChart, zipExcel: JSZip, zip: JSZip): Promise<string> {
	return new Promise((resolve, reject) => {
		zipExcel
			.generateAsync({ type: 'base64' })
			.then(content => {
				// 1: Create the embedded Excel worksheet with labels and data
				zip.file(`ppt/embeddings/Microsoft_Excel_Worksheet${chartObject.globalId}.xlsx`, content, { base64: true })

				// 2: Create the chart.xml and rel files
				// PowerPoint finds the style and colour-style parts by relationship type, so they are
				// declared here rather than referenced from inside `chartN.xml`.
				// Classic charts write them only when requested; ChartEx layouts always need them.
				const isChartex = isChartexType(chartObject.opts._type)
				const writeStyleParts = wantsChartStyleParts(chartObject.opts)
				const styleRels = writeStyleParts
					? `<Relationship Id="rId2" Type="${CHART_STYLE.colorsRelType}" Target="colors${chartObject.globalId}.xml"/>` +
					`<Relationship Id="rId3" Type="${CHART_STYLE.styleRelType}" Target="style${chartObject.globalId}.xml"/>`
					: ''
				zip.file(
					'ppt/charts/_rels/' + chartObject.fileName + '.rels',
					'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
					'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
					`<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/package" Target="../embeddings/Microsoft_Excel_Worksheet${chartObject.globalId}.xlsx"/>` +
					styleRels +
					'</Relationships>'
				)
				zip.file(`ppt/charts/${chartObject.fileName}`, isChartex ? makeXmlChartEx(chartObject) : makeXmlCharts(chartObject))
				if (writeStyleParts) {
					zip.file(chartColorsPartName(chartObject.globalId), makeXmlChartColors(chartObject.opts?.chartColorStyle))
					zip.file(chartStylePartName(chartObject.globalId), makeXmlChartStyle(chartObject.opts?.chartStyle))
				}

				// 3: Done
				resolve('')
			})
			.catch(strErr => {
				reject(strErr)
			})
	})
}
