/**
 * Semantic contracts for the core slide and chart paths.
 *
 * Unlike golden XML snapshots, these checks document the OOXML that matters and allow harmless
 * serializer changes without regenerating fixture files.
 */
import { test, before } from 'node:test'
import assert from 'node:assert/strict'
import { readFileSync } from 'node:fs'
import { JSZip } from '@node-projects/jszip'
import pptxgen from '../src/pptxgen'
import { assertEmbeddedFontContracts, assertEmbeddedXlsxContracts, assertModernCommentPartContracts, assertNoEmbeddedFonts, assertPptxPackageContracts, readPart } from './pptx-contracts'

let zip: JSZip

before(async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addText('Contract', { x: 0.5, y: 0.3, w: 6, h: 0.5, fontSize: 18, color: '0000FF', bold: true })
	slide.addShape(pptx.ShapeType.rect, { x: 1, y: 1.2, w: 2, h: 1, fill: { color: 'FF0000' } })
	slide.addTable([['A', 'B'], ['1', '2']], { x: 0.5, y: 2.6, w: 5 })
	slide.addChart(pptx.ChartType.bar, [{ name: 'Sales', labels: ['Q1', 'Q2'], values: [10, 20] }], { x: 0.5, y: 4, w: 6, h: 3 })
	zip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
})

test('contract: package parts and relationships are coherent', async () => {
	await assertPptxPackageContracts(zip)
	await assertEmbeddedXlsxContracts(zip)
})

test('contract: every slide part is a PresentationML slide Override', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addText('one', { x: 0.5, y: 0.5, w: 4, h: 1 })
	pptx.addSlide().addText('two', { x: 0.5, y: 0.5, w: 4, h: 1 })
	pptx.addSlide().addText('three', { x: 0.5, y: 0.5, w: 4, h: 1 })
	const multi = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(multi)
})

test('contract: default export embeds no fonts', async () => {
	await assertNoEmbeddedFonts(zip)
})

test('contract: addFont embeds font parts, content types, and rels', async () => {
	const buf = readFileSync(new URL('./fonts/IBMPlexSans-Regular.ttf', import.meta.url))
	const fontFile = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength)
	const pptx = new pptxgen()
	await pptx.addFont({ fontFace: 'IBM Plex Sans', fontFile, fontType: 'ttf' })
	pptx.addSlide().addText('Hello', { x: 0.5, y: 0.5, w: 3, h: 1, fontFace: 'IBM Plex Sans', fontSize: 24 })
	const fontZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertEmbeddedFontContracts(fontZip, 'IBM Plex Sans')
})

test('contract: rejects a part without a declared content type', async () => {
	const invalidZip = await JSZip.loadAsync(await zip.generateAsync({ type: 'nodebuffer' }))
	invalidZip.file('ppt/undeclared.bin', 'invalid')
	await assert.rejects(assertPptxPackageContracts(invalidZip), /package part has no content type/)
})

test('contract: validates relationship references with any legal ID', async () => {
	const invalidZip = await JSZip.loadAsync(await zip.generateAsync({ type: 'nodebuffer' }))
	const slideXml = await readPart(invalidZip, 'ppt/slides/slide1.xml')
	const referencePattern = /r:(id|embed|link)="rId\d+"/
	assert.match(slideXml, referencePattern, 'test presentation has no relationship reference')
	invalidZip.file('ppt/slides/slide1.xml', slideXml.replace(referencePattern, (_match, attribute) => `r:${attribute}="custom-id"`))
	await assert.rejects(assertPptxPackageContracts(invalidZip), /missing custom-id relationship/)
})

test('contract: slide keeps text, shape, and table semantics', async () => {
	const xml = await readPart(zip, 'ppt/slides/slide1.xml')
	assert.match(xml, /<a:t>Contract<\/a:t>/, 'text content missing')
	assert.match(xml, /<a:prstGeom prst="rect">/, 'rectangle shape missing')
	assert.match(xml, /<a:srgbClr val="FF0000"\/>/, 'shape fill missing')
	assert.match(xml, /<a:tbl>/, 'table missing')
	assert.equal([...xml.matchAll(/<a:gridCol /g)].length, 2, 'table grid width changed')
})

test('contract: default tableStyles.xml stays an empty list with the built-in def GUID', async () => {
	const xml = await readPart(zip, 'ppt/tableStyles.xml')
	assert.ok(xml.trimEnd().endsWith('def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"/>'), `default tableStyles.xml changed: ${xml}`)
	assert.ok(!xml.includes('<a:tblStyle '), 'a default deck gained a table style')
})

test('contract: custom tableStyles write a:tblStyle and keep the built-in def GUID', async () => {
	const STYLE_ID = '{A1B2C3D4-1111-2222-3333-444455556666}'
	const pptx = new pptxgen()
	pptx.tableStyles = [{ id: STYLE_ID, name: 'Contract Blue', firstRow: { bold: true, fill: { color: '4472C4' } } }]
	pptx.addSlide().addTable([['H'], ['a']], { x: 1, y: 1, w: 4, tableStyleId: STYLE_ID, firstRow: true })
	const stylesZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	const xml = await readPart(stylesZip, 'ppt/tableStyles.xml')
	assert.ok(xml.includes(`styleId="${STYLE_ID}"`) && xml.includes('styleName="Contract Blue"'), 'custom tblStyle missing')
	assert.ok(xml.includes('def="{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}"'), '@def was repointed at a custom style')
	assert.ok(xml.includes('<a:firstRow>') && xml.includes('b="on"'), 'firstRow text style missing')
	const slide = await readPart(stylesZip, 'ppt/slides/slide1.xml')
	assert.ok(slide.includes(`<a:tableStyleId>${STYLE_ID}</a:tableStyleId>`), 'table did not reference the custom style')
})

test('contract: library version is generated from package.json', () => {
	const pptx = new pptxgen()
	assert.match(pptx.version, /^\d+\.\d+\.\d+/)
	assert.doesNotMatch(readFileSync(new URL('../src/pptxgen.ts', import.meta.url), 'utf8'), /const VERSION = '/)
})

test('contract: default slides omit classification and design nvPr extLst URIs', async () => {
	const xml = await readPart(zip, 'ppt/slides/slide1.xml')
	const presXml = await readPart(zip, 'ppt/presentation.xml')
	assert.doesNotMatch(xml, /\{1162E1C5-73C7-4A58-AE30-91384D911F3F\}/, 'classification URI must stay opt-in')
	assert.doesNotMatch(xml, /<p184:classification/, 'classification element must stay opt-in')
	assert.doesNotMatch(xml, /\{386F3935-93C4-4BCD-93E2-E3B085C9AB24\}/, 'designElem URI must stay opt-in')
	assert.doesNotMatch(xml, /\{E7BDC344-281C-4309-B0C6-D0EE65EED2A8\}/, 'designPr URI must stay opt-in')
	assert.doesNotMatch(presXml, /\{E3EDB536-0D56-4F60-86BA-61A60CA02DAB\}/, 'sldId designTagLst URI must stay opt-in')
	assert.match(presXml, /<p:sldId id="\d+" r:id="rId\d+"\/>/, 'default sldId must stay a self-closing element')
})

test('contract: bar chart keeps its data and chart type', async () => {
	const xml = await readPart(zip, 'ppt/charts/chart1.xml')
	assert.match(xml, /<c:barChart>/, 'bar chart missing')
	assert.match(xml, /<c:v>Sales<\/c:v>/, 'series name missing')
	assert.match(xml, /<c:v>Q1<\/c:v>/, 'category label missing')
	assert.match(xml, /<c:v>20<\/c:v>/, 'series value missing')
	assert.match(xml, /<c:cat>\s*<c:strRef>/, 'single-level categories must use c:strRef')
	assert.doesNotMatch(xml, /<c:multiLvlStrRef>/, 'single-level categories must not use multiLvlStrRef')
})

test('contract: slide chart relationship Target is relative; Content_Types PartName stays absolute', async () => {
	const rels = await readPart(zip, 'ppt/slides/_rels/slide1.xml.rels')
	assert.match(
		rels,
		/Type="http:\/\/schemas\.openxmlformats\.org\/officeDocument\/2006\/relationships\/chart" Target="\.\.\/charts\/chart\d+\.xml"/,
		'chart relationship Target must be relative to the slide part',
	)
	assert.doesNotMatch(rels, /Target="\/ppt\/charts\//, 'absolute chart relationship Target is non-idiomatic')
	const types = await readPart(zip, '[Content_Types].xml')
	assert.match(types, /PartName="\/ppt\/charts\/chart\d+\.xml"/, 'Content_Types Override PartName must stay package-absolute')
})

test('contract: gradient stops and color transforms keep DrawingML fill semantics', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addShape(pptx.ShapeType.rect, {
		x: 0.5, y: 0.5, w: 2, h: 1,
		fill: {
			type: 'gradient',
			gradient: {
				type: 'linear',
				angle: 90,
				stops: [
					{ color: { baseColor: 'accent1', tint: 40 }, pos: 0 },
					{ color: { baseColor: 'accent2', shade: 25, lumMod: 80 }, pos: 100 },
				],
			},
		},
	})
	slide.addShape(pptx.ShapeType.rect, {
		x: 3, y: 0.5, w: 2, h: 1,
		fill: { type: 'solid', color: '00FF00' },
	})
	slide.addShape(pptx.ShapeType.rect, {
		x: 5.5, y: 0.5, w: 2, h: 1,
		fill: { type: 'pattern', pattern: { prst: 'ltHorz', color: '000000', bgColor: 'FFFFFF' } },
	})
	slide.addShape(pptx.ShapeType.rect, {
		x: 0.5, y: 2, w: 2, h: 1,
		fill: { type: 'solid', color: { baseColor: 'bg1', shade: 10 } },
	})

	const fillsZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(fillsZip)
	const xml = await readPart(fillsZip, 'ppt/slides/slide1.xml')

	assert.match(xml, /<a:gradFill[^>]*>/, 'gradFill missing')
	const gsLst = xml.match(/<a:gsLst>([\s\S]*?)<\/a:gsLst>/)
	assert.ok(gsLst, 'gsLst missing')
	const stops = [...gsLst[1].matchAll(/<a:gs pos="(\d+)">/g)]
	assert.ok(stops.length >= 2, 'CT_GradientStopList requires at least 2 gs')
	for (const [, pos] of stops) {
		const n = Number(pos)
		assert.ok(n >= 0 && n <= 100000, `gs@pos ${pos} outside ST_PositiveFixedPercentage`)
	}
	assert.match(
		xml,
		/<a:gs pos="0"><a:schemeClr val="accent1"><a:tint val="40000"\/>/,
		'tint transform must be a child of schemeClr, not of gs'
	)
	assert.match(
		xml,
		/<a:schemeClr val="accent2"><a:lumMod val="80000"\/><a:shade val="25000"\/>/,
		'shade/lumMod transforms missing on gradient stop'
	)
	assert.match(xml, /<a:lin ang="5400000"/, 'linear shade angle missing')

	assert.match(xml, /<a:solidFill><a:srgbClr val="00FF00"\/>/, 'solid fill regressing')
	assert.match(xml, /<a:pattFill prst="ltHorz">/, 'pattern fill regressing')
	assert.match(xml, /<a:solidFill><a:schemeClr val="bg1"><a:shade val="10000"\/>/, 'modified-theme fill regressing')
})

const CHART_TRACKING_REF_URI = '{FD5EFAAD-0ECE-453E-9831-46B23BE46B34}'

function addSampleChart (pptx: pptxgen) {
	pptx.addSlide().addChart(pptx.ChartType.bar, [{ name: 'Sales', labels: ['Q1', 'Q2'], values: [10, 20] }], { x: 0.5, y: 4, w: 6, h: 3 })
}

async function writePptx (pptx: pptxgen): Promise<JSZip> {
	return await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
}

async function readChartXml (pkg: JSZip): Promise<string> {
	const file = pkg.file(/ppt\/charts\/chart\d+\.xml$/)[0]
	assert.ok(file, 'missing chart part')
	return await file.async('string')
}

test('contract: unset chartTrackingRefBased does not emit the MS-PPTX tracking URI', async () => {
	const presPrXml = await readPart(zip, 'ppt/presProps.xml')
	assert.doesNotMatch(presPrXml, /FD5EFAAD-0ECE-453E-9831-46B23BE46B34/, 'chartTrackingRefBased URI must be absent when unset')
	assert.doesNotMatch(presPrXml, /chartTrackingRefBased/, 'chartTrackingRefBased element must be absent when unset')
})

test('contract: chartTrackingRefBased emits the MS-PPTX presentationPr URI', async () => {
	const unset = new pptxgen()
	addSampleChart(unset)
	const optIn = new pptxgen()
	optIn.chartTrackingRefBased = true
	addSampleChart(optIn)

	const [unsetZip, optInZip] = await Promise.all([writePptx(unset), writePptx(optIn)])
	const unsetPresPr = await readPart(unsetZip, 'ppt/presProps.xml')
	const optInPresPr = await readPart(optInZip, 'ppt/presProps.xml')

	assert.doesNotMatch(unsetPresPr, /FD5EFAAD-0ECE-453E-9831-46B23BE46B34/, 'unset deck must not emit the tracking URI')
	assert.match(optInPresPr, new RegExp(`uri="${CHART_TRACKING_REF_URI.replace(/[{}]/g, '\\$&')}"`), `missing MS-PPTX ┬º2.2.12 URI ${CHART_TRACKING_REF_URI}: ${optInPresPr}`)
	assert.match(optInPresPr, /<p15:chartTrackingRefBased\b[^>]*\bval="1"/, `chartTrackingRefBased val missing: ${optInPresPr}`)
	assert.match(optInPresPr, /xmlns:p15="http:\/\/schemas\.microsoft\.com\/office\/powerpoint\/2012\/main"/, 'p15 2012/main namespace missing')

	const [unsetChart, optInChart] = await Promise.all([readChartXml(unsetZip), readChartXml(optInZip)])
	assert.equal(optInChart, unsetChart, 'chart XML must not change when chartTrackingRefBased is set')
})

test('contract: modern comment parts satisfy MS-PPTX ┬º2.1.5ΓÇô2.1.6', async () => {
	const pptx = new pptxgen()
	pptx.commentAuthors = [{ name: 'Ada Lovelace', initials: 'AL' }]
	pptx.addSlide().addComment({ text: 'Review this', author: 0, x: 1, y: 1, replies: [{ text: 'Done', author: 'Grace' }] })
	const commentsZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(commentsZip)
	await assertModernCommentPartContracts(commentsZip)
})

test('contract: rejects a comments part without the ┬º2.1.5 content type', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addComment({ text: 'x', author: 'Ada' })
	const invalidZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	const ctXml = await readPart(invalidZip, '[Content_Types].xml')
	invalidZip.file('[Content_Types].xml', ctXml.replace('application/vnd.ms-powerpoint.comments+xml', 'application/xml'))
	await assert.rejects(assertModernCommentPartContracts(invalidZip), /Comment part content type/)
})

test('contract: rejects an authors relationship that is not Internal', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addComment({ text: 'x', author: 'Ada' })
	const invalidZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	const relsXml = await readPart(invalidZip, 'ppt/_rels/presentation.xml.rels')
	invalidZip.file(
		'ppt/_rels/presentation.xml.rels',
		relsXml.replace(
			'Type="http://schemas.microsoft.com/office/2018/10/relationships/authors"',
			'Type="http://schemas.microsoft.com/office/2018/10/relationships/authors" TargetMode="External"'
		)
	)
	await assert.rejects(assertModernCommentPartContracts(invalidZip), /Author relationship TargetMode MUST be Internal/)
})

test('contract: rejects a comments part without a slide relationship', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addComment({ text: 'x', author: 'Ada' })
	const invalidZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	const slideRels = await readPart(invalidZip, 'ppt/slides/_rels/slide1.xml.rels')
	invalidZip.file(
		'ppt/slides/_rels/slide1.xml.rels',
		slideRels.replace(/<Relationship[^>]*relationships\/comments[^>]*\/>/, '')
	)
	await assert.rejects(assertModernCommentPartContracts(invalidZip), /MUST be the target of an explicit relationship/)
})

// ChartEx (PowerPoint 2016+) charts - MS-ODRAWXML 2.1
// The package contract is what makes or breaks these: a wrong content type, relationship type, or a
// missing sidecar part makes PowerPoint declare the file damaged rather than degrade gracefully.

/** Chart part ids are global to the process, so the part name is discovered rather than assumed */
async function buildChartEx (type: string, data: object[], opts: object = {}): Promise<{ zip: JSZip, part: string }> {
	const pptx = new pptxgen()
	pptx.addSlide().addChart(type as never, data as never, { x: 0.5, y: 0.5, w: 6, h: 4, ...opts })
	const zip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	const part = Object.keys(zip.files).find(name => /^ppt\/charts\/chartEx\d+\.xml$/.test(name)) ?? ''
	return { zip, part }
}

/** The chartex part XML for one chart, asserting the part exists under a chartEx name */
async function chartExXml (type: string, data: object[], opts: object = {}): Promise<string> {
	const { zip, part } = await buildChartEx(type, data, opts)
	assert.ok(part, `${type}: chartex parts are named chartExN.xml, not chartN.xml`)
	return await readPart(zip, part)
}

const WATERFALL_DATA = [{ name: 'Cash', labels: ['Start', 'Q1', 'Q2', 'End'], values: [100, 30, -20, 110] }]

test('contract: a chartex chart is a separate part with its own content and relationship types', async () => {
	const { zip: cxZip, part } = await buildChartEx('waterfall', WATERFALL_DATA)
	await assertPptxPackageContracts(cxZip)
	await assertEmbeddedXlsxContracts(cxZip)

	assert.ok(part, 'chartex parts are named chartExN.xml, not chartN.xml')
	assert.equal(Object.keys(cxZip.files).filter(name => /^ppt\/charts\/chart\d+\.xml$/.test(name)).length, 0, 'a chartex chart must not also write an ECMA-376 chart part')

	const contentTypes = await readPart(cxZip, '[Content_Types].xml')
	assert.match(contentTypes, new RegExp(`PartName="/${part}" ContentType="application/vnd\\.ms-office\\.chartex\\+xml"`), 'chartex content type missing')

	const rels = await readPart(cxZip, 'ppt/slides/_rels/slide1.xml.rels')
	assert.match(rels, /Type="http:\/\/schemas\.microsoft\.com\/office\/2014\/relationships\/chartEx" Target="\.\.\/charts\/chartEx\d+\.xml"/, 'chartex relationship type missing')
	assert.doesNotMatch(rels, /relationships\/chart"/, 'a chartex part must not use the ECMA-376 chart relationship type')
})

test('contract: a chartex chart relates to the style and color parts PowerPoint requires', async () => {
	const { zip: cxZip, part } = await buildChartEx('treemap', [{ name: 'Revenue', labels: ['A', 'B', 'C'], values: [10, 20, 30] }])

	const chartRels = await readPart(cxZip, `ppt/charts/_rels/${part.split('/').pop() ?? ''}.rels`)
	const styleRel = /Type="http:\/\/schemas\.microsoft\.com\/office\/2011\/relationships\/chartStyle" Target="(style\d+\.xml)"/.exec(chartRels)
	const colorsRel = /Type="http:\/\/schemas\.microsoft\.com\/office\/2011\/relationships\/chartColorStyle" Target="(colors\d+\.xml)"/.exec(chartRels)
	assert.ok(styleRel, 'chart style relationship missing')
	assert.ok(colorsRel, 'chart color style relationship missing')
	assert.match(chartRels, /Id="rId1" Type="[^"]*relationships\/package"/, 'the embedded workbook must stay rId1')

	const contentTypes = await readPart(cxZip, '[Content_Types].xml')
	assert.match(contentTypes, new RegExp(`/ppt/charts/${styleRel[1]}" ContentType="application/vnd\\.ms-office\\.chartstyle\\+xml"`), 'chart style content type missing')
	assert.match(contentTypes, new RegExp(`/ppt/charts/${colorsRel[1]}" ContentType="application/vnd\\.ms-office\\.chartcolorstyle\\+xml"`), 'chart color style content type missing')
	assert.match(await readPart(cxZip, `ppt/charts/${styleRel[1]}`), /^<\?xml[^>]*\?><cs:chartStyle /, 'chart style part is not a cs:chartStyle document')
	assert.match(await readPart(cxZip, `ppt/charts/${colorsRel[1]}`), /^<\?xml[^>]*\?><cs:colorStyle /, 'chart color style part is not a cs:colorStyle document')
})

test('contract: a chartex frame is offered through mc:AlternateContent with a fallback', async () => {
	const { zip: cxZip } = await buildChartEx('waterfall', WATERFALL_DATA)
	const slideXml = await readPart(cxZip, 'ppt/slides/slide1.xml')

	const block = slideXml.slice(slideXml.indexOf('<mc:AlternateContent'), slideXml.indexOf('</mc:AlternateContent>'))
	assert.ok(block, 'a chartex frame must be wrapped in mc:AlternateContent')
	assert.match(block, /<mc:Choice xmlns:cx1="http:\/\/schemas\.microsoft\.com\/office\/drawing\/2015\/9\/8\/chartex" Requires="cx1">/, 'chartex Choice condition missing')
	assert.match(block, /<a:graphicData uri="http:\/\/schemas\.microsoft\.com\/office\/drawing\/2014\/chartex">\s*<cx:chart [^>]*r:id="rId\d+"\/>/, 'chartex graphicData uri or chart reference missing')
	assert.match(block, /<mc:Fallback><p:sp>/, 'a chartex frame must offer a fallback shape')
	assert.match(block, /<a:spLocks noTextEdit="1"\/>/, 'the fallback must not be editable into something that no longer matches the chart')

	const funnelXml = await readPart((await buildChartEx('funnel', [{ name: 'Pipeline', labels: ['Leads', 'Won'], values: [500, 30] }])).zip, 'ppt/slides/slide1.xml')
	assert.match(funnelXml, /xmlns:cx1="http:\/\/schemas\.microsoft\.com\/office\/drawing\/2015\/10\/21\/chartex"/, 'funnel must require its own chartex generation')
})

test('contract: each chartex layout emits the markup PowerPoint keys off', async () => {
	const waterfall = await chartExXml('waterfall', WATERFALL_DATA, { chartExSubtotals: [0, 3] })
	assert.match(waterfall, /<cx:series layoutId="waterfall" uniqueId="\{[0-9A-F-]{36}\}">/, 'waterfall layoutId or uniqueId missing')
	assert.match(waterfall, /<cx:subtotals><cx:idx val="0"\/><cx:idx val="3"\/><\/cx:subtotals>/, 'waterfall subtotals missing')
	assert.match(waterfall, /<cx:strDim type="cat"><cx:f>Sheet1!\$A\$2:\$A\$5<\/cx:f>/, 'category dimension does not point at the worksheet label column')
	assert.match(waterfall, /<cx:numDim type="val"><cx:f>Sheet1!\$B\$2:\$B\$5<\/cx:f>/, 'value dimension does not point at the worksheet series column')
	assert.match(waterfall, /<cx:axis id="0"><cx:catScaling gapWidth="0\.5"\/>.*<cx:axis id="1"><cx:valScaling\/>/s, 'waterfall axes missing')

	const histogram = await chartExXml('histogram', [{ name: 'Ages', labels: ['a', 'b', 'c'], values: [3, 7, 12] }], { chartExBinCount: 4 })
	assert.match(histogram, /<cx:series layoutId="clusteredColumn"/, 'histogram layoutId missing')
	assert.match(histogram, /<cx:binning intervalClosed="r"><cx:binCount val="4"\/><\/cx:binning>/, 'histogram bin count missing')
	assert.doesNotMatch(histogram, /<cx:strDim type="cat"/, 'a histogram bins raw values and has no category dimension')

	const treemap = await chartExXml('treemap', [{ name: 'Revenue', labels: ['A', 'B'], values: [10, 20] }], { chartExParentLabels: 'banner' })
	assert.match(treemap, /<cx:numDim type="size">/, 'treemap must size segments through a size dimension')
	assert.match(treemap, /<cx:parentLabelLayout val="banner"\/>/, 'treemap parent label layout missing')
	assert.doesNotMatch(treemap, /<cx:axis /, 'a treemap has no axes')

	const sunburst = await chartExXml('sunburst', [{ name: 'Revenue', labels: ['A', 'B'], values: [10, 20] }])
	assert.match(sunburst, /<cx:series layoutId="sunburst"/, 'sunburst layoutId missing')

	const box = await chartExXml('boxWhisker', [
		{ name: 'A', labels: ['x', 'y'], values: [1, 5] },
		{ name: 'B', labels: ['x', 'y'], values: [2, 6] },
	], { chartExMeanLine: true })
	assert.match(box, /<cx:data id="0">.*<cx:data id="1">/s, 'box & whisker needs one data block per series')
	assert.match(box, /<cx:numDim type="val"><cx:f>Sheet1!\$C\$2:\$C\$3<\/cx:f>/, 'the second series must read the second worksheet column')
	assert.match(box, /<cx:dataId val="1"\/>/, 'the second series must name its own data block')
	assert.match(box, /<cx:visibility meanLine="1" meanMarker="1" nonoutliers="0" outliers="1"\/>/, 'box & whisker visibility missing')
	assert.match(box, /<cx:statistics quartileMethod="exclusive"\/>/, 'box & whisker quartile method missing')

	const funnel = await chartExXml('funnel', [{ name: 'Pipeline', labels: ['Leads', 'Won'], values: [500, 30] }])
	assert.match(funnel, /<cx:axis id="1"><cx:catScaling/, 'funnel category axis missing')
	assert.doesNotMatch(funnel, /<cx:axis id="0"/, 'a funnel has no value axis')
})

test('contract: chartex options that PowerPoint would reject are dropped with a warning', async () => {
	const warnings: string[] = []
	const origWarn = console.warn
	let xml = ''
	try {
		console.warn = (msg: string) => warnings.push(String(msg))
		xml = await chartExXml('waterfall', WATERFALL_DATA, { chartExSubtotals: [0, 9], chartExBinCount: 0, chartExParentLabels: 'sideways' })
	} finally {
		console.warn = origWarn
	}

	assert.equal(warnings.length, 3, `expected one warning per rejected option, got: ${warnings.join(' | ')}`)
	assert.match(xml, /<cx:subtotals><cx:idx val="0"\/><\/cx:subtotals>/, 'the valid subtotal index must survive')
	assert.doesNotMatch(xml, /val="9"/, 'an out-of-range subtotal index must not be emitted')
})

test('contract: a chartex type cannot be smuggled into a multi-type chart', () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	assert.throws(
		() => slide.addChart([
			{ type: 'bar' as never, data: [{ name: 'A', labels: ['x'], values: [1] }], options: {} },
			{ type: 'waterfall' as never, data: [{ name: 'B', labels: ['x'], values: [2] }], options: {} },
		], []),
		/waterfall.*multi-type/,
		'a chartex layout owns the whole plot area and cannot share one'
	)
})

test('contract: classic charts are untouched by the chartex path', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addChart(pptx.ChartType.bar, [{ name: 'Sales', labels: ['Q1', 'Q2'], values: [10, 20] }], { x: 0.5, y: 0.5, w: 6, h: 4 })
	const barZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	const slideXml = await readPart(barZip, 'ppt/slides/slide1.xml')

	assert.doesNotMatch(slideXml, /mc:AlternateContent/, 'an ECMA-376 chart needs no compatibility wrapper')
	assert.match(slideXml, /<a:graphicData uri="http:\/\/schemas\.openxmlformats\.org\/drawingml\/2006\/chart">/, 'classic chart graphicData uri changed')
	const chartParts = Object.keys(barZip.files).filter(name => /^ppt\/charts\/chart\d+\.xml$/.test(name))
	assert.equal(chartParts.length, 1, 'classic chart part missing')
	assert.equal(Object.keys(barZip.files).filter(name => /chartEx/.test(name)).length, 0, 'no chartex part may appear without a chartex chart')
	assert.doesNotMatch(await readPart(barZip, '[Content_Types].xml'), /ms-office\.chartex/, 'no chartex content type may be declared without a chartex chart')
	assert.equal(Object.keys(barZip.files).filter(name => /\/(colors|style)\d+\.xml$/.test(name)).length, 0, 'default classic charts must not write style or colors parts')
})

test('contract: chartex data that PowerPoint would reject is normalized, not emitted', async () => {
	const warnings: string[] = []
	const origWarn = console.warn
	let ragged = ''
	let multi = ''
	let extra = ''
	let noLabels = ''
	try {
		console.warn = (msg: string) => warnings.push(String(msg))
		ragged = await chartExXml('waterfall', [{ name: 'Cash', labels: ['a', 'b', 'c', 'd'], values: [10, null, 30] }])
		multi = await chartExXml('treemap', [{ name: 'R', labels: [['Gear', 'Berg', 'Motr'], ['Mech', '', '']], values: [1, 2, 3] }])
		extra = await chartExXml('waterfall', [{ name: 'A', labels: ['a', 'b'], values: [1, 2] }, { name: 'B', labels: ['a', 'b'], values: [3, 4] }])
		noLabels = await chartExXml('histogram', [{ name: 'Ages', values: [1, 2, 3] }])
	} finally {
		console.warn = origWarn
	}

	assert.doesNotMatch(ragged, />null<|>undefined</, 'a missing point must be blank, not the string "null"')
	assert.match(ragged, /<cx:pt idx="1"><\/cx:pt>/, 'a missing point must still hold its index')
	assert.match(ragged, /<cx:f>Sheet1!\$B\$2:\$B\$4<\/cx:f><cx:lvl ptCount="3"/, 'the value range must be as long as its ptCount')
	assert.match(ragged, /<cx:f>Sheet1!\$A\$2:\$A\$5<\/cx:f><cx:lvl ptCount="4"/, 'the category range must be as long as its ptCount')

	assert.match(multi, /<cx:f>Sheet1!\$A\$2:\$A\$4<\/cx:f><cx:lvl ptCount="3"><cx:pt idx="0">Gear/, 'the leaf label level must address column A')
	assert.match(multi, /type="size"><cx:f>Sheet1!\$B\$2:\$B\$4</, 'values must stay in column B after flattening')
	assert.doesNotMatch(multi, /Mech/, 'the extra label level must be dropped, not emitted')

	assert.equal((extra.match(/<cx:series /g) ?? []).length, 1, 'a waterfall plots one series')
	assert.match(noLabels, /<cx:numDim type="val"><cx:f>Sheet1!\$B\$2:\$B\$4</, 'a label-less histogram must still address its values')

	assert.equal(warnings.length, 2, `expected one warning per normalization, got: ${warnings.join(' | ')}`)

	const box = await chartExXml('boxWhisker', [{ name: 'A', labels: ['a', 'b'], values: [1, 2] }, { name: 'B', labels: ['a', 'b'], values: [3, 4] }])
	assert.equal((box.match(/<cx:series /g) ?? []).length, 2, 'box & whisker plots one series per distribution')
})

test('contract: chartex options use the cx vocabulary, not the ECMA-376 one', async () => {
	assert.match(
		await chartExXml('treemap', [{ name: 'R', labels: ['a'], values: [1] }], { showLegend: true, legendPos: 'tr' }),
		/<cx:legend pos="r" align="min" overlay="0"\/>/,
		'top-right must become a top-aligned right legend'
	)

	const warnings: string[] = []
	const origWarn = console.warn
	let honoured = ''
	let rejected = ''
	try {
		console.warn = (msg: string) => warnings.push(String(msg))
		honoured = await chartExXml('waterfall', [{ name: 'C', labels: ['a', 'b'], values: [1, 2] }], { showValue: true, dataLabelPosition: 'ctr' })
		rejected = await chartExXml('waterfall', [{ name: 'C', labels: ['a', 'b'], values: [1, 2] }], { showValue: true, dataLabelPosition: 'bestFit' })
	} finally {
		console.warn = origWarn
	}

	assert.match(honoured, /<cx:dataLabels pos="ctr">/, 'a valid chartex label position must survive')
	assert.equal(warnings.length, 1, `only the unsupported position may warn, got: ${warnings.join(' | ')}`)
	assert.match(rejected, /<cx:dataLabels pos="outEnd">/, 'an unsupported position falls back to the layout default')

	const coloured = await chartExXml('funnel', [{ name: 'P', labels: ['a', 'b'], values: [5, 3], color: 'FF0000' }], { chartColors: ['00FF00'] })
	assert.match(coloured, /<cx:spPr><a:solidFill><a:srgbClr val="FF0000"\/><\/a:solidFill><\/cx:spPr>/, 'an explicit series colour must be emitted')
	assert.doesNotMatch(coloured, /00FF00/, 'chartColors does not apply to a chartex layout')
})

test('contract: the chartex fallback shape mirrors the chart frame it replaces', async () => {
	const { zip } = await buildChartEx('waterfall', WATERFALL_DATA, { objectName: 'WF', title: 'Quarterly cash', showTitle: true })
	const slideXml = await readPart(zip, 'ppt/slides/slide1.xml')
	const fallback = slideXml.slice(slideXml.indexOf('<mc:Fallback>'), slideXml.indexOf('</mc:Fallback>'))

	assert.match(fallback, /name="WF"/, 'the fallback must carry the chart name')
	assert.doesNotMatch(fallback, /title="Quarterly cash"/, 'the chart title must not leak into the alt-text title')
	assert.match(fallback, /<a:spLocks noTextEdit="1"\/>/, 'the fallback must never be text-editable')
})

const HOVER_PNG = 'image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAAC0lEQVR42mNkAAIAAAoAAv/lxKUAAAAASUVORK5CYII='
const HOVER_WAV = 'audio/wav;base64,UklGRiQAAABXQVZFZm10IBAAAAABAAEAQB8AAEAfAAABAAgAZGF0YQAAAAA='

test('contract: text fields emit a:fld with a cached value', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addText([
		{ text: '22/08/2026', options: { field: 'datetime1' } },
		{ text: ' page ' },
		{ text: '1', options: { field: 'slidenum' } },
	], { x: 1, y: 1, w: 4, h: 1 })
	const fldZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(fldZip)
	const xml = await readPart(fldZip, 'ppt/slides/slide1.xml')

	assert.doesNotMatch(xml, /NaN|undefined/, 'field options must not leak invalid values')
	assert.match(xml, /<a:fld id="\{[0-9a-f]{8}-0000-0000-0000-[0-9a-f]{12}\}" type="datetime1"><a:rPr[^>]*>[\s\S]*?<a:t>22\/08\/2026<\/a:t><\/a:fld>/, 'datetime field missing')
	assert.match(xml, /<a:fld id="\{[0-9a-f-]+\}" type="slidenum">[\s\S]*?<a:t>1<\/a:t><\/a:fld>/, 'slide-number field missing')
	assert.match(xml, /<\/a:fld><a:r>[\s\S]*?<a:t> page <\/a:t><\/a:r><a:fld/, 'a field must coexist with plain runs')
	const ids = [...xml.matchAll(/<a:fld id="(\{[^"]+\})"/g)].map(match => match[1])
	assert.equal(new Set(ids).size, 2, 'each field needs its own id')

	const warnings: string[] = []
	const origWarn = console.warn
	console.warn = (msg: string) => warnings.push(String(msg))
	let badXml = ''
	try {
		const bad = new pptxgen()
		bad.addSlide().addText([{ text: 'x', options: { field: 'lunchtime' as unknown as 'slidenum' } }], { x: 1, y: 1, w: 2, h: 1 })
		badXml = await readPart(await JSZip.loadAsync((await bad.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')
	} finally {
		console.warn = origWarn
	}
	assert.ok(warnings.some(w => w.includes('unknown text field "lunchtime"')), 'unknown field type must warn')
	assert.doesNotMatch(badXml, /a:fld/, 'an unknown field type must not be emitted')
	assert.match(badXml, /<a:t>x<\/a:t>/, 'the text still renders as a plain run')
})

test('contract: bullets support colour, size, font, and pictures', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addText('picture bullet', { x: 1, y: 1, w: 4, h: 1, bullet: { image: HOVER_PNG, color: 'FF0000', size: 150 } })
	slide.addText('char bullet', { x: 1, y: 3, w: 4, h: 1, bullet: { characterCode: '25BA', fontFace: 'Wingdings', sizePts: 14 } })
	slide.addText('numbered', { x: 1, y: 5, w: 4, h: 1, bullet: { type: 'number', color: '0000FF' } })
	const buZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(buZip)
	const xml = await readPart(buZip, 'ppt/slides/slide1.xml')

	assert.match(xml, /<a:buClr><a:srgbClr val="FF0000"\/><\/a:buClr><a:buSzPct val="150000"\/><a:buBlip><a:blip r:embed="rId\d+"\/><\/a:buBlip>/, 'picture bullet missing')
	const rid = /<a:blip r:embed="(rId\d+)"\/><\/a:buBlip>/.exec(xml)?.[1]
	assert.ok(rid, 'picture bullet has no relationship')
	assert.match(await readPart(buZip, 'ppt/slides/_rels/slide1.xml.rels'), new RegExp(`<Relationship Id="${rid}" Type="[^"]*\\/image"`), 'bullet image relationship missing')
	assert.match(xml, /<a:buSzPts val="1400"\/><a:buFont typeface="Wingdings"\/><a:buChar char="&#x25BA;"\/>/, 'char bullet font/size missing')
	assert.match(xml, /<a:buClr><a:srgbClr val="0000FF"\/><\/a:buClr><a:buSzPct val="100000"\/><a:buFont typeface="\+mj-lt"\/><a:buAutoNum/, 'numbered bullet changed')
})

test('contract: rtlCol follows rtlMode, and kumimoji is settable', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addText('rtl', { x: 1, y: 1, w: 4, h: 1, columns: 2, rtlMode: true, kumimoji: true })
	slide.addText('ltr', { x: 1, y: 3, w: 4, h: 1, columns: 2 })
	slide.addText('override', { x: 1, y: 5, w: 4, h: 1, rtlMode: true, rtlColumns: false })
	const xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')

	assert.match(xml, /<a:bodyPr wrap="square" numCol="2" rtlCol="1"/, 'rtlCol must follow rtlMode')
	assert.match(xml, /<a:bodyPr wrap="square" numCol="2" rtlCol="0"/, 'a non-RTL box must stay rtlCol="0"')
	assert.match(xml, /kumimoji="1"/, 'kumimoji missing')
	assert.equal([...xml.matchAll(/rtlCol="0"/g)].length, 2, 'the explicit override must produce rtlCol="0"')
})

test('contract: text without fields, bullet extras, or RTL is unchanged', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addText('plain', { x: 1, y: 1, w: 3, h: 1, bullet: { characterCode: '2022' } })
	const xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')

	assert.doesNotMatch(xml, /a:fld|a:buClr|a:buBlip|a:buSzPts|kumimoji/, 'nothing new may appear unasked')
	assert.match(xml, /<a:buSzPct val="100000"\/><a:buChar char="&#x2022;"\/>/, 'default bullet output changed')
	assert.match(xml, /rtlCol="0"/, 'default rtlCol changed')
})

test('contract: mouse-over actions use the right element for each host', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addShape(pptx.ShapeType.rect, {
		x: 1, y: 1, w: 2, h: 1,
		hyperlink: { url: 'https://example.com', tooltip: 'go' },
		hyperlinkHover: { slide: 2, tooltip: 'peek' },
	})
	slide.addImage({ data: HOVER_PNG, x: 4, y: 1, w: 1, h: 1, hyperlinkHover: { url: 'https://hover.test' } })
	slide.addText([{ text: 'link', options: { hyperlink: { url: 'https://a.test' }, hyperlinkHover: { slide: 2 } } }], { x: 1, y: 3, w: 3, h: 1 })
	pptx.addSlide()

	const hlZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(hlZip)
	const xml = await readPart(hlZip, 'ppt/slides/slide1.xml')

	assert.doesNotMatch(xml, /NaN|undefined/, 'hyperlink options must not leak invalid values')
	assert.match(xml, /<a:hlinkClick [^>]*tooltip="go"[^>]*\/><a:hlinkHover r:id="rId\d+" action="ppaction:\/\/hlinksldjump" tooltip="peek"\/>/, 'shape hover must use a:hlinkHover, after the click link')
	assert.match(xml, /<a:hlinkHover r:id="rId\d+" invalidUrl="" action="" tgtFrame="" tooltip="" history="1"\/>/, 'image hover link missing')
	assert.match(xml, /<a:hlinkMouseOver r:id="rId\d+" action="ppaction:\/\/hlinksldjump" tooltip=""\/>/, 'text-run hover must use a:hlinkMouseOver')
	assert.equal([...xml.matchAll(/<a:hlinkMouseOver/g)].length, 1, 'exactly one run-level hover expected')
	assert.equal([...xml.matchAll(/<a:hlinkHover/g)].length, 2, 'exactly two shape-level hovers expected')
	assert.doesNotMatch(xml, /<p:cNvPr[^>]*>(?:(?!<\/p:cNvPr>)[\s\S])*<a:hlinkMouseOver/, 'a:hlinkMouseOver must not appear in p:cNvPr')

	const rels = await readPart(hlZip, 'ppt/slides/_rels/slide1.xml.rels')
	;[...xml.matchAll(/<a:hlink\w+ r:id="(rId\d+)"/g)].map(match => match[1]).forEach(rid => {
		assert.match(rels, new RegExp(`<Relationship Id="${rid}"`), `${rid} has no relationship`)
	})
})

test('contract: action sounds and click attributes are emitted', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addText([{
		text: 'noisy',
		options: { hyperlink: { url: 'https://a.test', highlightClick: true, stopSoundsOnClick: true, sound: { data: HOVER_WAV, name: 'ding.wav' } } },
	}], { x: 1, y: 1, w: 3, h: 1 })
	const sndZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(sndZip)
	const xml = await readPart(sndZip, 'ppt/slides/slide1.xml')

	assert.match(xml, /<a:hlinkClick [^>]*highlightClick="1" endSnd="1">/, 'click attributes must be settable')
	assert.match(xml, /<a:snd r:embed="rId\d+" name="ding\.wav"\/>/, 'action sound missing')
	assert.equal(Object.keys(sndZip.files).filter(file => /^ppt\/media\/.+\.wav$/.test(file)).length, 1, 'wav part missing')
	assert.match(await readPart(sndZip, '[Content_Types].xml'), /Extension="wav"/, 'wav content type missing')

	const plain = new pptxgen()
	plain.addSlide().addText([{ text: 'quiet', options: { hyperlink: { url: 'https://a.test' } } }], { x: 1, y: 1, w: 3, h: 1 })
	const plainXml = await readPart(await JSZip.loadAsync((await plain.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')
	assert.doesNotMatch(plainXml, /highlightClick|endSnd|a:snd|hlinkHover|hlinkMouseOver/, 'a plain link must not gain hover, sound, or click attributes')
})

test('contract: invalid hover and sound input is dropped with a warning', async () => {
	const warnings: string[] = []
	const origWarn = console.warn
	console.warn = (msg: string) => warnings.push(String(msg))
	let xml = ''
	try {
		const pptx = new pptxgen()
		const slide = pptx.addSlide()
		slide.addShape(pptx.ShapeType.rect, { x: 1, y: 1, w: 2, h: 1, hyperlinkHover: {} })
		slide.addShape(pptx.ShapeType.rect, { x: 4, y: 1, w: 2, h: 1, hyperlinkHover: { url: 'https://b.test', sound: {} } })
		slide.addShape(pptx.ShapeType.rect, { x: 1, y: 3, w: 2, h: 1, hyperlinkHover: { url: 'https://c.test', sound: { data: 'not-base64' } } })
		xml = await readPart(await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer), 'ppt/slides/slide1.xml')
	} finally {
		console.warn = origWarn
	}

	assert.ok(warnings.some(w => w.includes('hyperlink requires either `url` or `slide`')), 'a target-less link must warn')
	assert.ok(warnings.some(w => w.includes('`sound` requires `data` or `path`')), 'a sound without data must warn')
	assert.ok(warnings.some(w => w.includes('`sound.data` lacks a base64 header')), 'bad sound data must warn')

	assert.equal([...xml.matchAll(/<a:hlinkHover/g)].length, 2, 'links with a target survive; the target-less one does not')
	assert.doesNotMatch(xml, /<a:snd /, 'no invalid sound may be written')
	assert.doesNotMatch(xml, /r:id="rId0"/, 'no link may reference a non-existent relationship')
})
