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
	assert.match(optInPresPr, new RegExp(`uri="${CHART_TRACKING_REF_URI.replace(/[{}]/g, '\\$&')}"`), `missing MS-PPTX §2.2.12 URI ${CHART_TRACKING_REF_URI}: ${optInPresPr}`)
	assert.match(optInPresPr, /<p15:chartTrackingRefBased\b[^>]*\bval="1"/, `chartTrackingRefBased val missing: ${optInPresPr}`)
	assert.match(optInPresPr, /xmlns:p15="http:\/\/schemas\.microsoft\.com\/office\/powerpoint\/2012\/main"/, 'p15 2012/main namespace missing')

	const [unsetChart, optInChart] = await Promise.all([readChartXml(unsetZip), readChartXml(optInZip)])
	assert.equal(optInChart, unsetChart, 'chart XML must not change when chartTrackingRefBased is set')
})

test('contract: modern comment parts satisfy MS-PPTX §2.1.5–2.1.6', async () => {
	const pptx = new pptxgen()
	pptx.commentAuthors = [{ name: 'Ada Lovelace', initials: 'AL' }]
	pptx.addSlide().addComment({ text: 'Review this', author: 0, x: 1, y: 1, replies: [{ text: 'Done', author: 'Grace' }] })
	const commentsZip = await JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	await assertPptxPackageContracts(commentsZip)
	await assertModernCommentPartContracts(commentsZip)
})

test('contract: rejects a comments part without the §2.1.5 content type', async () => {
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
