/**
 * Measured table pagination and the HTML importer.
 *
 * Both helpers are typed structurally, so the presentation, the slides and the `<table>` are all
 * plain objects here - the assertions are about which rows land on which slide, not about XML.
 */
import { test } from 'bun:test'
import assert from 'node:assert/strict'
import type PptxGenJS from 'pptxgenjs-plus'
import pptxgen from '../../../src/pptxgen'
import { paginateTable, tableFromHtml, cssColorToHex } from '../src/tables'

interface RecordedTable {
	rows: PptxGenJS.TableRow[]
	options?: PptxGenJS.TableProps
}

/** A presentation double that records what would have been added */
function fakePres (): { addSlide: (options?: { masterName?: string }) => { addTable: (rows: PptxGenJS.TableRow[], options?: PptxGenJS.TableProps) => unknown, tables: RecordedTable[] }, slides: Array<{ masterName?: string, tables: RecordedTable[] }> } {
	const slides: Array<{ masterName?: string, tables: RecordedTable[] }> = []
	return {
		slides,
		addSlide (options) {
			const slide = {
				masterName: options?.masterName,
				tables: [] as RecordedTable[],
				addTable (rows: PptxGenJS.TableRow[], tableOptions?: PptxGenJS.TableProps) {
					slide.tables.push({ rows, options: tableOptions })
					return slide
				},
			}
			slides.push(slide)
			return slide
		},
	}
}

const row = (text: string): PptxGenJS.TableRow => [{ text }]

test('paginateTable: short tables stay on one slide', () => {
	const pres = fakePres()
	const result = paginateTable(pres, [row('a'), row('b'), row('c')], { x: 0.5, y: 0.5, w: 9 })
	assert.equal(result.slides.length, 1)
	assert.deepEqual(result.rowsPerSlide, [3])
	assert.equal(pres.slides[0].tables[0].rows.length, 3)
})

test('paginateTable: rows that do not fit go to a new slide', () => {
	const pres = fakePres()
	// 40 rows of ~0.3" each cannot fit the ~4.6" usable height of a 16:9 slide
	const rows = Array.from({ length: 40 }, (_, i) => row(`row ${i}`))
	const result = paginateTable(pres, rows, { x: 0.5, y: 0.5, w: 9, fontSize: 12 })

	assert.ok(result.slides.length > 1, 'expected more than one slide')
	assert.equal(result.rowsPerSlide.reduce((a, b) => a + b, 0), rows.length, 'every row is placed exactly once')
	const placed = pres.slides.flatMap(slide => slide.tables[0].rows.map(r => r[0].text))
	assert.deepEqual(placed, rows.map(r => r[0].text), 'rows keep their order across slides')
})

test('paginateTable: taller cells mean fewer rows per slide', () => {
	const short = fakePres()
	const tall = fakePres()
	const rows = Array.from({ length: 30 }, (_, i) => row(`row ${i}`))
	const wrapping = Array.from({ length: 30 }, () => row('a sentence long enough to wrap onto several lines inside a narrow column'))

	paginateTable(short, rows, { x: 0.5, y: 0.5, w: 3, fontSize: 12 })
	paginateTable(tall, wrapping, { x: 0.5, y: 0.5, w: 3, fontSize: 12 })
	assert.ok(tall.slides.length > short.slides.length, `wrapped rows should need more slides (${tall.slides.length} vs ${short.slides.length})`)
})

test('paginateTable: header rows repeat on every slide', () => {
	const pres = fakePres()
	const rows = [row('HEADER'), ...Array.from({ length: 40 }, (_, i) => row(`row ${i}`))]
	const result = paginateTable(pres, rows, { x: 0.5, y: 0.5, w: 9, repeatHeaderRows: 1 })

	assert.ok(result.slides.length > 1)
	for (const slide of pres.slides) assert.equal(slide.tables[0].rows[0][0].text, 'HEADER')
	// the header is not counted as a body row on any slide but the first
	const body = pres.slides.flatMap(slide => slide.tables[0].rows.slice(1).map(r => r[0].text))
	assert.equal(body.length, 40)
})

test('paginateTable: continueY places the tables after the first', () => {
	const pres = fakePres()
	const rows = Array.from({ length: 40 }, (_, i) => row(`row ${i}`))
	paginateTable(pres, rows, { x: 0.5, y: 2, w: 9, continueY: 0.4 })

	assert.equal(pres.slides[0].tables[0].options?.y, 2)
	for (const slide of pres.slides.slice(1)) assert.equal(slide.tables[0].options?.y, 0.4)
})

test('paginateTable: passes options through and resolves column widths', () => {
	const pres = fakePres()
	paginateTable(pres, [[{ text: 'a' }, { text: 'b' }]], { x: 1, y: 1, w: 8, fontSize: 9, border: { pt: 1, color: 'CCCCCC' } })
	const options = pres.slides[0].tables[0].options
	assert.equal(options?.fontSize, 9)
	assert.deepEqual(options?.border, { pt: 1, color: 'CCCCCC' })
	assert.deepEqual(options?.colW, [4, 4], 'width is split evenly when colW is not given')
})

test('paginateTable: explicit colW is honored and validated', () => {
	const pres = fakePres()
	paginateTable(pres, [[{ text: 'a' }, { text: 'b' }]], { x: 1, y: 1, w: 8, colW: [6, 2] })
	assert.deepEqual(pres.slides[0].tables[0].options?.colW, [6, 2])
	assert.throws(() => paginateTable(fakePres(), [[{ text: 'a' }, { text: 'b' }]], { colW: [1, 2, 3] }), /3 widths for 2 columns/)
})

test('paginateTable: masterName is applied to every slide it creates', () => {
	const pres = fakePres()
	paginateTable(pres, Array.from({ length: 40 }, (_, i) => row(`row ${i}`)), { masterName: 'DATA_SLIDE' })
	assert.ok(pres.slides.length > 1)
	for (const slide of pres.slides) assert.equal(slide.masterName, 'DATA_SLIDE')
})

test('paginateTable: reports when a measurement was a guess', () => {
	assert.equal(paginateTable(fakePres(), [row('a')], { w: 9 }).estimated, false)
	assert.equal(paginateTable(fakePres(), [row('a')], { w: 9, fontFace: 'No Such Typeface 9000' }).estimated, true)
})

test('paginateTable: slideHeight and presLayout both bound the usable height', () => {
	const rows = Array.from({ length: 30 }, (_, i) => row(`row ${i}`))
	const short = fakePres()
	const tall = fakePres()
	paginateTable(short, rows, { slideHeight: 3 })
	paginateTable(tall, rows, { slideHeight: 10 })
	assert.ok(short.slides.length > tall.slides.length, 'a shorter slide holds fewer rows')

	const fromLayout = { ...fakePres(), presLayout: { width: 10 * 914400, height: 3 * 914400 } }
	const layoutResult = paginateTable(fromLayout, rows)
	assert.deepEqual(layoutResult.rowsPerSlide, paginateTable(fakePres(), rows, { slideHeight: 3 }).rowsPerSlide)
})

test('paginateTable: rejects inputs it cannot lay out', () => {
	assert.throws(() => paginateTable(fakePres(), []), /at least one row/)
	assert.throws(() => paginateTable(fakePres(), [row('a')], { repeatHeaderRows: 1 }), /leaves no body rows/)
	assert.throws(() => paginateTable(fakePres(), [row('a')], { repeatHeaderRows: -1 }), /integer >= 0/)
	assert.throws(() => paginateTable(fakePres(), [[]]), /no cells/)
	assert.throws(() => paginateTable(fakePres(), [row('a')], { w: -1 }), /width must be > 0/)
})

/** A `<table>` double: the same shape a real HTMLTableElement presents */
function fakeTable (rows: Array<Array<{ text: string, tag?: string, colSpan?: number, rowSpan?: number }>>) {
	return {
		rows: rows.map(cells => ({
			cells: cells.map(cell => ({
				textContent: cell.text,
				tagName: cell.tag ?? 'TD',
				colSpan: cell.colSpan ?? 1,
				rowSpan: cell.rowSpan ?? 1,
			})),
		})),
	}
}

test('tableFromHtml: cells become table rows, trimmed', () => {
	const pres = fakePres()
	tableFromHtml(pres, fakeTable([[{ text: '  a  ' }, { text: 'b' }]]), { w: 9 })
	const rows = pres.slides[0].tables[0].rows
	assert.equal(rows.length, 1)
	assert.deepEqual(rows[0].map(cell => cell.text), ['a', 'b'])
})

test('tableFromHtml: th cells are bold and leading th rows repeat as headers', () => {
	const pres = fakePres()
	const body = Array.from({ length: 40 }, (_, i) => [{ text: `row ${i}` }])
	tableFromHtml(pres, fakeTable([[{ text: 'H1', tag: 'TH' }], ...body]), { w: 9 })

	assert.ok(pres.slides.length > 1, 'expected pagination')
	for (const slide of pres.slides) {
		assert.equal(slide.tables[0].rows[0][0].text, 'H1')
		assert.equal(slide.tables[0].rows[0][0].options?.bold, true)
	}
})

test('tableFromHtml: header detection can be turned off', () => {
	const pres = fakePres()
	const body = Array.from({ length: 40 }, (_, i) => [{ text: `row ${i}` }])
	tableFromHtml(pres, fakeTable([[{ text: 'H1', tag: 'TH' }], ...body]), { w: 9, detectHeaderRows: false })
	assert.notEqual(pres.slides[1].tables[0].rows[0][0].text, 'H1')
})

test('tableFromHtml: colspan and rowspan carry over', () => {
	const pres = fakePres()
	tableFromHtml(pres, fakeTable([[{ text: 'wide', colSpan: 2 }, { text: 'tall', rowSpan: 2 }], [{ text: 'a' }, { text: 'b' }]]), { w: 9 })
	const [first] = pres.slides[0].tables[0].rows
	assert.equal(first[0].options?.colspan, 2)
	assert.equal(first[1].options?.rowspan, 2)
	// a colspan of 1 adds no option at all
	assert.equal(pres.slides[0].tables[0].rows[1][0].options?.colspan, undefined)
})

test('tableFromHtml: a styleOf hook supplies cell formatting', () => {
	const pres = fakePres()
	tableFromHtml(pres, fakeTable([[{ text: 'a' }]]), {
		w: 9,
		styleOf: () => ({ color: 'FF0000', fill: '00FF00', bold: true, fontSize: 14, align: 'center' }),
	})
	const cell = pres.slides[0].tables[0].rows[0][0]
	assert.equal(cell.options?.color, 'FF0000')
	assert.deepEqual(cell.options?.fill, { color: '00FF00' })
	assert.equal(cell.options?.bold, true)
	assert.equal(cell.options?.fontSize, 14)
	assert.equal(cell.options?.align, 'center')
})

test('tableFromHtml: rejects an empty table', () => {
	assert.throws(() => tableFromHtml(fakePres(), { rows: [] }), /no rows/)
})

test('cssColorToHex: converts the forms a browser reports', () => {
	assert.equal(cssColorToHex('rgb(255, 0, 0)'), 'FF0000')
	assert.equal(cssColorToHex('rgba(0, 128, 255, 0.5)'), '0080FF')
	assert.equal(cssColorToHex('#abc'), 'AABBCC')
	assert.equal(cssColorToHex('#A1B2C3'), 'A1B2C3')
	// transparent is "no fill", not black
	assert.equal(cssColorToHex('rgba(0, 0, 0, 0)'), undefined)
	assert.equal(cssColorToHex('transparent'), undefined)
	assert.equal(cssColorToHex(''), undefined)
	assert.equal(cssColorToHex(null), undefined)
})

test('paginateTable: drives the real core, producing one slide per chunk', async () => {
	const pptx = new pptxgen()
	pptx.layout = 'LAYOUT_16x9'
	const rows = Array.from({ length: 60 }, (_, i) => row(`row ${i} with enough text to take up a line`))
	const result = paginateTable(pptx, rows, { x: 0.5, y: 0.5, w: 9, repeatHeaderRows: 0, fontSize: 11 })

	assert.ok(result.slides.length > 1, 'expected the rows to need more than one slide')
	assert.equal(pptx.slides.length, result.slides.length)
	const written = (await pptx.write({ outputType: 'nodebuffer' })) as Buffer
	assert.ok(written.length > 0)
})

test('paginateTable: reads the real presentation layout instead of assuming 16:9', () => {
	const wide = new pptxgen()
	wide.layout = 'LAYOUT_16x9'
	const tall = new pptxgen()
	tall.defineLayout({ name: 'TALL', width: 10, height: 12 })
	tall.layout = 'TALL'

	const rows = Array.from({ length: 40 }, (_, i) => row(`row ${i}`))
	const onWide = paginateTable(wide, rows, { x: 0.5, y: 0.5, w: 9 })
	const onTall = paginateTable(tall, rows, { x: 0.5, y: 0.5, w: 9 })
	assert.ok(onTall.slides.length < onWide.slides.length, `a 12" slide should hold more rows (${onTall.slides.length} vs ${onWide.slides.length})`)
})
