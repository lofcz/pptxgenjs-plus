/** Contracts for `pptxgenjs-plus-std/layout` — pure arithmetic, asserted directly. */
import { test } from 'bun:test'
import assert from 'node:assert/strict'
import pptxgen from '../../../src/pptxgen'
import { grid, gridFor, row, column, cm, pt } from '../src/layout'

/** Grid maths lands on repeating fractions; compare at a tolerance no slide can render past. */
const near = (actual: number, expected: number, what: string): void => {
	assert.ok(Math.abs(actual - expected) < 1e-9, `${what}: expected ${expected}, got ${actual}`)
}

test('grid: spans tile the slide without gaps or overlap', () => {
	const at = grid({ w: 10, h: 5.625, cols: 12, rows: 6, gutter: 0.2, margin: 0.5 })
	const full = at(0, 0, 12, 6)
	assert.equal(full.x, 0.5)
	assert.equal(full.y, 0.5)
	near(full.w, 9, 'a full-width span fills between the margins')
	near(full.h, 4.625, 'a full-height span fills between the margins')

	const left = at(0, 0, 6, 1)
	const right = at(6, 0, 6, 1)
	near(left.x + left.w + 0.2, right.x, 'adjacent columns are exactly one gutter apart')
	near(right.x + right.w, 9.5, 'the last column ends at the right margin')

	const top = at(0, 0, 1, 3)
	const bottom = at(0, 3, 1, 3)
	near(top.y + top.h + 0.2, bottom.y, 'adjacent rows are exactly one gutter apart')
	near(bottom.y + bottom.h, 5.125, 'the last row ends at the bottom margin')
})

test('grid: spans absorb the gutters they cross', () => {
	const at = grid({ w: 10, h: 5.625, cols: 4, rows: 4, gutter: 0.5, margin: 0.5 })
	const one = at(0, 0)
	const two = at(0, 0, 2)
	near(two.w, one.w * 2 + 0.5, 'a colSpan of 2 swallows one gutter')
	near(at(0, 0, 1, 3).h, one.h * 3 + 1, 'a rowSpan of 3 swallows two gutters')
})

test('grid: a single cell fills the whole usable area', () => {
	const cell = grid({ w: 8, h: 6, cols: 1, rows: 1, gutter: 0.3, margin: 1 })(0, 0)
	assert.deepEqual(cell, { x: 1, y: 1, w: 6, h: 4 }, 'one cell means margins only, no gutters')
})

test('grid: custom geometry is honored rather than snapped to the defaults', () => {
	const at = grid({ w: 13.333, h: 7.5, cols: 3, rows: 2, gutter: 0.25, margin: 0.4 })
	const cell = at(2, 1)
	near(cell.w, (13.333 - 0.8 - 0.5) / 3, 'column width divides the space left by margins and gutters')
	near(cell.h, (7.5 - 0.8 - 0.25) / 2, 'row height divides the space left by margins and gutters')
	near(cell.x + cell.w, 13.333 - 0.4, 'the last column ends at the right margin')
	near(cell.y + cell.h, 7.5 - 0.4, 'the last row ends at the bottom margin')
})

test('grid: out-of-range placement throws instead of drifting off-slide', () => {
	const at = grid({ cols: 12, rows: 6 })
	assert.throws(() => at(12, 0), /exceeds 12 columns/)
	assert.throws(() => at(0, 6), /exceeds 6 rows/)
	assert.throws(() => at(10, 0, 3), /exceeds 12 columns/)
	assert.throws(() => at(0, 4, 1, 3), /exceeds 6 rows/)
	assert.throws(() => at(-1, 0), /col and row must be >= 0/)
	assert.throws(() => at(0, -1), /col and row must be >= 0/)
	assert.throws(() => at(0, 0, 0), /spans must be >= 1/)
	assert.throws(() => at(0, 0, 1, 0), /spans must be >= 1/)
	assert.throws(() => at(1.5, 0), /must be integers/)
	assert.throws(() => at(0, 0.5), /must be integers/)
})

test('grid: an unusable geometry is rejected at construction, not at placement', () => {
	assert.throws(() => grid({ w: 10, h: 5, margin: 6 }), /leave no room/)
	assert.throws(() => grid({ w: 10, h: 5, cols: 40, gutter: 0.5 }), /leave no room/)
	assert.throws(() => grid({ cols: 0 }), /must be >= 1/)
	assert.throws(() => grid({ rows: 0 }), /must be >= 1/)
})

test('grid: every call returns a fresh area', () => {
	const at = grid()
	const first = at(0, 0)
	const second = at(0, 0)
	assert.notEqual(first, second, 'areas must not be shared - callers spread and mutate them')
	assert.deepEqual(first, second)
})

test('gridFor: tracks the presentation layout', () => {
	const pptx = new pptxgen()
	pptx.layout = 'LAYOUT_4x3' // 10 x 7.5 inches
	const cell = gridFor(pptx, { cols: 2, rows: 2, gutter: 0, margin: 0 })(0, 0)
	assert.equal(cell.w, 5)
	assert.equal(cell.h, 3.75)
})

test('gridFor: follows a custom layout', () => {
	const pptx = new pptxgen()
	pptx.defineLayout({ name: 'A4', width: 11.7, height: 8.3 })
	pptx.layout = 'A4'
	const cell = gridFor(pptx, { cols: 1, rows: 1, gutter: 0, margin: 0 })(0, 0)
	near(cell.w, 11.7, 'width comes from the defined layout')
	near(cell.h, 8.3, 'height comes from the defined layout')
})

test('row: equal slots span the area with gaps between them', () => {
	const area = { x: 1, y: 2, w: 9, h: 4 }
	const [a, b, c] = row(area, 3, 0.3)
	assert.equal(a.x, 1)
	assert.equal(a.y, 2)
	assert.equal(a.h, 4, 'the cross axis is untouched')
	// 9" less two 0.3" gaps, split three ways
	assert.ok(Math.abs(a.w - (9 - 0.6) / 3) < 1e-9)
	assert.ok(Math.abs(b.x - (a.x + a.w + 0.3)) < 1e-9)
	assert.ok(Math.abs(c.x + c.w - (area.x + area.w)) < 1e-9, 'the last slot ends where the area does')
})

test('row: weights divide the remaining space proportionally', () => {
	const [left, right] = row({ x: 0, y: 0, w: 10, h: 5 }, [1, 3], 0)
	assert.ok(Math.abs(left.w - 2.5) < 1e-9)
	assert.ok(Math.abs(right.w - 7.5) < 1e-9)
	assert.ok(Math.abs(right.x - 2.5) < 1e-9)
})

test('column: stacks slots down the area', () => {
	const [top, bottom] = column({ x: 1, y: 1, w: 8, h: 4.5 }, [1, 2], 0.3)
	assert.equal(top.x, 1)
	assert.equal(top.w, 8, 'the cross axis is untouched')
	assert.equal(top.y, 1)
	assert.ok(Math.abs(top.h - (4.5 - 0.3) / 3) < 1e-9)
	assert.ok(Math.abs(bottom.y - (top.y + top.h + 0.3)) < 1e-9)
	assert.ok(Math.abs(bottom.y + bottom.h - 5.5) < 1e-9)
})

test('row/column: slots nest, because output and input are the same shape', () => {
	const [, body] = column({ x: 0.5, y: 0.5, w: 9, h: 4.625 }, [1, 4])
	const cells = row(body, 3)
	assert.equal(cells.length, 3)
	for (const cell of cells) {
		assert.ok(cell.y >= body.y - 1e-9 && cell.y + cell.h <= body.y + body.h + 1e-9, 'nested slots stay inside their parent')
	}
})

test('row/column: a single slot is the area minus nothing', () => {
	const area = { x: 1, y: 1, w: 4, h: 2 }
	assert.deepEqual(row(area, 1), [area])
	assert.deepEqual(column(area, 1, 5), [area], 'a gap needs two slots to matter')
})

test('row/column: reject what would silently misplace content', () => {
	const area = { x: 0, y: 0, w: 4, h: 2 }
	assert.throws(() => row(area, 0), /slot count must be an integer >= 1/)
	assert.throws(() => row(area, 1.5), /slot count must be an integer >= 1/)
	assert.throws(() => row(area, []), /weights must not be empty/)
	assert.throws(() => row(area, [1, 0]), /every weight must be > 0/)
	assert.throws(() => row(area, [1, -1]), /every weight must be > 0/)
	assert.throws(() => row(area, 2, -1), /gap must be >= 0/)
	assert.throws(() => row(area, 5, 1), /leaves no room across 4/)
	assert.throws(() => column(area, 5, 1), /leaves no room across 2/)
})

test('cm/pt: convert to the inches every addX option wants', () => {
	near(cm(2.54), 1, 'one inch of centimetres')
	near(cm(0), 0, 'zero converts')
	near(cm(-1.27), -0.5, 'negative offsets are legal positions')
	near(pt(72), 1, 'one inch of points')
	near(pt(18), 0.25, 'a quarter inch')
	// A4 landscape, the metric page these exist for
	near(cm(29.7), 11.692913385826772, 'A4 width')
})

test('cm/pt: reject what would place content at NaN', () => {
	assert.throws(() => cm(Number.NaN), /cm: value must be a finite number/)
	assert.throws(() => cm(Number.POSITIVE_INFINITY), /cm: value must be a finite number/)
	assert.throws(() => pt(Number.NaN), /pt: value must be a finite number/)
})
