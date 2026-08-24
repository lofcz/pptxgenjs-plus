/**
 * Text measurement and fitting.
 *
 * No DOM here, so `measureText` always takes the bundled-metrics tier. The canvas tier is
 * exercised by the same wrapping code, only with a different width source.
 */
import { test } from 'bun:test'
import assert from 'node:assert/strict'
import { measureText, registerFontMetrics, fitText, checkOverflow } from '../src/text'

/** Calibri advance widths, per mille of the em, from the generated table */
const SPACE = 0.226
const DIGIT = 0.507

test('measureText: uses bundled metrics for Calibri, not a guess', () => {
	const result = measureText('0', { fontSize: 72 })
	assert.equal(result.source, 'metrics')
	// 72pt is exactly one inch of em, so an em-width is an inch-width
	assert.ok(Math.abs(result.w - DIGIT) < 1e-9, `expected ${DIGIT}", got ${result.w}"`)
	assert.equal(result.lines.length, 1)
})

test('measureText: width scales with font size and is additive across characters', () => {
	const one = measureText('0', { fontSize: 12 }).w
	const four = measureText('0000', { fontSize: 12 }).w
	assert.ok(Math.abs(four - one * 4) < 1e-9)
	assert.ok(Math.abs(measureText('0', { fontSize: 24 }).w - one * 2) < 1e-9)
})

test('measureText: a space is narrower than a digit - proportional, not monospace', () => {
	const space = measureText(' ', { fontSize: 72 }).w
	assert.ok(Math.abs(space - SPACE) < 1e-9)
	assert.ok(space < measureText('0', { fontSize: 72 }).w, 'monospace metrics would make these equal')
})

test('measureText: bold is wider than regular at the same size', () => {
	const regular = measureText('Handgloves', { fontSize: 18 }).w
	const bold = measureText('Handgloves', { fontSize: 18, bold: true }).w
	assert.ok(bold > regular, `expected bold (${bold}) > regular (${regular})`)
})

test('measureText: without w, text is one line and newlines still split', () => {
	assert.equal(measureText('a very long line of text with no wrap width given').lines.length, 1)
	assert.deepEqual(measureText('one\ntwo\nthree').lines, ['one', 'two', 'three'])
})

test('measureText: wraps at the given width and never exceeds it', () => {
	const text = 'The quick brown fox jumps over the lazy dog and keeps going for a while'
	const result = measureText(text, { w: 2, fontSize: 12 })
	assert.ok(result.lines.length > 1, 'expected the text to wrap')
	assert.ok(result.w <= 2 + 1e-9, `widest line ${result.w}" exceeds the 2" box`)
	// wrapping is lossless: the words come back in order
	assert.equal(result.lines.join(' ').replace(/\s+/g, ' ').trim(), text)
})

test('measureText: height is line count times line height', () => {
	const result = measureText('The quick brown fox jumps over the lazy dog', { w: 1.5, fontSize: 12 })
	assert.ok(Math.abs(result.h - result.lines.length * result.lineHeight) < 1e-9)
	// Calibri: ascent 952 + |descent| 269 = 1.221em, so 12pt lines are ~0.2035"
	assert.ok(Math.abs(result.lineHeight - (1.221 * 12) / 72) < 1e-3, `unexpected line height ${result.lineHeight}`)
})

test('measureText: lineSpacingMultiple scales height, not width', () => {
	const single = measureText('some text here', { w: 1, fontSize: 12 })
	const double = measureText('some text here', { w: 1, fontSize: 12, lineSpacingMultiple: 2 })
	assert.ok(Math.abs(double.h - single.h * 2) < 1e-9)
	assert.ok(Math.abs(double.w - single.w) < 1e-9)
})

test('measureText: a word wider than the box is broken instead of overflowing', () => {
	const result = measureText('Donaudampfschifffahrtsgesellschaftskapitaen', { w: 0.5, fontSize: 14 })
	assert.ok(result.lines.length > 1, 'expected the word to be broken')
	assert.ok(result.w <= 0.5 + 1e-9, `line ${result.w}" exceeds the 0.5" box`)
	assert.equal(result.lines.join(''), 'Donaudampfschifffahrtsgesellschaftskapitaen')
})

test('measureText: an unknown font falls back to an estimate and says so', () => {
	const result = measureText('some text', { fontFace: 'No Such Typeface 9000' })
	assert.equal(result.source, 'estimate')
	assert.ok(result.w > 0)
})

test('measureText: rejects impossible inputs instead of returning nonsense', () => {
	assert.throws(() => measureText('x', { fontSize: 0 }), /fontSize must be > 0/)
	assert.throws(() => measureText('x', { w: 0 }), /w must be > 0/)
	assert.throws(() => measureText('x', { lineSpacingMultiple: 0 }), /lineSpacingMultiple must be > 0/)
})

test('registerFontMetrics: a registered font is measured, not estimated', () => {
	registerFontMetrics('Test Mono', { widths: { ' ': 0.6, a: 0.6, b: 0.6 }, fallbackWidth: 0.6, ascent: 1, descent: -0.25 })
	const result = measureText('aab', { fontFace: 'Test Mono', fontSize: 72 })
	assert.equal(result.source, 'metrics')
	assert.ok(Math.abs(result.w - 1.8) < 1e-9, `expected 1.8", got ${result.w}"`)
	assert.ok(Math.abs(result.lineHeight - 1.25) < 1e-9)
})

test('registerFontMetrics: codepoint keys work and bad keys are rejected', () => {
	registerFontMetrics('Test Codepoints', { widths: { 65: 0.9 } })
	assert.ok(Math.abs(measureText('A', { fontFace: 'Test Codepoints', fontSize: 72 }).w - 0.9) < 1e-9)
	assert.throws(() => registerFontMetrics('Test Bad', { widths: { ab: 1 } }), /neither a single character nor a codepoint/)
	assert.throws(() => registerFontMetrics('Test Empty', { widths: {} }), /widths is empty/)
	assert.throws(() => registerFontMetrics('', { widths: { a: 1 } }), /fontFace is required/)
})

test('registerFontMetrics: a style falls back to the regular weight, not to an estimate', () => {
	registerFontMetrics('Test Fallback', { widths: { a: 0.5 } })
	assert.equal(measureText('a', { fontFace: 'Test Fallback', bold: true }).source, 'metrics')
})

test('fitText: picks the largest whole point size that fits', () => {
	const area = { w: 3, h: 1 }
	const result = fitText(area, 'A headline that has to be shrunk to fit this box', { max: 40 })
	assert.equal(result.overflows, false)
	assert.ok(result.h <= area.h + 1e-9, `fitted height ${result.h}" exceeds ${area.h}"`)
	assert.ok(result.w <= area.w + 1e-9)
	assert.equal(result.fontSize, Math.round(result.fontSize), 'whole points only')

	// one point larger must not fit, or it was not the largest
	const next = measureText('A headline that has to be shrunk to fit this box', { fontSize: result.fontSize + 1, w: area.w })
	assert.ok(next.h > area.h || next.w > area.w, `${result.fontSize + 1}pt also fits - not the largest`)
})

test('fitText: short text keeps the maximum size', () => {
	const result = fitText({ w: 8, h: 3 }, 'Hi', { max: 32 })
	assert.equal(result.fontSize, 32)
	assert.equal(result.overflows, false)
})

test('fitText: reports overflow rather than pretending to fit', () => {
	const result = fitText({ w: 1, h: 0.3 }, 'far too much text to ever fit into a box this small at any size', { min: 10 })
	assert.equal(result.overflows, true)
	assert.equal(result.fontSize, 10)
})

test('fitText: margin shrinks the usable area', () => {
	const plain = fitText({ w: 4, h: 2 }, 'Some text that needs a few lines to lay out', { max: 40 })
	const inset = fitText({ w: 4, h: 2 }, 'Some text that needs a few lines to lay out', { max: 40, margin: 0.5 })
	assert.ok(inset.fontSize < plain.fontSize, `margin should force a smaller size (${inset.fontSize} vs ${plain.fontSize})`)
})

test('fitText: rejects impossible areas and ranges', () => {
	assert.throws(() => fitText({ w: 0, h: 1 }, 'x'), /positive w and h/)
	assert.throws(() => fitText({ w: 1, h: 1 }, 'x', { min: 0 }), /min must be > 0/)
	assert.throws(() => fitText({ w: 1, h: 1 }, 'x', { min: 20, max: 10 }), /below min/)
	assert.throws(() => fitText({ w: 1, h: 1 }, 'x', { margin: 1 }), /leaves no room/)
})

test('checkOverflow: reports the spill, not just that there is one', () => {
	const long = 'A paragraph long enough that eleven point type cannot possibly fit it into half an inch of height, no matter how wide the box gets.'
	const spills = checkOverflow({ w: 3, h: 0.5 }, long, { fontSize: 11 })
	assert.ok(spills.overflows, 'text taller than its box overflows')
	assert.ok(spills.overflowBy > 0, `expected a positive overflow, got ${spills.overflowBy}`)
	assert.equal(spills.overflowBy, spills.h - 0.5, 'overflowBy is the height beyond the area')

	const fits = checkOverflow({ w: 6, h: 4 }, 'Short line', { fontSize: 11 })
	assert.equal(fits.overflows, false)
	assert.equal(fits.overflowBy, 0)
})

test('checkOverflow: a word too long for the line is broken, then counted in height', () => {
	const result = checkOverflow({ w: 0.4, h: 0.3 }, 'Wirtschaftspruefungsgesellschaft', { fontSize: 18 })
	assert.ok(result.lines.length > 1, 'the word was broken across lines')
	assert.ok(result.overflowBy > 0, 'those lines do not fit 0.3in')
})

test('checkOverflow: agrees with fitText at the size fitText picked', () => {
	const text = 'Ergebnisqualitaet und Marktposition im Vergleich zum Wettbewerb'
	const area = { w: 3.5, h: 1.2 }
	const { fontSize } = fitText(area, text, { max: 40 })
	assert.equal(checkOverflow(area, text, { fontSize }).overflows, false, 'the chosen size must fit')
	assert.equal(checkOverflow(area, text, { fontSize: fontSize + 1 }).overflows, true, 'one point larger must not')
})

test('checkOverflow: rejects impossible areas', () => {
	assert.throws(() => checkOverflow({ w: 0, h: 1 }, 'x'), /positive w and h/)
	assert.throws(() => checkOverflow({ w: 1, h: 1 }, 'x', { margin: 1 }), /leaves no room/)
})
