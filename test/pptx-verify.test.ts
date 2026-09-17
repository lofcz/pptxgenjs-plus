/**
 * Generates PPTX files and verifies them through the shared `verifyPptx` report.
 * Package contracts always run. The PowerPoint sidecar is opt-in via `test:powerpoint`.
 */
import assert from 'node:assert/strict'
import { test } from 'node:test'
import { JSZip } from '@node-projects/jszip'
import pptxgen from '../src/pptxgen'
import { assertVerified, isPowerPointResult, verifyPptx, writeAndVerify } from './pptx-verify'

const PIXEL_PNG = 'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg=='

test('verify: text slide returns package inventory and extracted text', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addText('Hello verify', { x: 0.5, y: 0.5, w: 6, h: 0.5 })
	const report = await writeAndVerify(pptx)

	assertVerified(report)
	assert.equal(report.zip, true)
	assert.equal(report.slides.length, 1)
	assert.equal(report.presentation?.slideCount, 1)
	assert.ok(report.slides[0].texts.includes('Hello verify'))
	assert.ok(report.slides[0].objects.some(object => object.kind === 'text' && object.text.includes('Hello verify')))
	assert.ok(report.parts.some(part => part.name === 'ppt/slides/slide1.xml'))
	assert.ok(report.masters.length >= 1)
	assert.ok(report.layouts.length >= 1)
	assert.ok((report.presentation?.slideWidthIn ?? 0) > 0)
	assert.ok((report.presentation?.slideHeightIn ?? 0) > 0)
	assert.equal(report.issues.length, 0)
})

test('verify: multi-slide deck preserves sldId order and per-slide text', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addText('one', { x: 0.5, y: 0.5, w: 4, h: 0.4 })
	pptx.addSlide().addText('two', { x: 0.5, y: 0.5, w: 4, h: 0.4 })
	pptx.addSlide().addText('three', { x: 0.5, y: 0.5, w: 4, h: 0.4 })
	const report = await writeAndVerify(pptx)

	assertVerified(report)
	assert.deepEqual(report.slides.map(slide => slide.part), [
		'ppt/slides/slide1.xml',
		'ppt/slides/slide2.xml',
		'ppt/slides/slide3.xml',
	])
	assert.deepEqual(report.slides.map(slide => slide.texts.join('')), ['one', 'two', 'three'])
	assert.deepEqual(report.presentation?.sldIds.map(entry => entry.part), report.slides.map(slide => slide.part))
	assert.ok(report.presentation?.sldIds.every(entry => entry.rId && entry.id))
})

test('verify: shape, table, chart, and image show up as typed objects', async () => {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addShape(pptx.ShapeType.rect, { x: 0.5, y: 0.4, w: 2, h: 1, fill: { color: 'FF0000' } })
	slide.addTable([['A', 'B'], ['1', '2']], { x: 0.5, y: 1.6, w: 4 })
	slide.addChart(pptx.ChartType.bar, [{ name: 'Sales', labels: ['Q1', 'Q2'], values: [10, 20] }], { x: 5, y: 0.4, w: 4, h: 3 })
	slide.addImage({ data: `image/png;base64,${PIXEL_PNG}`, x: 0.5, y: 4.2, w: 1, h: 1 })
	const report = await writeAndVerify(pptx)

	assertVerified(report)
	const kinds = new Set(report.slides[0].objects.map(object => object.kind))
	assert.ok(kinds.has('shape'), JSON.stringify(report.slides[0].objects))
	assert.ok(kinds.has('table'), JSON.stringify(report.slides[0].objects))
	assert.ok(kinds.has('chart'), JSON.stringify(report.slides[0].objects))
	assert.ok(kinds.has('image'), JSON.stringify(report.slides[0].objects))
	assert.ok(report.charts.length >= 1)
	assert.ok(report.embeddings.length >= 1)
	assert.ok(report.media.length >= 1)
	assert.ok(report.slides[0].objects.some(object => object.kind === 'shape' && object.preset === 'rect'))
	assert.ok(report.slides[0].texts.includes('A'))
	assert.ok(report.slides[0].texts.includes('1'))
})

test('verify: undeclared package part fails contracts and stays in the inventory', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addText('ok', { x: 0.5, y: 0.5, w: 3, h: 0.4 })
	const good = await writeAndVerify(pptx)
	assertVerified(good)

	const zip = await JSZip.loadAsync(await zipBytes(pptx))
	zip.file('ppt/undeclared.bin', 'invalid', { createFolders: false })
	const report = await verifyPptx(zip)

	assert.equal(report.ok, false)
	assert.equal(report.zip, true)
	assert.ok(report.parts.some(part => part.name === 'ppt/undeclared.bin'))
	assert.ok(report.issues.some(item => item.code === 'package-contract' && /package part has no content type/.test(item.message)))
})

test('verify: non-package bytes are reported as not-zip', async () => {
	const report = await verifyPptx(Buffer.from('this is not a pptx'))
	assert.equal(report.ok, false)
	assert.equal(report.zip, false)
	assert.equal(report.slides.length, 0)
	assert.ok(report.issues.some(item => item.code === 'not-zip'))
})

test('verify: repair fixtures carry the intended mutations', async () => {
	const { buildRepairFixtures, fixturePath } = await import('./repair-fixtures')
	await buildRepairFixtures()
	const invalid = await verifyPptx(fixturePath('invalid-cnvid'))
	assert.ok(invalid.slides[0].objects.some(object => object.id === 'abc'), JSON.stringify(invalid.slides[0].objects))
	const emptyId = await verifyPptx(fixturePath('empty-cnvid'))
	assert.ok(emptyId.slides[0].objects.some(object => object.id === ''), JSON.stringify(emptyId.slides[0].objects))
	const preset = await verifyPptx(fixturePath('invalid-preset'))
	assert.ok(preset.slides[0].objects.some(object => object.preset === 'notashape'), JSON.stringify(preset.slides[0].objects))
	const baseline = await verifyPptx(fixturePath('ok-baseline'))
	assertVerified(baseline)
})

test('verify: PowerPoint sidecar stays skipped unless opted in', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addText('sidecar off', { x: 0.5, y: 0.5, w: 4, h: 0.4 })
	const report = await writeAndVerify(pptx)
	assertVerified(report)
	assert.equal(report.powerpoint, undefined)
	assert.equal(isPowerPointResult(report.powerpoint), false)
})

async function zipBytes (pptx: pptxgen): Promise<Buffer> {
	return (await pptx.write({ outputType: 'nodebuffer' })) as Buffer
}
