/**
 * Opens a generated presentation with LibreOffice and converts it to PDF.
 * Run explicitly with PPTXGENJS_OFFICE_BIN set to libreoffice or soffice.
 */
import assert from 'node:assert/strict'
import { execFile as execFileCallback } from 'node:child_process'
import { mkdtemp, rm, stat, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { promisify } from 'node:util'
import { test } from 'node:test'
import pptxgen from '../src/pptxgen'

const officeBinary = process.env.PPTXGENJS_OFFICE_BIN
const execFile = promisify(execFileCallback)
const run = officeBinary ? test : test.skip

run('office: LibreOffice opens and converts a generated presentation', async () => {
	const directory = await mkdtemp(join(tmpdir(), 'pptxgenjs-office-'))
	const presentationPath = join(directory, 'smoke.pptx')

	try {
		const pptx = new pptxgen()
		const slide = pptx.addSlide()
		slide.addText('OOXML consumer smoke test', { x: 0.5, y: 0.5, w: 5, h: 0.5 })
		slide.addText(
			[
				{ text: 'Math: ' },
				{ text: '', options: { omml: '<m:oMath><m:r><m:t>E=mc^2</m:t></m:r></m:oMath>' } },
			],
			{ x: 0.5, y: 1.0, w: 5, h: 0.4 },
		)
		slide.addTable([['Region', 'Sales'], ['West', '20']], { x: 0.5, y: 1.5, w: 5 })
		slide.addChart(pptx.ChartType.bar, [{ name: 'Sales', labels: ['Q1', 'Q2'], values: [10, 20] }], { x: 0.5, y: 3, w: 6, h: 3 })
		slide.addText([{ text: '22/08/2026', options: { field: 'datetime1' } }, { text: ' page ' }, { text: '1', options: { field: 'slidenum' } }], { x: 6, y: 7.2, w: 3, h: 0.4 })
		slide.addShape(pptx.ShapeType.rect, { x: 6.5, y: 5.5, w: 2, h: 0.6, hyperlink: { url: 'https://example.com', tooltip: 'open' }, hyperlinkHover: { slide: 2, tooltip: 'peek' } })
		pptx.addSlide().addText('hover target', { x: 0.5, y: 0.5, w: 3, h: 0.4 })
		await writeFile(presentationPath, (await pptx.write({ outputType: 'nodebuffer' })) as Buffer)

		await execFile(officeBinary, ['--headless', '--convert-to', 'pdf', '--outdir', directory, presentationPath], { timeout: 60_000 })
		assert.ok((await stat(join(directory, 'smoke.pdf'))).size > 0, 'LibreOffice did not produce a PDF')
	} finally {
		await rm(directory, { recursive: true, force: true })
	}
})

test('office: LibreOffice opens a presentation with animation and transition', async () => {
	const directory = await mkdtemp(join(tmpdir(), 'pptxgenjs-office-anim-'))
	const presentationPath = join(directory, 'anim-trans.pptx')

	try {
		const pptx = new pptxgen()
		const slide = pptx.addSlide()
		slide.addText('Office animation/transition round-trip', {
			x: 0.5, y: 0.5, w: 8, h: 1,
			animation: { type: pptx.AnimationPreset.fadein, duration: 500 },
		})
		slide.addTransition({ type: pptx.TransitionType.fade, speed: 'med' })
		await writeFile(presentationPath, (await pptx.write({ outputType: 'nodebuffer' })) as Buffer)

		await execFile(officeBinary, ['--headless', '--convert-to', 'pdf', '--outdir', directory, presentationPath], { timeout: 60_000 })
		assert.ok((await stat(join(directory, 'anim-trans.pdf'))).size > 0, 'LibreOffice did not produce a PDF')
	} finally {
		await rm(directory, { recursive: true, force: true })
	}
})

