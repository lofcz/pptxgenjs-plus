/**
 * Opens generated and deliberately broken PPTX files with real PowerPoint
 * (hidden desktop sidecar). Skips when Office is not installed.
 */
import assert from 'node:assert/strict'
import { mkdtemp, readFile, rm, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { test } from 'node:test'
import { JSZip } from '@node-projects/jszip'
import pptxgen from '../src/pptxgen'
import { isPowerPointSidecarAvailable } from './powerpoint-verify'
import { isPowerPointResult, verifyPptx, writeAndVerify } from './pptx-verify'

const available = isPowerPointSidecarAvailable()
const run = available ? test : test.skip

async function writeGenerated (directory: string, name: string): Promise<string> {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addText('PowerPoint sidecar smoke', { x: 0.5, y: 0.5, w: 8, h: 0.5 })
	const path = join(directory, name)
	await writeFile(path, (await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
	return path
}

run('powerpoint: generated deck opens without a repair dialog', async () => {
	const pptx = new pptxgen()
	pptx.addSlide().addText('PowerPoint sidecar smoke', { x: 0.5, y: 0.5, w: 8, h: 0.5 })
	const report = await writeAndVerify(pptx, { powerpoint: true })
	assert.equal(report.ok, true, JSON.stringify(report.issues))
	assert.ok(isPowerPointResult(report.powerpoint), JSON.stringify(report.powerpoint))
	assert.equal(report.powerpoint.verdict, 'ok', JSON.stringify(report.powerpoint))
	assert.equal(report.powerpoint.opened, true)
})

run('powerpoint: invalid cNvPr id shows the Repair dialog', async () => {
	const directory = await mkdtemp(join(tmpdir(), 'pptxgenjs-ppt-repair-'))
	try {
		const path = await writeGenerated(directory, 'good.pptx')
		const zip = await JSZip.loadAsync(await readFile(path))
		const slideXml = await zip.file('ppt/slides/slide1.xml')!.async('string')
		zip.file('ppt/slides/slide1.xml', slideXml.replace('<p:cNvPr id="2"', '<p:cNvPr id="abc"'), { createFolders: false })
		const broken = join(directory, 'repair-cnvid.pptx')
		await writeFile(broken, await zip.generateAsync({ type: 'nodebuffer' }))
		const report = await verifyPptx(broken, { powerpoint: true })
		assert.ok(isPowerPointResult(report.powerpoint), JSON.stringify(report.powerpoint))
		assert.equal(report.powerpoint.verdict, 'repair', JSON.stringify(report.powerpoint))
		assert.ok(report.powerpoint.signals.includes('repair-dialog'), JSON.stringify(report.powerpoint.signals))
		assert.ok(report.powerpoint.repairSummary, JSON.stringify(report.powerpoint))
		assert.match(report.powerpoint.repairSummary, /found a problem with content/i)
		if (report.powerpoint.packageDiff) {
			assert.ok(
				report.powerpoint.packageDiff.changed.some(change => change.part.includes('ppt/slides/slide1.xml')),
				JSON.stringify(report.powerpoint.packageDiff),
			)
		}
	} finally {
		await rm(directory, { recursive: true, force: true })
	}
})

run('powerpoint: non-package file is rejected or offered repair then fails', async () => {
	const directory = await mkdtemp(join(tmpdir(), 'pptxgenjs-ppt-reject-'))
	try {
		const path = join(directory, 'reject-notzip.pptx')
		await writeFile(path, 'this is not a pptx')
		const report = await verifyPptx(path, { powerpoint: true })
		assert.ok(isPowerPointResult(report.powerpoint), JSON.stringify(report.powerpoint))
		assert.ok(report.powerpoint.verdict === 'repair' || report.powerpoint.verdict === 'reject', JSON.stringify(report.powerpoint))
		assert.equal(report.powerpoint.opened, false)
	} finally {
		await rm(directory, { recursive: true, force: true })
	}
})
