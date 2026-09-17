/**
 * Builds small PPTX files that PowerPoint has been observed to offer to repair.
 * Only mutations that produced a Repair dialog on a real Office 16 sidecar run
 * stay in the set; silent-ok and COM-crash cases were dropped.
 */
import { mkdir, writeFile } from 'node:fs/promises'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { JSZip } from '@node-projects/jszip'
import pptxgen from '../src/pptxgen'

export type RepairExpect = 'ok' | 'repair' | 'reject'

export type RepairFixture = {
	id: string
	title: string
	expect: RepairExpect
	mutate?: (zip: JSZip) => Promise<void>
}

const here = dirname(fileURLToPath(import.meta.url))
export const REPAIR_FIXTURE_DIR = join(here, 'fixtures', 'repair')

function mustReplace (xml: string, pattern: RegExp | string, replacement: string, label: string): string {
	if (typeof pattern === 'string') {
		if (!xml.includes(pattern)) throw new Error(`${label}: missing ${JSON.stringify(pattern)}`)
		return xml.replace(pattern, replacement)
	}
	if (!pattern.test(xml)) throw new Error(`${label}: missing ${pattern}`)
	return xml.replace(pattern, replacement)
}

/** First `p:cNvPr` after the reserved `p:nvGrpSpPr` id="1". */
function firstObjectCnvPrOpen (xml: string, label: string): string {
	const match = [...xml.matchAll(/<p:cNvPr id="(\d+)"/g)].find(item => item[1] !== '1')
	if (!match) throw new Error(`${label}: no object p:cNvPr`)
	return match[0]
}

async function rewrite (zip: JSZip, part: string, rewriteXml: (xml: string) => string): Promise<void> {
	const file = zip.file(part)
	if (!file) throw new Error(`missing part ${part}`)
	zip.file(part, rewriteXml(await file.async('string')), { createFolders: false })
}

export const REPAIR_FIXTURES: RepairFixture[] = [
	{
		id: 'ok-baseline',
		title: 'Unmutated control deck',
		expect: 'ok',
	},
	{
		id: 'invalid-cnvid',
		title: 'Non-numeric p:cNvPr@id',
		expect: 'repair',
		mutate: zip => rewrite(zip, 'ppt/slides/slide1.xml', xml =>
			mustReplace(xml, firstObjectCnvPrOpen(xml, 'invalid-cnvid'), '<p:cNvPr id="abc"', 'invalid-cnvid')),
	},
	{
		id: 'empty-txbody',
		title: 'Empty p:txBody with no paragraph',
		expect: 'repair',
		mutate: zip => rewrite(zip, 'ppt/slides/slide1.xml', xml =>
			mustReplace(xml, /<p:txBody>[\s\S]*?<\/p:txBody>/, '<p:txBody/>', 'empty-txbody')),
	},
	{
		id: 'nan-cell-margin',
		title: 'Table cell marL="NaN"',
		expect: 'repair',
		mutate: zip => rewrite(zip, 'ppt/slides/slide1.xml', xml =>
			mustReplace(xml, /marL="\d+"/, 'marL="NaN"', 'nan-cell-margin')),
	},
	{
		id: 'nan-shape-cx',
		title: 'Shape a:ext cx="NaN"',
		expect: 'repair',
		mutate: zip => rewrite(zip, 'ppt/slides/slide1.xml', xml =>
			mustReplace(xml, /<a:ext cx="\d+" cy="\d+"\/>/, '<a:ext cx="NaN" cy="914400"/>', 'nan-shape-cx')),
	},
	{
		id: 'empty-cnvid',
		title: 'Empty p:cNvPr@id',
		expect: 'repair',
		mutate: zip => rewrite(zip, 'ppt/slides/slide1.xml', xml =>
			mustReplace(xml, firstObjectCnvPrOpen(xml, 'empty-cnvid'), '<p:cNvPr id=""', 'empty-cnvid')),
	},
	{
		id: 'invalid-preset',
		title: 'Unknown a:prstGeom@prst',
		expect: 'repair',
		mutate: zip => rewrite(zip, 'ppt/slides/slide1.xml', xml =>
			mustReplace(xml, 'prst="rect"', 'prst="notashape"', 'invalid-preset')),
	},
	{
		id: 'empty-table-cell',
		title: 'Table cell txBody with no paragraphs',
		expect: 'repair',
		mutate: zip => rewrite(zip, 'ppt/slides/slide1.xml', xml =>
			mustReplace(
				xml,
				/<a:tc>(<a:txBody>[\s\S]*?<\/a:txBody>)/,
				'<a:tc><a:txBody><a:bodyPr/><a:lstStyle/></a:txBody>',
				'empty-table-cell',
			)),
	},
]

async function basePackage (): Promise<JSZip> {
	const pptx = new pptxgen()
	const slide = pptx.addSlide()
	slide.addNotes('repair fixture notes')
	slide.background = { color: 'F5F5F5' }
	slide.addText('Repair fixture', { x: 0.5, y: 0.3, w: 8, h: 0.5, fontSize: 20, bold: true })
	slide.addShape(pptx.ShapeType.rect, { x: 0.5, y: 1.0, w: 2, h: 1, fill: { color: 'FF0000' } })
	slide.addTable([['A', 'B'], ['1', '2']], { x: 0.5, y: 2.3, w: 4 })
	slide.addChart(pptx.ChartType.bar, [{ name: 'Sales', labels: ['Q1', 'Q2'], values: [10, 20] }], { x: 5.2, y: 1.0, w: 5, h: 3 })
	return JSZip.loadAsync((await pptx.write({ outputType: 'nodebuffer' })) as Buffer)
}

export function fixturePath (id: string): string {
	return join(REPAIR_FIXTURE_DIR, `${id}.pptx`)
}

export async function buildRepairFixtures (directory = REPAIR_FIXTURE_DIR): Promise<string[]> {
	await mkdir(directory, { recursive: true })
	const written: string[] = []
	for (const fixture of REPAIR_FIXTURES) {
		const zip = await basePackage()
		if (fixture.mutate) await fixture.mutate(zip)
		const path = join(directory, `${fixture.id}.pptx`)
		await writeFile(path, await zip.generateAsync({ type: 'nodebuffer' }))
		written.push(path)
	}
	const manifest = {
		directory: 'test/fixtures/repair',
		fixtures: REPAIR_FIXTURES.map(fixture => ({
			id: fixture.id,
			file: `${fixture.id}.pptx`,
			title: fixture.title,
			expect: fixture.expect,
		})),
	}
	await writeFile(join(directory, 'manifest.json'), JSON.stringify(manifest, null, 2) + '\n')
	return written
}

async function main (): Promise<void> {
	const paths = await buildRepairFixtures()
	console.log(`wrote ${paths.length} fixtures to ${REPAIR_FIXTURE_DIR}`)
	for (const fixture of REPAIR_FIXTURES) {
		console.log(`  ${fixture.id}.pptx  expect=${fixture.expect}  ${fixture.title}`)
	}
}

const invokedDirectly = process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]
if (invokedDirectly) void main()
