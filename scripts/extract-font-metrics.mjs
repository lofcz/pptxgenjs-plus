// Extracts advance widths from TrueType fonts into the table `packages/pptxgenjs-std/src/text/metrics-data.ts`
// consumes. Dev tooling: not shipped, not run by any lifecycle hook - run it by hand when a font
// is added, then commit the generated file.
//
//   node scripts/extract-font-metrics.mjs out.ts "Calibri=/path/Carlito-Regular.ttf" ...
//
// The key of each entry is the font name *as used in slide text options*, so a metric-compatible
// substitute can stand in for the font it clones (Carlito for Calibri, Arimo for Arial, ...).
//
// Advance widths only: no kerning (GPOS) and no ligatures, so measured runs are marginally wide
// where a real renderer would pull pairs together. That is the right direction to err for
// overflow checks.
import { readFileSync, writeFileSync } from 'node:fs'

/** Codepoints worth carrying: ASCII printable plus Latin-1 letters and common punctuation */
const COVERED = [
	[0x20, 0x7e],
	[0xa0, 0xff],
	[0x2013, 0x2014], // en/em dash
	[0x2018, 0x201d], // curly quotes
	[0x2026, 0x2026], // ellipsis
	[0x20ac, 0x20ac], // euro
]

const u16 = (buf, at) => buf.readUInt16BE(at)
const i16 = (buf, at) => buf.readInt16BE(at)
const u32 = (buf, at) => buf.readUInt32BE(at)

/** Table directory: tag -> byte offset */
function readTables (buf) {
	const tables = {}
	const count = u16(buf, 4)
	for (let i = 0; i < count; i++) {
		const at = 12 + i * 16
		tables[buf.toString('ascii', at, at + 4)] = u32(buf, at + 8)
	}
	return tables
}

/** Character-to-glyph map. Format 4 only - every font shipping Latin text has one. */
function readCmap (buf, cmapAt) {
	let subtableAt = 0
	for (let i = 0; i < u16(buf, cmapAt + 2); i++) {
		const rec = cmapAt + 4 + i * 8
		const platform = u16(buf, rec)
		const encoding = u16(buf, rec + 2)
		// (3,1) Windows BMP is the canonical Unicode subtable; (0,x) Unicode is an acceptable fallback
		if ((platform === 3 && encoding === 1) || platform === 0) subtableAt = cmapAt + u32(buf, rec + 4)
		if (platform === 3 && encoding === 1) break
	}
	if (!subtableAt) throw new Error('cmap: no Unicode subtable')
	if (u16(buf, subtableAt) !== 4) throw new Error(`cmap: unsupported format ${u16(buf, subtableAt)}`)

	const segCount = u16(buf, subtableAt + 6) / 2
	const ends = subtableAt + 14
	const starts = ends + segCount * 2 + 2
	const deltas = starts + segCount * 2
	const ranges = deltas + segCount * 2

	return function glyphFor (code) {
		for (let s = 0; s < segCount; s++) {
			const end = u16(buf, ends + s * 2)
			if (code > end) continue
			const start = u16(buf, starts + s * 2)
			if (code < start) return 0
			const rangeOffset = u16(buf, ranges + s * 2)
			if (rangeOffset === 0) return (code + i16(buf, deltas + s * 2)) & 0xffff
			const glyphAt = ranges + s * 2 + rangeOffset + (code - start) * 2
			const glyph = u16(buf, glyphAt)
			return glyph === 0 ? 0 : (glyph + i16(buf, deltas + s * 2)) & 0xffff
		}
		return 0
	}
}

/** Advance width in font units for a glyph id; the last hmtx entry covers every glyph after it */
function widthReader (buf, hmtxAt, numHMetrics) {
	return glyph => u16(buf, hmtxAt + Math.min(glyph, numHMetrics - 1) * 4)
}

function extract (file) {
	const buf = readFileSync(file)
	const tables = readTables(buf)
	for (const tag of ['head', 'hhea', 'hmtx', 'cmap']) {
		if (!tables[tag]) throw new Error(`${file}: missing ${tag} table`)
	}

	const unitsPerEm = u16(buf, tables.head + 18)
	if (unitsPerEm < 16 || unitsPerEm > 16384) throw new Error(`${file}: implausible unitsPerEm ${unitsPerEm}`)

	const glyphFor = readCmap(buf, tables.cmap)
	const widthFor = widthReader(buf, tables.hmtx, u16(buf, tables.hhea + 34))
	/** Font units -> per-mille of em, so the table is unit-agnostic and integer */
	const perMille = units => Math.round((units / unitsPerEm) * 1000)

	const widths = {}
	let missing = 0
	for (const [from, to] of COVERED) {
		for (let code = from; code <= to; code++) {
			const glyph = glyphFor(code)
			if (glyph === 0) { missing++; continue }
			widths[code] = perMille(widthFor(glyph))
		}
	}
	if (Object.keys(widths).length < 90) throw new Error(`${file}: only ${Object.keys(widths).length} glyphs mapped`)

	const space = widths[0x20]
	if (!space) throw new Error(`${file}: no space glyph`)

	return {
		widths,
		missing,
		// hhea ascender/descender/lineGap: the em-relative box a line of this font occupies
		ascent: perMille(i16(buf, tables.hhea + 4)),
		descent: perMille(i16(buf, tables.hhea + 6)),
		lineGap: perMille(i16(buf, tables.hhea + 8)),
		fallbackWidth: space,
	}
}

/** Runs of equal width collapse to `start:count:width`, which is most of this table */
function encode (widths) {
	const codes = Object.keys(widths).map(Number).sort((a, b) => a - b)
	const runs = []
	for (const code of codes) {
		const last = runs[runs.length - 1]
		if (last && last.width === widths[code] && last.start + last.count === code) last.count++
		else runs.push({ start: code, count: 1, width: widths[code] })
	}
	return runs.map(run => `${run.start}:${run.count}:${run.width}`).join(',')
}

const [outFile, ...specs] = process.argv.slice(2)
if (!outFile || specs.length === 0) {
	console.error('usage: node scripts/extract-font-metrics.mjs <out.ts> "<Font Name>=<file.ttf>" ...')
	process.exit(1)
}

const entries = specs.map(spec => {
	const split = spec.indexOf('=')
	if (split < 1) throw new Error(`bad spec "${spec}" - expected "<Font Name>=<file.ttf>"`)
	const name = spec.slice(0, split)
	const font = extract(spec.slice(split + 1))
	console.error(`${name}: ${Object.keys(font.widths).length} glyphs, ${font.missing} uncovered, ascent ${font.ascent}`)
	return { name, font }
})

const body = entries
	.map(({ name, font }) => `\t'${name.toLowerCase()}': {
		ascent: ${font.ascent},
		descent: ${font.descent},
		lineGap: ${font.lineGap},
		fallbackWidth: ${font.fallbackWidth},
		widths: '${encode(font.widths)}',
	},`)
	.join('\n')

writeFileSync(outFile, `/**
 * GENERATED by scripts/extract-font-metrics.mjs - do not edit by hand.
 *
 * Advance widths in per-mille of the em, run-length encoded as \`codepoint:count:width\`.
 * No kerning or ligature data, so a measured run is marginally wider than a real renderer's -
 * the safe direction for overflow checks.
 *
 * Carlito (c) 2013 The Carlito Project Authors, SIL Open Font License 1.1. Carlito is
 * metric-compatible with Calibri, so it stands in for it here; no glyph outlines are included.
 */
export interface RawFontMetrics {
	ascent: number
	descent: number
	lineGap: number
	fallbackWidth: number
	widths: string
}

export const RAW_METRICS: Record<string, RawFontMetrics> = {
${body}
}
`)
console.error(`wrote ${outFile}`)
