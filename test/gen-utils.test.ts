/**
 * PptxGenJS: Unit tests for utility methods
 * Run with: `bun run test`
 */
import { test } from 'node:test'
import assert from 'node:assert/strict'
import { readdirSync, readFileSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import {
	debugLog,
	importNodeBuiltin,
	isDebugEnabled,
	isNodeRuntime,
	getSmartParseNumber,
	getUuid,
	encodeXmlEntities,
	inch2Emu,
	valToPts,
	convertRotationDegrees,
	componentToHex,
	rgbToHex,
	createColorElement,
	createGlowElement,
	genXmlColorSelection,
	getNewRelId,
	correctShadowOptions,
	resolveThemeColors,
	DEF_THEME_COLORS,
	marginToEmu,
	inch2Emu,
	base64ToBytes,
	binaryStringToBase64,
	bytesToBase64,
	utf8ToBase64,
} from '../src/gen-utils'
import { PresLayout, PresSlide, ShadowProps } from '../src/core-interfaces'
import { DEF_FONT_COLOR } from '../src/core-enums'
import { genXmlLine } from '../src/xml/line'

// 10in x 7.5in layout expressed in EMU
const LAYOUT = { name: 'TEST', width: 9144000, height: 6858000 } as PresLayout

test('inch2Emu', () => {
	assert.equal(inch2Emu(1), 914400)
	assert.equal(inch2Emu(0.5), 457200)
	assert.equal(inch2Emu('2'), 1828800)
	assert.equal(inch2Emu('1in'), 914400)
	assert.equal(inch2Emu(200), 200, 'values > 100 are assumed to be EMU already')
	assert.equal(inch2Emu(-0.5), -457200, 'negative inches stay inches')
	assert.equal(inch2Emu(-200), -200, 'negative EMU is not re-multiplied')
})

test('valToPts', () => {
	assert.equal(valToPts(1), 12700)
	assert.equal(valToPts('2'), 25400)
	assert.equal(valToPts('not-a-number'), 0)
	assert.equal(valToPts(undefined as unknown as number), 0)
})

test('marginToEmu: dual-unit like table cells (>=1 points, <1 inches)', () => {
	assert.equal(marginToEmu(10), valToPts(10), '10 → points')
	assert.equal(marginToEmu(0.1), inch2Emu(0.1), '0.1 → inches')
	assert.equal(marginToEmu(0), 0)
})

test('convertRotationDegrees', () => {
	assert.equal(convertRotationDegrees(0), 0)
	assert.equal(convertRotationDegrees(90), 5400000)
	assert.equal(convertRotationDegrees(360), 21600000)
	assert.equal(convertRotationDegrees(361), 60000, 'wraps values over 360')
	assert.equal(convertRotationDegrees(undefined as unknown as number), 0)
})

test('componentToHex', () => {
	assert.equal(componentToHex(0), '00')
	assert.equal(componentToHex(15), '0f')
	assert.equal(componentToHex(255), 'ff')
})

test('rgbToHex', () => {
	assert.equal(rgbToHex(255, 0, 0), 'FF0000')
	assert.equal(rgbToHex(0, 255, 0), '00FF00')
	assert.equal(rgbToHex(0, 0, 0), '000000')
})

test('encodeXmlEntities', () => {
	assert.equal(encodeXmlEntities('a & b < c > d "e" \'f\''), 'a &amp; b &lt; c &gt; d &quot;e&quot; &apos;f&apos;')
	assert.equal(encodeXmlEntities(null as unknown as string), '')
	assert.equal(encodeXmlEntities(undefined as unknown as string), '')
})

test('getSmartParseNumber', () => {
	assert.equal(getSmartParseNumber(1, 'X', LAYOUT), 914400, 'small numbers are inches')
	assert.equal(getSmartParseNumber(-0.5, 'X', LAYOUT), -457200, 'negative inches stay inches')
	assert.equal(getSmartParseNumber(914400, 'X', LAYOUT), 914400, 'large numbers are already EMU')
	assert.equal(getSmartParseNumber(-914400, 'X', LAYOUT), -914400, 'negative EMU is not re-multiplied')
	assert.equal(getSmartParseNumber('50%', 'X', LAYOUT), 4572000, 'percent of width')
	assert.equal(getSmartParseNumber('50%', 'Y', LAYOUT), 3429000, 'percent of height')
	assert.equal(getSmartParseNumber('garbage', 'X', LAYOUT), 0)
})

test('getSmartParseNumber accepts unit-suffixed lengths', () => {
	assert.equal(getSmartParseNumber('2.54cm', 'X', LAYOUT), 914400, 'cm')
	assert.equal(getSmartParseNumber('25.4mm', 'Y', LAYOUT), 914400, 'mm')
	assert.equal(getSmartParseNumber('72pt', 'X', LAYOUT), 914400, 'pt')
	assert.equal(getSmartParseNumber('1in', 'X', LAYOUT), 914400, 'in')
	assert.equal(getSmartParseNumber('-1.27cm', 'X', LAYOUT), -457200, 'negative offsets are allowed')
	assert.equal(getSmartParseNumber('300cm', 'X', LAYOUT), 108000000, 'a stated unit bypasses the inches-vs-EMU heuristic')
	assert.equal(getSmartParseNumber('5px', 'X', LAYOUT), 0, 'unsupported units are not guessed at')
	assert.equal(getSmartParseNumber('.5in', 'X', LAYOUT), 457200, 'a leading decimal point')
	const digits = '0'.repeat(50000)
	const start = process.hrtime.bigint()
	assert.equal(getSmartParseNumber(`${digits}x`, 'X', LAYOUT), 0, 'a long non-matching digit string is rejected')
	assert.ok(Number(process.hrtime.bigint() - start) / 1e6 < 100, 'and rejected in linear time, not by backtracking')
})

test('inch2Emu accepts unit-suffixed lengths', () => {
	assert.equal(inch2Emu('2.54cm'), 914400, 'the inch-valued options (colW, rowH, inset, cell margin) share the parse')
	assert.equal(inch2Emu('300cm'), 108000000, 'no magnitude heuristic when the unit is stated')
})

test('getUuid', () => {
	assert.match(getUuid('xxxxxxxx'), /^[0-9a-f]{8}$/)
	assert.match(getUuid('y'), /^[89ab]$/, 'the "y" nibble is constrained per RFC4122')
	assert.notEqual(getUuid('xxxxxxxx-xxxx'), getUuid('xxxxxxxx-xxxx'), 'values are random')
})

test('base64 helpers are byte-equivalent to Node Buffer', () => {
	const bytes = Uint8Array.from([0, 1, 255, 128, 10])
	assert.equal(bytesToBase64(bytes), Buffer.from(bytes).toString('base64'))
	assert.deepEqual(base64ToBytes(Buffer.from(bytes).toString('base64')), bytes)
	assert.deepEqual(
		base64ToBytes('data:application/octet-stream;base64,' + Buffer.from(bytes).toString('base64')),
		bytes,
	)
	assert.equal(utf8ToBase64('Grüße'), Buffer.from('Grüße', 'utf8').toString('base64'))
	assert.equal(binaryStringToBase64('\x00\xff'), Buffer.from('\x00\xff', 'binary').toString('base64'))
})

test('createColorElement: hex', () => {
	assert.equal(createColorElement('FF0000'), '<a:srgbClr val="FF0000"/>')
	assert.equal(createColorElement('#ff0000'), '<a:srgbClr val="FF0000"/>', 'strips # and uppercases')
	assert.equal(createColorElement('FF0000', '<a:alpha val="50000"/>'), '<a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr>')
})

test('createColorElement: scheme color', () => {
	assert.equal(createColorElement('accent1'), '<a:schemeClr val="accent1"/>')
})

test('createColorElement: ModifiedThemeColor tint/shade (percent ×1000)', () => {
	assert.equal(
		createColorElement({ baseColor: 'accent1', tint: 40 }),
		'<a:schemeClr val="accent1"><a:tint val="40000"/></a:schemeClr>'
	)
	assert.equal(
		createColorElement({ baseColor: 'FF0000', shade: 50, alpha: 80 }),
		'<a:srgbClr val="FF0000"><a:alpha val="80000"/><a:shade val="50000"/></a:srgbClr>'
	)
})

test('createColorElement: ModifiedThemeColor hue uses deg×60000 (not ×1000)', () => {
	assert.equal(
		createColorElement({ baseColor: 'accent2', hue: 90 }),
		'<a:schemeClr val="accent2"><a:hue val="5400000"/></a:schemeClr>'
	)
	assert.equal(
		createColorElement({ baseColor: 'accent2', hueOff: 30 }),
		'<a:schemeClr val="accent2"><a:hueOff val="1800000"/></a:schemeClr>'
	)
	// hueMod is a percent, unlike hue/hueOff
	assert.equal(
		createColorElement({ baseColor: 'accent2', hueMod: 50 }),
		'<a:schemeClr val="accent2"><a:hueMod val="50000"/></a:schemeClr>'
	)
})

test('createColorElement: ModifiedThemeColor allows lumMod > 100', () => {
	assert.equal(
		createColorElement({ baseColor: 'accent1', lumMod: 110 }),
		'<a:schemeClr val="accent1"><a:lumMod val="110000"/></a:schemeClr>'
	)
})

test('genXmlColorSelection: bare ModifiedThemeColor (not mistaken for ShapeFillProps)', () => {
	assert.equal(
		genXmlColorSelection({ baseColor: 'accent1', tint: 25 }),
		'<a:solidFill><a:schemeClr val="accent1"><a:tint val="25000"/></a:schemeClr></a:solidFill>'
	)
	assert.equal(
		genXmlColorSelection({ type: 'solid', color: { baseColor: 'bg1', shade: 10 } }),
		'<a:solidFill><a:schemeClr val="bg1"><a:shade val="10000"/></a:schemeClr></a:solidFill>'
	)
})

test('resolveThemeColors: requires exactly 12 hex colors', () => {
	assert.deepEqual(resolveThemeColors(undefined), [...DEF_THEME_COLORS])
	const orig = console.warn
	console.warn = () => {}
	try {
		assert.deepEqual(resolveThemeColors({ themeColors: ['FF0000'] }), [...DEF_THEME_COLORS], 'short array falls back')
		assert.deepEqual(resolveThemeColors({ hlinkColor: 'not-hex' }), [...DEF_THEME_COLORS], 'invalid hlinkColor is ignored')
	} finally {
		console.warn = orig
	}
	const custom = [
		'111111', '222222', '333333', '444444',
		'555555', '666666', '777777', '888888', '999999', 'AAAAAA',
		'BBBBBB', 'CCCCCC',
	]
	assert.deepEqual(resolveThemeColors({ themeColors: custom }), custom)
	assert.deepEqual(
		resolveThemeColors({ themeColors: custom.map(c => `#${c.toLowerCase()}`) }),
		custom,
		'strips # and uppercases'
	)
	const withHlink = [...DEF_THEME_COLORS]
	withHlink[10] = 'FF0000'
	assert.deepEqual(resolveThemeColors({ hlinkColor: '#ff0000' }), withHlink, 'hlinkColor overrides scheme hlink')
	// eslint-disable-next-line @typescript-eslint/no-explicit-any
	assert.deepEqual(resolveThemeColors({ hlinkColor: 0 as any }), [...DEF_THEME_COLORS], 'non-string hlinkColor is ignored')
})

test('createColorElement: invalid falls back to default font color', () => {
	const orig = console.warn
	console.warn = () => {} // silence the expected warning
	try {
		assert.equal(createColorElement('NOPE'), '<a:srgbClr val="000000"/>')
	} finally {
		console.warn = orig
	}
})

test('createGlowElement', () => {
	assert.equal(
		createGlowElement({ size: 8, color: 'FFFFFF', opacity: 0.75 }, { size: 8, color: 'FFFFFF', opacity: 0.75 }),
		'<a:glow rad="101600"><a:srgbClr val="FFFFFF"><a:alpha val="75000"/></a:srgbClr></a:glow>'
	)
})

test('genXmlColorSelection', () => {
	assert.equal(genXmlColorSelection('FF0000'), '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>')
	assert.equal(
		genXmlColorSelection({ type: 'solid', color: 'FF0000', transparency: 50 }),
		'<a:solidFill><a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr></a:solidFill>'
	)
	// Gradient with too few stops degrades to a solid fill (PowerPoint repairs files with <2-stop gsLst)
	const orig = console.warn
	console.warn = () => {}
	try {
		assert.equal(genXmlColorSelection({ type: 'gradient', color: 'FF0000' }), '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>')
	} finally {
		console.warn = orig
	}
	// Valid 2-stop linear gradient
	assert.equal(
		genXmlColorSelection({
			type: 'gradient',
			gradient: {
				type: 'linear',
				angle: 90,
				stops: [
					{ color: 'FF0000', pos: 0 },
					{ color: '0000FF', pos: 100, transparency: 50 },
				],
			},
		}),
		'<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="FF0000"/></a:gs><a:gs pos="100000"><a:srgbClr val="0000FF"><a:alpha val="50000"/></a:srgbClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill>'
	)
	// sambauers/gradients flat API (`type:'linearGradient'`, `position`, tileRect, flip)
	assert.equal(
		genXmlColorSelection({
			type: 'linearGradient',
			angle: 45,
			scaled: true,
			flip: 'x',
			tileRect: { t: 10, l: 0 },
			stops: [
				{ position: 0, color: '000000', transparency: 10 },
				{ position: 100, color: '333333', transparency: 50 },
			],
		}),
		'<a:gradFill rotWithShape="1" flip="x"><a:gsLst><a:gs pos="0"><a:srgbClr val="000000"><a:alpha val="90000"/></a:srgbClr></a:gs><a:gs pos="100000"><a:srgbClr val="333333"><a:alpha val="50000"/></a:srgbClr></a:gs></a:gsLst><a:lin ang="2700000" scaled="1"/><a:tileRect t="10000" l="0"/></a:gradFill>'
	)
	// angle 0 still emits `<a:lin>` (upstream sambauers skipped falsy angle)
	assert.ok(
		genXmlColorSelection({
			type: 'linearGradient',
			stops: [
				{ position: 0, color: 'FF0000' },
				{ position: 100, color: '0000FF' },
			],
		}).includes('<a:lin ang="0" scaled="0"/>')
	)
	// Radial gradient
	assert.equal(
		genXmlColorSelection({
			type: 'gradient',
			gradient: {
				type: 'radial',
				stops: [
					{ color: 'FFFFFF', pos: 0 },
					{ color: '000000', pos: 100 },
				],
			},
		}),
		'<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:srgbClr val="FFFFFF"/></a:gs><a:gs pos="100000"><a:srgbClr val="000000"/></a:gs></a:gsLst><a:path path="circle"></a:path></a:gradFill>'
	)
	// Pattern fill (hakrueger/pattern)
	assert.equal(
		genXmlColorSelection({
			type: 'pattern',
			pattern: { prst: 'ltHorz', color: 'FF0000', bgColor: '00FF00' },
		}),
		'<a:pattFill prst="ltHorz"><a:bgClr><a:srgbClr val="00FF00"/></a:bgClr><a:fgClr><a:srgbClr val="FF0000"/></a:fgClr></a:pattFill>'
	)
	assert.equal(
		genXmlColorSelection({ type: 'pattern', color: '112233' }),
		'<a:pattFill prst="cross"><a:bgClr><a:srgbClr val="FFFFFF"/></a:bgClr><a:fgClr><a:srgbClr val="112233"/></a:fgClr></a:pattFill>'
	)
	// Gradient stops accept ModifiedThemeColor; transforms are children of EG_ColorChoice, not of a:gs
	assert.equal(
		genXmlColorSelection({
			type: 'gradient',
			gradient: {
				type: 'linear',
				angle: 0,
				stops: [
					{ color: { baseColor: 'accent1', tint: 40 }, pos: 0 },
					{ color: { baseColor: 'FF0000', shade: 50, lumMod: 110 }, pos: 100 },
				],
			},
		}),
		'<a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="accent1"><a:tint val="40000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:srgbClr val="FF0000"><a:lumMod val="110000"/><a:shade val="50000"/></a:srgbClr></a:gs></a:gsLst><a:lin ang="0" scaled="0"/></a:gradFill>'
	)
	assert.equal(genXmlColorSelection({ type: 'none' }), '<a:noFill/>', 'type:none is DrawingML noFill')
	// gs@pos is ST_PositiveFixedPercentage (0–100000); API percents outside 0–100 are clamped
	const clampedStops = genXmlColorSelection({
		type: 'gradient',
		gradient: {
			stops: [
				{ color: 'FF0000', pos: -10 },
				{ color: '0000FF', pos: 150 },
			],
		},
	})
	assert.match(clampedStops, /<a:gs pos="0">/)
	assert.match(clampedStops, /<a:gs pos="100000">/)
	assert.match(clampedStops, /<a:gsLst>.*<a:gs .*<a:gs /)
})

test('getNewRelId', () => {
	const slide = { _rels: [1, 2], _relsChart: [1], _relsMedia: [] } as unknown as PresSlide
	assert.equal(getNewRelId(slide), 4, 'sum of all rel arrays + 1')
})

test('correctShadowOptions', () => {
	assert.equal(correctShadowOptions(undefined as unknown as ShadowProps), undefined)
	assert.equal(correctShadowOptions('nope' as unknown as ShadowProps), undefined)
	assert.equal(correctShadowOptions({ type: 'bogus' } as unknown as ShadowProps)?.type, 'outer', 'invalid type corrected')
	assert.equal(correctShadowOptions({ type: 'outer', angle: 400 } as ShadowProps)?.angle, 270, 'out-of-range angle corrected')
	assert.equal(correctShadowOptions({ type: 'outer', opacity: 2 } as ShadowProps)?.opacity, 0.75, 'out-of-range opacity corrected')
	assert.equal(correctShadowOptions({ type: 'outer', color: '#FF0000' } as ShadowProps)?.color, 'FF0000', 'strips leading #')
})

test('debugLog: silent unless PPTXGENJS_DEBUG or NODE_DEBUG is set', () => {
	const { PPTXGENJS_DEBUG, NODE_DEBUG } = process.env
	const calls: unknown[][] = []
	const orig = console.debug
	console.debug = (...args: unknown[]) => calls.push(args)
	try {
		delete process.env.PPTXGENJS_DEBUG
		delete process.env.NODE_DEBUG
		assert.equal(isDebugEnabled(), false)
		debugLog('quiet')
		assert.equal(calls.length, 0)

		process.env.NODE_DEBUG = 'http,pptxgenjs'
		assert.equal(isDebugEnabled(), true)
		debugLog('loud')
		assert.deepEqual(calls, [['[pptxgenjs]', 'loud']])

		process.env.NODE_DEBUG = 'http'
		assert.equal(isDebugEnabled(), false, 'NODE_DEBUG matched on a partial section name')

		process.env.PPTXGENJS_DEBUG = '1'
		assert.equal(isDebugEnabled(), true)
	} finally {
		console.debug = orig
		if (PPTXGENJS_DEBUG === undefined) delete process.env.PPTXGENJS_DEBUG; else process.env.PPTXGENJS_DEBUG = PPTXGENJS_DEBUG
		if (NODE_DEBUG === undefined) delete process.env.NODE_DEBUG; else process.env.NODE_DEBUG = NODE_DEBUG
	}
})

test('isNodeRuntime: true in this Node test process', () => {
	assert.equal(isNodeRuntime(), true)
})

test('importNodeBuiltin: loads fs without a static node: specifier', async () => {
	const fs = await importNodeBuiltin<typeof import('node:fs')>('fs')
	assert.equal(typeof fs.readFileSync, 'function')
})

test('runtime sources do not statically import Node builtins', () => {
	const root = join(dirname(fileURLToPath(import.meta.url)), '..')
	for (const rel of ['src/gen-utils.ts', 'src/gen-media.ts', 'src/pptxgen.ts']) {
		const runtime = readFileSync(join(root, rel), 'utf8')
			.replace(/\/\*[\s\S]*?\*\//g, '')
			.replace(/\/\/.*$/gm, '')
			.replace(/typeof\s+import\s*\(\s*['"]node:[^'"]+['"]\s*\)/g, '')
		assert.doesNotMatch(
			runtime,
			/import\s*\(\s*(?:\/\*[\s\S]*?\*\/\s*)*['"]node:/,
			`${rel} still has a statically analyzable node: import`
		)
	}
})

function listBrowserSrcFiles (dir: string): string[] {
	const out: string[] = []
	for (const ent of readdirSync(dir, { withFileTypes: true })) {
		if (ent.name === 'vendor') continue
		const p = join(dir, ent.name)
		if (ent.isDirectory()) out.push(...listBrowserSrcFiles(p))
		else if (ent.name.endsWith('.ts') || ent.name.endsWith('.js')) out.push(p)
	}
	return out
}

test('createColorElement: ColorProps object forms', () => {
	assert.equal(createColorElement({ hex: 'FF0000' }), '<a:srgbClr val="FF0000"/>')
	assert.equal(createColorElement({ scheme: 'accent1' }), '<a:schemeClr val="accent1"/>')
	assert.equal(createColorElement({ scheme: 'hlink' }), '<a:schemeClr val="hlink"/>')
	assert.equal(createColorElement({ scheme: 'folHlink' }), '<a:schemeClr val="folHlink"/>')
	assert.equal(createColorElement({ scheme: 'phClr' }), '<a:schemeClr val="phClr"/>')
	assert.equal(createColorElement({ scheme: 'dk1' }), '<a:schemeClr val="dk1"/>')
	assert.equal(createColorElement({ system: 'windowText' }), '<a:sysClr val="windowText"/>')
	assert.equal(createColorElement({ system: 'windowText', lastColor: '#000000' }), '<a:sysClr val="windowText" lastClr="000000"/>')
	assert.equal(createColorElement({ preset: 'cornflowerBlue' }), '<a:prstClr val="cornflowerBlue"/>')
	assert.equal(createColorElement({ hsl: { hue: 210, sat: 80, lum: 50 } }), '<a:hslClr hue="12600000" sat="80000" lum="50000"/>')
	assert.equal(createColorElement({ scrgb: { r: 100, g: 50, b: 0 } }), '<a:scrgbClr r="100000" g="50000" b="0"/>')
	// string form of the extra scheme slots
	assert.equal(createColorElement('dk1'), '<a:schemeClr val="dk1"/>')
	assert.equal(createColorElement('hlink'), '<a:schemeClr val="hlink"/>')
})

test('createColorElement: ColorProps transforms and their unit ranges', () => {
	assert.equal(createColorElement({ hex: 'FF0000', alpha: 50 }), '<a:srgbClr val="FF0000"><a:alpha val="50000"/></a:srgbClr>')
	assert.equal(createColorElement({ scheme: 'accent1', lumMod: 60, lumOff: 40 }), '<a:schemeClr val="accent1"><a:lumMod val="60000"/><a:lumOff val="40000"/></a:schemeClr>')
	assert.equal(createColorElement({ scheme: 'accent1', satMod: 170 }), '<a:schemeClr val="accent1"><a:satMod val="170000"/></a:schemeClr>')
	assert.equal(createColorElement({ hex: '00FF00', lumOff: -25, satOff: 200 }), '<a:srgbClr val="00FF00"><a:satOff val="100000"/><a:lumOff val="-25000"/></a:srgbClr>')
	assert.equal(createColorElement({ hex: '00FF00', hueOff: -30 }), '<a:srgbClr val="00FF00"><a:hueOff val="-1800000"/></a:srgbClr>')
	assert.equal(createColorElement({ hex: '00FF00', inverse: true, grayscale: true, complement: true, gamma: true, inverseGamma: true }),
		'<a:srgbClr val="00FF00"><a:inv/><a:gray/><a:comp/><a:gamma/><a:invGamma/></a:srgbClr>')
	assert.equal(createColorElement({ hex: '00FF00', alpha: 150, tint: -20 }), '<a:srgbClr val="00FF00"><a:tint val="0"/><a:alpha val="100000"/></a:srgbClr>')
})

test('createColorElement: unknown ColorProps names fall back rather than emit invalid enums', () => {
	const orig = console.warn
	const warnings: string[] = []
	console.warn = (msg: string) => warnings.push(String(msg))
	try {
		assert.equal(createColorElement({ scheme: 'nope' as unknown as 'accent1' }), `<a:srgbClr val="${DEF_FONT_COLOR}"/>`)
		assert.equal(createColorElement({ system: 'chartreuse' }), `<a:srgbClr val="${DEF_FONT_COLOR}"/>`)
		assert.equal(createColorElement({ preset: 'ultraviolet' }), `<a:srgbClr val="${DEF_FONT_COLOR}"/>`)
		assert.equal(createColorElement({ hex: 'ZZZ' }), `<a:srgbClr val="${DEF_FONT_COLOR}"/>`)
		assert.equal(createColorElement({ system: 'windowText', lastColor: 'nope' }), '<a:sysClr val="windowText"/>')
	} finally {
		console.warn = orig
	}
	assert.equal(warnings.length, 4, 'each rejected name warns once')
	assert.ok(warnings.some(w => w.includes('is not a theme color slot')))
	assert.ok(warnings.some(w => w.includes('is not a system color')))
	assert.ok(warnings.some(w => w.includes('is not a preset color name')))
})

test('createColorElement: ModifiedThemeColor still works beside ColorProps', () => {
	assert.equal(
		createColorElement({ baseColor: 'accent1', tint: 40 }),
		'<a:schemeClr val="accent1"><a:tint val="40000"/></a:schemeClr>'
	)
})

test('genXmlColorSelection: a color object is a solid fill, a fill object is not', () => {
	assert.equal(genXmlColorSelection({ scheme: 'accent1', lumMod: 75 }), '<a:solidFill><a:schemeClr val="accent1"><a:lumMod val="75000"/></a:schemeClr></a:solidFill>')
	assert.equal(genXmlColorSelection({ preset: 'gold' }), '<a:solidFill><a:prstClr val="gold"/></a:solidFill>')
	assert.equal(genXmlColorSelection({ type: 'solid', color: 'FF0000' }), '<a:solidFill><a:srgbClr val="FF0000"/></a:solidFill>')
	assert.equal(genXmlColorSelection({ type: 'solid', color: { scheme: 'accent2', shade: 50 } }), '<a:solidFill><a:schemeClr val="accent2"><a:shade val="50000"/></a:schemeClr></a:solidFill>')
	assert.equal(genXmlColorSelection({ baseColor: 'accent1', tint: 25 }), '<a:solidFill><a:schemeClr val="accent1"><a:tint val="25000"/></a:schemeClr></a:solidFill>')
})

test('genXmlLine: preset dash, cap, compound, joins, and arrow sizing', () => {
	assert.equal(genXmlLine({ width: 2, color: 'FF0000', dashType: 'dash' }),
		'<a:ln w="25400"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="dash"/></a:ln>')
	assert.equal(genXmlLine({ width: 2, color: 'FF0000', cap: 'rnd' }),
		'<a:ln w="25400" cap="rnd"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln>')
	assert.equal(genXmlLine({ width: 3, color: '000000', compound: 'thickThin' }),
		'<a:ln w="38100" cmpd="thickThin"><a:solidFill><a:srgbClr val="000000"/></a:solidFill></a:ln>')
	assert.equal(genXmlLine({ color: '000000', join: 'round' }), '<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:round/></a:ln>')
	assert.equal(genXmlLine({ color: '000000', join: 'bevel' }), '<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:bevel/></a:ln>')
	assert.equal(genXmlLine({ color: '000000', join: 'miter', miterLimit: 400 }), '<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:miter lim="400000"/></a:ln>')
	assert.equal(genXmlLine({ color: '000000', join: 'miter' }), '<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:miter lim="800000"/></a:ln>')
	assert.equal(genXmlLine({ color: '000000', beginArrowType: 'arrow', beginArrowSize: { width: 'lg', length: 'lg' }, endArrowType: 'triangle', endArrowSize: { width: 'sm' } }),
		'<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:headEnd type="arrow" w="lg" len="lg"/><a:tailEnd type="triangle" w="sm"/></a:ln>')
	assert.equal(genXmlLine({ color: '000000', beginArrowType: 'arrow', beginArrowSize: 'med' }),
		'<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:headEnd type="arrow" w="med" len="med"/></a:ln>')
})

test('genXmlLine: a custom dash replaces the preset one', () => {
	assert.equal(genXmlLine({ color: '000000', dashType: 'dash', custDash: [{ dash: 400, space: 300 }, { dash: 100, space: 300 }] }),
		'<a:ln><a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:custDash><a:ds d="400000" sp="300000"/><a:ds d="100000" sp="300000"/></a:custDash></a:ln>')
})

test('genXmlLine: the same content model serves a:uLn', () => {
	assert.equal(genXmlLine({ width: 1, color: 'FF0000', dashType: 'solid' }, 'a:uLn'),
		'<a:uLn w="12700"><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill><a:prstDash val="solid"/></a:uLn>')
})

test('genXmlLine: invalid line options are dropped with a warning', () => {
	const orig = console.warn
	const warnings: string[] = []
	console.warn = (msg: string) => warnings.push(String(msg))
	let compound = ''
	let dash = ''
	let join = ''
	try {
		compound = genXmlLine({ color: '000000', compound: 'quad' as 'sng' })
		dash = genXmlLine({ color: '000000', custDash: [{ dash: -1, space: 100 }, { dash: 200, space: 100 }] })
		join = genXmlLine({ color: '000000', join: 'sharp' as 'miter' })
	} finally {
		console.warn = orig
	}
	assert.ok(warnings.some(w => w.includes('`compound` must be one of')), 'bad compound must warn')
	assert.ok(warnings.some(w => w.includes('each `custDash` stop needs')), 'bad dash stop must warn')
	assert.ok(warnings.some(w => w.includes('`join` must be')), 'bad join must warn')
	assert.doesNotMatch(compound, /cmpd=/, 'an invalid compound must not be emitted')
	assert.match(dash, /<a:custDash><a:ds d="200000" sp="100000"\/><\/a:custDash>/, 'valid dash stops must survive')
	assert.doesNotMatch(join, /a:round|a:bevel|a:miter/, 'an invalid join must not be emitted')
	assert.doesNotMatch(compound + dash + join, /NaN/, 'no NaN attributes')
})

test('browser-bundled helpers do not reference Buffer', () => {
	const root = join(dirname(fileURLToPath(import.meta.url)), '..')
	const srcRoot = join(root, 'src')
	for (const file of listBrowserSrcFiles(srcRoot)) {
		const runtime = readFileSync(file, 'utf8')
			.replace(/\/\*[\s\S]*?\*\//g, '')
			.replace(/\/\/.*$/gm, '')
		assert.doesNotMatch(
			runtime,
			/\bBuffer\b/,
			`${file.slice(srcRoot.length + 1)} still references Buffer in browser-bundled code`
		)
	}
})
