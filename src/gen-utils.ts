/**
 * PptxGenJS: Utility Methods
 */

import { parseXml } from '@rgrove/parse-xml'

import { EMU, REGEX_HEX_COLOR, DEF_FONT_COLOR, DEF_TEXT_GLOW, ONEPT, SchemeColor, SCHEME_COLORS, TILE_ALIGNMENTS } from './core-enums'
import { PresLayout, TextGlowProps, PresSlide, SlideLayout, ShapeFillProps, Color, ShapeLineProps, Coord, ShadowProps, ShapeGradientProps, ShapePatternProps, ShapeImageFillProps, ModifiedThemeColor, ThemeProps, HexColor } from './core-interfaces'

/** debug namespace, used for both the log prefix and the `NODE_DEBUG` section name */
const DEBUG_NS = 'pptxgenjs'

/**
 * True Node process, not a browser bundle with a `process` polyfill.
 * Bundlers often inject `process.versions.node`, so also require no `document`.
 */
export function isNodeRuntime (): boolean {
	return typeof document === 'undefined'
		&& typeof process !== 'undefined'
		&& !!process.versions?.node
		&& process.release?.name === 'node'
}

/**
 * Dynamic Node builtin import that browser bundlers cannot statically resolve.
 *
 * Same idea as avoiding `Buffer.from`: a literal `import('node:https')` (or even
 * `import('node:' + name)` that Rollup later constant-folds) puts Node builtins
 * in the browser graph. Build the specifier at runtime and load it through
 * `Function` so Vite/Rsbuild/Rspack never see an import they can resolve.
 */
export function importNodeBuiltin<T> (name: 'fs' | 'http' | 'https'): Promise<T & { default: T }> {
	const specifier = ['node', name].join(':')
	return (Function('s', 'return import(s)') as (s: string) => Promise<T & { default: T }>)(specifier)
}

/**
 * Whether verbose diagnostics are enabled
 * - set `PPTXGENJS_DEBUG=1`, or include `pptxgenjs` in Node's `NODE_DEBUG`
 * @returns {boolean} debug enabled
 */
export function isDebugEnabled (): boolean {
	if (typeof process === 'undefined' || !process.env) return false
	return Boolean(process.env.PPTXGENJS_DEBUG) || (process.env.NODE_DEBUG ?? '').split(/[\s,]+/).includes(DEBUG_NS)
}

/**
 * Log a diagnostic message (no-op unless debug is enabled)
 * @param {unknown[]} args - console.debug arguments
 */
export function debugLog (...args: unknown[]): void {
	if (isDebugEnabled()) console.debug(`[${DEBUG_NS}]`, ...args)
}

/**
 * Encode bytes as base64 without referencing the `Buffer` identifier.
 * Browser bundlers (Vite) inject a broken `buffer` polyfill when they see `Buffer.from`.
 */
export function bytesToBase64 (bytes: Uint8Array): string {
	let bin = ''
	const chunk = 0x8000
	for (let i = 0; i < bytes.length; i += chunk) {
		bin += String.fromCharCode(...bytes.subarray(i, i + chunk))
	}
	return btoa(bin)
}

/** Decode base64 (or a data-URL) to bytes — browser + Node, no `Buffer`. */
export function base64ToBytes (strData: string): Uint8Array {
	const idxHdr = strData.indexOf('base64,')
	const strB64 = idxHdr > -1 ? strData.substring(idxHdr + 'base64,'.length) : strData
	const strBin = atob(strB64)
	const bytes = new Uint8Array(strBin.length)
	for (let idx = 0; idx < strBin.length; idx++) bytes[idx] = strBin.charCodeAt(idx)
	return bytes
}

/** UTF-8 string → base64 */
export function utf8ToBase64 (text: string): string {
	return bytesToBase64(new TextEncoder().encode(text))
}

/** base64 → UTF-8 string */
export function base64ToUtf8 (b64: string): string {
	return new TextDecoder().decode(base64ToBytes(b64))
}

/** Latin-1 / binary string → base64 (Node `res.setEncoding('binary')` payloads) */
export function binaryStringToBase64 (raw: string): string {
	const bytes = new Uint8Array(raw.length)
	for (let i = 0; i < raw.length; i++) bytes[i] = raw.charCodeAt(i) & 0xff
	return bytesToBase64(bytes)
}

/**
 * Recolor SVG markup for Martin-N `image.fill` (hex only)
 * - Replaces existing `"#RRGGBB"` / `"#RGB"` attribute values, or adds `fill` on bare `<path>` tags
 */
export function applySvgFillColor (svgText: string, color: string | undefined): string {
	if (!color || typeof color !== 'string') return svgText
	const hex = color.replace(/^#/, '')
	if (!/^[0-9A-Fa-f]{3}$|^[0-9A-Fa-f]{6}$/.test(hex)) return svgText

	const hexAttr = /"#[0-9A-Fa-f]{3,6}"/g
	if (hexAttr.test(svgText)) {
		return svgText.replace(/"#[0-9A-Fa-f]{3,6}"/g, `"#${hex}"`)
	}
	return svgText.replace(/<path(\s)/g, `<path fill="#${hex}"$1`)
}

/**
 * Apply `fill.color` to a base64 / data-URL SVG payload; returns unchanged input when not SVG text
 */
export function applySvgFillToDataUrl (data: string, color: string | undefined): string {
	if (!color || !data) return data
	const hdr = 'base64,'
	const idx = data.indexOf(hdr)
	const b64 = idx > -1 ? data.substring(idx + hdr.length) : data
	try {
		const svgText = base64ToUtf8(b64)
		if (!svgText.includes('<svg') && !svgText.includes('<path')) return data
		const recolored = applySvgFillColor(svgText, color)
		const outB64 = utf8ToBase64(recolored)
		return idx > -1 ? data.substring(0, idx + hdr.length) + outB64 : outB64
	} catch {
		return data
	}
}

/** friendly `dataLabelPosition` names mapped to their OOXML `c:dLblPos` codes (the codes stay accepted too) */
const DATA_LABEL_POS_CODES: Record<string, string> = {
	bottom: 'b',
	center: 'ctr',
	left: 'l',
	right: 'r',
	top: 't',
	insideEnd: 'inEnd',
	insideBase: 'inBase',
	outsideEnd: 'outEnd',
	bestFit: 'bestFit',
}

/**
 * Resolve a `dataLabelPosition` to the OOXML code valid for this chart type
 * - a value the chart type does not accept makes PowerPoint declare the file corrupt, so it is dropped
 * @param {string} position - user value: a friendly name (`'outsideEnd'`) or an OOXML code (`'outEnd'`)
 * @param {string} chartType - chart type being rendered, or undefined for a multi-type chart (translate only)
 * @param {string} barGrouping - bar grouping (stacked bars accept fewer positions than clustered)
 * @returns {string | undefined} OOXML code, or undefined when not valid for this chart type
 */
export function resolveDataLabelPosition (position: string, chartType?: string, barGrouping?: string): string | undefined {
	const code = DATA_LABEL_POS_CODES[position] ?? position

	// a multi-type chart has no single type to validate against - each sub-chart validates its own options
	if (!chartType) return code
	// REFERENCE: https://docs.microsoft.com/en-us/openspecs/office_standards/ms-oi29500/e2b1697c-7adc-463d-9081-3daef72f656f
	let valid: string[]
	switch (chartType) {
		case 'pie':
			valid = ['bestFit', 'ctr', 'inEnd', 'outEnd']
			break
		case 'bubble':
		case 'bubble3D':
		case 'line':
		case 'scatter':
			valid = ['b', 'ctr', 'l', 'r', 't']
			break
		case 'bar':
			// stacked bars have no "outside end" to sit against
			valid = (barGrouping ?? '').includes('tacked') ? ['ctr', 'inBase', 'inEnd'] : ['ctr', 'inBase', 'inEnd', 'outEnd']
			break
		default:
			// area, bar3D, doughnut, radar: PowerPoint takes no `c:dLblPos` at all
			valid = []
	}

	if (!valid.includes(code)) {
		console.warn(
			`[pptxgenjs] dataLabelPosition '${position}' is not valid for a '${chartType}' chart - ignoring it (valid: ${valid.length > 0 ? valid.join(', ') : 'none'})`
		)
		return undefined
	}

	return code
}

/**
 * Translates any type of `x`/`y`/`w`/`h` prop to EMU
 * - guaranteed to return a result regardless of undefined, null, etc. (0)
 * - {number} - 12800 (EMU)
 * - {number} - 0.5 (inches)
 * - {string} - "75%"
 * @param {number|string} size - numeric ("5.5") or percentage ("90%")
 * @param {'X' | 'Y'} xyDir - direction
 * @param {PresLayout} layout - presentation layout
 * @returns {number} calculated size
 */
export function getSmartParseNumber (size: Coord | undefined, xyDir: 'X' | 'Y', layout: PresLayout): number {
	// FIRST: Convert string numeric value if reqd
	if (typeof size === 'string' && !isNaN(Number(size))) size = Number(size)

	// CASE 1: Number in inches
	// Assume any number whose absolute value is less than 100 is inches (off-slide coords are negative)
	if (typeof size === 'number' && Math.abs(size) < 100) return inch2Emu(size)

	// CASE 2: Number is already converted to something other than inches
	// Assume any number whose absolute value is 100+ is EMU already.
	if (typeof size === 'number' && Math.abs(size) >= 100) return size

	// CASE 3: Percentage (ex: '50%')
	if (typeof size === 'string' && size.includes('%')) {
		if (xyDir && xyDir === 'X') return Math.round((parseFloat(size) / 100) * layout.width)
		if (xyDir && xyDir === 'Y') return Math.round((parseFloat(size) / 100) * layout.height)

		// Default: Assume width (x/cx)
		return Math.round((parseFloat(size) / 100) * layout.width)
	}

	// LAST: Default value
	return 0
}

/**
 * Basic UUID Generator Adapted
 * @link https://stackoverflow.com/questions/105034/create-guid-uuid-in-javascript#answer-2117523
 * @param {string} uuidFormat - UUID format
 * @returns {string} UUID
 */
export function getUuid (uuidFormat: string): string {
	return uuidFormat.replace(/[xy]/g, function (c) {
		// Web Crypto API - a global in browsers and in Node >=20 (this package's engines floor).
		// Mask the low nibble (0-15) - an unbiased reduction (no modulo) for a hex digit.
		const r = globalThis.crypto.getRandomValues(new Uint8Array(1))[0] & 0x0f
		const v = c === 'x' ? r : (r & 0x3) | 0x8
		return v.toString(16)
	})
}

/** deprecation keys already warned about - each fires once per process */
const _warnedDeprecations = new Set<string>()

/**
 * Warn about a deprecated option/usage - once per key, so migration guidance appears without flooding the console
 * @param {string} key - unique key for this deprecation
 * @param {string} message - migration guidance
 */
export function warnDeprecatedOnce (key: string, message: string): void {
	if (_warnedDeprecations.has(key)) return
	_warnedDeprecations.add(key)
	console.warn(`[pptxgenjs] DEPRECATED: ${message}`)
}

/**
 * Replace special XML characters with HTML-encoded strings
 * @param {string} xml - XML string to encode
 * @returns {string} escaped XML
 */
export function encodeXmlEntities (xml: string | undefined): string {
	// NOTE: Dont use short-circuit eval here as value c/b "0" (zero) etc.!
	if (typeof xml === 'undefined' || xml == null) return ''
	return xml.toString().replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;')
}

/**
 * Check that a string is a well-formed XML fragment (strict XML 1.0 via
 * `@rgrove/parse-xml`). Used to reject caller-supplied markup (e.g. OMML)
 * that would otherwise corrupt the package — PowerPoint shows a repair
 * prompt (or refuses to open) for malformed slide XML.
 * @param {string} xml - XML fragment to validate
 * @returns {string | null} description of the first problem, or null when well-formed
 */
export function validateXmlFragment (xml: string): string | null {
	try {
		// Synthetic root: fragments may have multiple top-level elements
		parseXml(`<x>${xml}</x>`)
		return null
	} catch (error) {
		return (error as Error).message.split('\n')[0]
	}
}

/**
 * Convert inches into EMU
 * @param {number|string} inches - as string or number
 * @returns {number} EMU value
 */
export function inch2Emu (inches: number | string): number {
	// NOTE: Provide Caller Safety: Numbers may get conv<->conv during flight, so be kind and do some simple checks to ensure inches were passed
	// Any absolute value over 100 is not inches (negative EMU is off-slide), so return it unchanged
	if (typeof inches === 'number' && Math.abs(inches) > 100) return inches
	if (typeof inches === 'string') inches = Number(inches.replace(/in*/gi, ''))
	return Math.round(EMU * inches)
}

/**
 * Convert `pt` into points (using `ONEPT`)
 * @param {number|string} pt
 * @returns {number} value in points (`ONEPT`)
 */
export function valToPts (pt: number | string | undefined): number {
	const points = Number(pt) || 0
	return isNaN(points) ? 0 : Math.round(points * ONEPT)
}

/**
 * Convert a margin component to EMU (same dual-unit rule as table cells)
 * - `>= 1` → points (`valToPts`)
 * - `< 1` → inches (`inch2Emu`)
 * Used so text margins stay consistent with table margins (mikemeerschaert/fix-inconsistent-margins)
 * @param {number} val margin component
 * @returns {number} EMU
 */
export function marginToEmu (val: number): number {
	const n = Number(val) || 0
	return n >= 1 ? valToPts(n) : inch2Emu(n)
}

/**
 * Convert degrees (0..360) to PowerPoint `rot` value
 * @param {number} d degrees
 * @returns {number} calculated `rot` value
 */
export function convertRotationDegrees (d: number | undefined): number {
	d = d || 0
	return Math.round((d > 360 ? d - 360 : d) * 60000)
}

/**
 * Converts component value to hex value
 * @param {number} c - component color
 * @returns {string} hex string
 */
export function componentToHex (c: number): string {
	const hex = c.toString(16)
	return hex.length === 1 ? '0' + hex : hex
}

/**
 * Converts RGB colors from css selectors to Hex for Presentation colors
 * @param {number} r - red value
 * @param {number} g - green value
 * @param {number} b - blue value
 * @returns {string} XML string
 */
export function rgbToHex (r: number, g: number, b: number): string {
	return (componentToHex(r) + componentToHex(g) + componentToHex(b)).toUpperCase()
}

/** Office theme defaults: dk1, lt1, dk2, lt2, accent1-6, hlink, folHlink */
export const DEF_THEME_COLORS: readonly HexColor[] = [
	'000000', 'FFFFFF', '44546A', 'E7E6E6',
	'4472C4', 'ED7D31', 'A5A5A5', 'FFC000', '5B9BD5', '70AD47',
	'0563C1', '954F72',
]

/**
 * Validate/normalize `theme.themeColors` to exactly 12 hex RGB values.
 * Invalid length or entries fall back to Office defaults (with a warning).
 */
export function resolveThemeColors (theme?: ThemeProps): HexColor[] {
	const input = theme?.themeColors
	let colors: HexColor[]
	if (!input) {
		colors = [...DEF_THEME_COLORS]
	} else if (input.length !== 12) {
		console.warn(`[pptxgenjs] theme.themeColors must contain exactly 12 hex colors (got ${input.length}); using Office defaults`)
		colors = [...DEF_THEME_COLORS]
	} else {
		colors = input.map((raw, idx) => {
			const hex = String(raw || '').replace('#', '').toUpperCase()
			if (!REGEX_HEX_COLOR.test(hex)) {
				console.warn(`[pptxgenjs] theme.themeColors[${idx}]="${raw}" is not a valid 6-digit hex; using default ${DEF_THEME_COLORS[idx]}`)
				return DEF_THEME_COLORS[idx]
			}
			return hex
		})
	}

	// Additive override for `a:hlink` (ECMA-376 clrScheme). Reject non-strings and invalid hex.
	if (typeof theme?.hlinkColor === 'string') {
		const hex = theme.hlinkColor.replace('#', '').toUpperCase()
		if (REGEX_HEX_COLOR.test(hex)) colors[10] = hex
		else console.warn(`[pptxgenjs] theme.hlinkColor="${theme.hlinkColor}" is not a valid 6-digit hex; using ${colors[10]}`)
	}
	return colors
}

export function isModifiedThemeColor (value: unknown): value is ModifiedThemeColor {
	return typeof value === 'object' && value !== null && 'baseColor' in value
}

/** Percent-based OOXML color transforms (0-100 → val*1000) */
const PERCENT_COLOR_MODIFIERS = [
	'alpha', 'alphaMod', 'alphaOff',
	'blue', 'blueMod', 'blueOff',
	'green', 'greenMod', 'greenOff',
	'red', 'redMod', 'redOff',
	'lum', 'lumMod', 'lumOff',
	'sat', 'satMod', 'satOff',
	'shade', 'tint',
	'hueMod', // hue modulation is a percentage, unlike hue/hueOff
] as const

const FLAG_COLOR_MODIFIERS = ['comp', 'gray', 'inv', 'gamma'] as const

/**
 * Emit DrawingML color transform children for a ModifiedThemeColor.
 * Careful vs upstream: hue/hueOff use fixed-angle units (deg×60000), not percent×1000.
 */
function handleModifiedColorProps (color: ModifiedThemeColor): string {
	let output = ''

	// OOXML percent units: API 0–100 → val*1000. Do not clamp — lumMod/satMod etc. routinely exceed 100.
	for (const modifier of PERCENT_COLOR_MODIFIERS) {
		const value = color[modifier]
		if (typeof value === 'number' && !Number.isNaN(value)) {
			output += `<a:${modifier} val="${Math.round(value * 1000)}"/>`
		}
	}

	// ST_PositiveFixedAngle / ST_FixedAngle — degrees × 60000
	if (typeof color.hue === 'number' && !Number.isNaN(color.hue)) {
		const deg = ((color.hue % 360) + 360) % 360
		output += `<a:hue val="${Math.round(deg * 60000)}"/>`
	}
	if (typeof color.hueOff === 'number' && !Number.isNaN(color.hueOff)) {
		output += `<a:hueOff val="${Math.round(color.hueOff * 60000)}"/>`
	}

	for (const modifier of FLAG_COLOR_MODIFIERS) {
		if (color[modifier]) output += `<a:${modifier}/>`
	}

	return output
}

/**  TODO: FUTURE: TODO-4.0:
 * @date 2022-04-10
 * @tldr this s/b a private method with all current calls switched to `genXmlColorSelection()`
 * @desc lots of code calls this method
 * @example [gen-charts.tx] `strXml += '<a:solidFill>' + createColorElement(seriesColor, `<a:alpha val="${Math.round(opts.chartColorsOpacity * 1000)}"/>`) + '</a:solidFill>'`
 * Thi sis wrong. We s/b calling `genXmlColorSelection()` instead as it returns `<a:solidfill>BLAH</a:solidFill>`!!
 */
/**
 * Create either a `a:schemeClr` - (scheme color) or `a:srgbClr` (hexa representation).
 * @param {Color|SCHEME_COLORS} colorInput - hex, scheme token, or ModifiedThemeColor
 * @param {string} innerElements - additional elements that adjust the color and are enclosed by the color element
 * @returns {string} XML string
 */
export function createColorElement (colorInput: Color | SCHEME_COLORS | undefined, innerElements?: string): string {
	const base = isModifiedThemeColor(colorInput) ? colorInput.baseColor : colorInput
	const colorStr = typeof base === 'string' ? base : DEF_FONT_COLOR
	let colorVal = (colorStr || '').replace('#', '')

	let kids = innerElements || ''
	if (isModifiedThemeColor(colorInput)) {
		kids += handleModifiedColorProps(colorInput)
	}

	if (
		!REGEX_HEX_COLOR.test(colorVal) &&
		colorVal !== SchemeColor.background1 &&
		colorVal !== SchemeColor.background2 &&
		colorVal !== SchemeColor.text1 &&
		colorVal !== SchemeColor.text2 &&
		colorVal !== SchemeColor.accent1 &&
		colorVal !== SchemeColor.accent2 &&
		colorVal !== SchemeColor.accent3 &&
		colorVal !== SchemeColor.accent4 &&
		colorVal !== SchemeColor.accent5 &&
		colorVal !== SchemeColor.accent6
	) {
		console.warn(`"${colorVal}" is not a valid scheme color or hex RGB! "${DEF_FONT_COLOR}" used instead. Only provide 6-digit RGB or 'pptx.SchemeColor' values!`)
		colorVal = DEF_FONT_COLOR
	}

	const tagName = REGEX_HEX_COLOR.test(colorVal) ? 'srgbClr' : 'schemeClr'
	const colorAttr = 'val="' + (REGEX_HEX_COLOR.test(colorVal) ? colorVal.toUpperCase() : colorVal) + '"'

	return kids ? `<a:${tagName} ${colorAttr}>${kids}</a:${tagName}>` : `<a:${tagName} ${colorAttr}/>`
}

/**
 * Creates `a:glow` element
 * @param {TextGlowProps} options glow properties
 * @param {TextGlowProps} defaults defaults for unspecified properties in `opts`
 * @see http://officeopenxml.com/drwSp-effects.php
 * { size: 8, color: 'FFFFFF', opacity: 0.75 };
 */
/**
 * Nominal ("coloured") brand for resolved glow options. The symbol is module-private and unexported,
 * so `ResolvedGlowProps` values can ONLY be produced by `resolveGlowOptions` below - a hand-built
 * object (even a full `Required<TextGlowProps>`) is not assignable. This statically guarantees that
 * anything reaching `createGlowElement` has passed through the defaults-merge boundary.
 */
const glowBrand: unique symbol = Symbol('resolvedGlow')
export type ResolvedGlowProps = Required<TextGlowProps> & { readonly [glowBrand]: boolean }

/**
 * Resolve boundary: merge user glow options over the documented defaults and brand the result. The
 * only constructor of `ResolvedGlowProps` (no cast - the brand is added by this factory).
 */
export function resolveGlowOptions (options: TextGlowProps | undefined): ResolvedGlowProps | undefined {
	if (!options) return undefined
	return { ...DEF_TEXT_GLOW, ...options, [glowBrand]: true }
}

export function createGlowElement (glow: ResolvedGlowProps): string {
	let strXml = ''
	const size = Math.round(glow.size * ONEPT)
	const color = glow.color
	const opacity = Math.round(glow.opacity * 100000)

	strXml += `<a:glow rad="${size}">`
	strXml += createColorElement(color, `<a:alpha val="${opacity}"/>`)
	strXml += '</a:glow>'

	return strXml
}

/**
 * Normalize nested `type:'gradient'` or flat `type:'linearGradient'` (sambauers) into ShapeGradientProps.
 */
function resolveGradientProps (props: ShapeFillProps): ShapeGradientProps | undefined {
	if (props.type === 'linearGradient') {
		return {
			type: 'linear',
			angle: props.angle,
			scaled: props.scaled,
			rotateWithShape: props.rotWithShape ?? props.gradient?.rotateWithShape,
			rotWithShape: props.rotWithShape,
			flip: props.flip ?? props.gradient?.flip,
			tileRect: props.tileRect ?? props.gradient?.tileRect,
			stops: props.stops ?? props.gradient?.stops ?? [],
		}
	}
	return props.gradient
}

/**
 * Create a DrawingML gradient fill element (`a:gradFill`)
 * @param {ShapeGradientProps | undefined} gradient gradient definition
 * @param {Color | string} fallbackColor color used if the gradient is unusable (fewer than 2 stops)
 * @param {string} fallbackInner alpha element(s) for the fallback solid fill
 * @returns {string} XML string
 * @see http://officeopenxml.com/drwSp-GradFill.php
 * @note MS-PPT (CT_GradientStopList) requires a minimum of 2 stops; we degrade to a solid fill rather than emit invalid XML
 */
function createGradientFillElement (gradient: ShapeGradientProps | undefined, fallbackColor: Color | string, fallbackInner: string): string {
	const stops = (gradient?.stops ?? [])
		.filter(stop => stop && stop.color !== undefined && stop.color !== null)
		.map(stop => {
			const pct = stop.pos ?? stop.position ?? 0
			return {
				color: stop.color,
				// @note `pos` is clamped: ST_PositiveFixedPercentage only permits 0-100000 (0-100%)
				pos: Math.round(Math.min(100, Math.max(0, Number(pct) || 0)) * 1000),
				transparency: stop.transparency,
			}
		})
		// @note MS-PPT renders stops in document order, so sort ascending to make `pos` authoritative
		.sort((a, b) => a.pos - b.pos)

	if (stops.length < 2) {
		// @note fall back rather than emit `<a:gsLst>` with too few stops, which MS-PPT flags as needing repair
		console.warn('`gradient.stops` requires at least 2 stops! A solid fill was used instead.')
		const soleColor = stops.length === 1 ? stops[0].color : fallbackColor
		return soleColor ? `<a:solidFill>${createColorElement(soleColor, fallbackInner)}</a:solidFill>` : ''
	}

	const gsLst = stops
		.map(stop => {
			const inner = typeof stop.transparency === 'number' ? `<a:alpha val="${Math.round((100 - Math.min(100, Math.max(0, stop.transparency))) * 1000)}"/>` : ''
			return `<a:gs pos="${stop.pos}">${createColorElement(stop.color, inner)}</a:gs>`
		})
		.join('')

	const rotateWithShape = gradient?.rotateWithShape === false || gradient?.rotWithShape === false ? 0 : 1
	const flip = gradient?.flip && gradient.flip !== 'none' ? ` flip="${gradient.flip}"` : ''

	let tileRectXml = ''
	const tr = gradient?.tileRect
	if (tr && (typeof tr.t === 'number' || typeof tr.r === 'number' || typeof tr.b === 'number' || typeof tr.l === 'number')) {
		const attrs: string[] = []
		if (typeof tr.t === 'number') attrs.push(`t="${Math.round(tr.t * 1000)}"`)
		if (typeof tr.r === 'number') attrs.push(`r="${Math.round(tr.r * 1000)}"`)
		if (typeof tr.b === 'number') attrs.push(`b="${Math.round(tr.b * 1000)}"`)
		if (typeof tr.l === 'number') attrs.push(`l="${Math.round(tr.l * 1000)}"`)
		tileRectXml = `<a:tileRect ${attrs.join(' ')}/>`
	}

	const geometry = gradient?.type === 'radial'
		? '<a:path path="circle"></a:path>'
		: (() => {
			// @note `ang` is ST_PositiveFixedAngle (0 to 21599999, in 60000ths of a degree), so normalize into 0-359 first
			// Always emit `<a:lin>` (unlike sambauers, which skipped angle 0 / falsy)
			const degrees = ((Number(gradient?.angle) || 0) % 360 + 360) % 360
			// truthy so sambauers demos using `scaled: 1` still work
			return `<a:lin ang="${Math.round(degrees * 60000)}" scaled="${gradient?.scaled ? 1 : 0}"/>`
		})()
	return `<a:gradFill rotWithShape="${rotateWithShape}"${flip}><a:gsLst>${gsLst}</a:gsLst>${geometry}${tileRectXml}</a:gradFill>`
}

/**
 * Create a DrawingML pattern fill element (`a:pattFill`)
 * Selective port of hakrueger/pattern — uses createColorElement so theme colors work.
 * @see http://officeopenxml.com/drwSp-patternFill.php
 */
function createPatternFillElement (pattern: ShapePatternProps | undefined, fallbackFg: string): string {
	const prst = pattern?.prst || 'cross'
	const fg = String(pattern?.color || fallbackFg || '000000')
	const bg = String(pattern?.bgColor || 'FFFFFF')
	return `<a:pattFill prst="${prst}"><a:bgClr>${createColorElement(bg)}</a:bgClr><a:fgClr>${createColorElement(fg)}</a:fgClr></a:pattFill>`
}

/**
 * Create a DrawingML `a:blipFill` for a picture fill (ECMA-376 20.1.8.14)
 * - the image relationship is resolved when the object is created
 */
function createImageFillElement (image: ShapeImageFillProps, rId: number): string {
	let mode = '<a:stretch><a:fillRect/></a:stretch>'
	if (image.sizing === 'tile') {
		const scale = typeof image.scale === 'number' && isFinite(image.scale) && image.scale > 0 ? Math.round(image.scale * 1000) : 100000
		const rawAlgn = image.alignment ?? 'tl'
		const algn = TILE_ALIGNMENTS.has(rawAlgn) ? rawAlgn : 'tl'
		if (image.alignment && !TILE_ALIGNMENTS.has(image.alignment)) {
			console.warn(`[pptxgenjs] unknown tile alignment "${String(image.alignment)}" - "tl" used instead`)
		}
		mode = `<a:tile tx="0" ty="0" sx="${scale}" sy="${scale}" flip="none" algn="${algn}"/>`
	}

	return `<a:blipFill rotWithShape="${image.rotateWithShape === false ? '0' : '1'}"><a:blip r:embed="rId${rId}"/>${mode}</a:blipFill>`
}

/**
 * Create color selection
 * @param {Color | ShapeFillProps | ShapeLineProps} props fill props
 * @returns XML string
 */
export function genXmlColorSelection (props: Color | ShapeFillProps | ShapeLineProps | undefined): string {
	let fillType = 'solid'
	let colorVal: Color = ''
	let internalElements = ''
	let outText = ''

	if (props) {
		// ModifiedThemeColor as a bare Color value (must not be treated as ShapeFillProps)
		if (typeof props === 'string') {
			colorVal = props
		} else if (isModifiedThemeColor(props)) {
			colorVal = props
		} else {
			if (props.type) fillType = props.type
			else if (props.image?._rId || props._rId) fillType = 'image'
			if (props.color) colorVal = props.color
			if (props.alpha) internalElements += `<a:alpha val="${Math.round((100 - props.alpha) * 1000)}"/>` // DEPRECATED: @deprecated v3.3.0
			if (props.transparency) internalElements += `<a:alpha val="${Math.round((100 - props.transparency) * 1000)}"/>`
		}

		switch (fillType) {
			case 'solid':
				outText += `<a:solidFill>${createColorElement(colorVal, internalElements)}</a:solidFill>`
				break
			case 'gradient':
			case 'linearGradient':
				outText += createGradientFillElement(
					typeof props === 'string' || isModifiedThemeColor(props) ? undefined : resolveGradientProps(props),
					colorVal || '',
					internalElements,
				)
				break
			case 'pattern':
				outText += createPatternFillElement(
					typeof props === 'string' || isModifiedThemeColor(props) ? undefined : props.pattern,
					typeof colorVal === 'string' ? colorVal : String(colorVal.baseColor),
				)
				break
			case 'image': {
				const image = typeof props === 'string' || isModifiedThemeColor(props) ? undefined : props.image
				const rId = image?._rId ?? (typeof props === 'string' || isModifiedThemeColor(props) ? undefined : props._rId)
				if (rId) {
					outText += createImageFillElement(image ?? {}, rId)
				} else {
					console.warn('[pptxgenjs] `fill.type:"image"` requires `fill.image.data` or `fill.image.path` - fill omitted')
				}
				break
			}
			case 'group':
				outText += '<a:grpFill/>'
				break
			case 'none':
				outText += '<a:noFill/>'
				break
			default: // keep a statement so empty-default is not stripped, then flagged by no-default
				outText += ''
				break
		}
	}

	return outText
}

/**
 * Get a new rel ID (rId) for charts, media, etc.
 * @param {PresSlide} target - the slide to use
 * @returns {number} count of all current rels plus 1 for the caller to use as its "rId"
 */
export function getNewRelId (target: PresSlide | SlideLayout): number {
	return target._rels.length + target._relsChart.length + target._relsMedia.length + 1
}

/**
 * Checks shadow options passed by user and performs corrections if needed.
 * @param {ShadowProps} ShadowProps - shadow options
 */
export function correctShadowOptions (ShadowProps: ShadowProps): ShadowProps | undefined {
	if (!ShadowProps || typeof ShadowProps !== 'object') {
		if (ShadowProps) console.warn('[pptxgenjs] `shadow` must be an object (ex: `{shadow: {type:\'outer\'}}`) - value ignored')
		return
	}

	// Work on a copy - never mutate the caller's options object
	ShadowProps = { ...ShadowProps }

	// OPT: `type`
	if (ShadowProps.type !== 'outer' && ShadowProps.type !== 'inner' && ShadowProps.type !== 'none' && ShadowProps.type !== 'preset') {
		console.warn('Warning: shadow.type options are `outer`, `inner`, `preset` or `none`.')
		ShadowProps.type = 'outer'
	}

	// OPT: `angle`
	if (ShadowProps.angle) {
		// A: REALITY-CHECK
		if (isNaN(Number(ShadowProps.angle)) || ShadowProps.angle < 0 || ShadowProps.angle > 359) {
			console.warn('Warning: shadow.angle can only be 0-359')
			ShadowProps.angle = 270
		}

		// B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
		ShadowProps.angle = Math.round(Number(ShadowProps.angle))
	}

	// OPT: `opacity`
	if (ShadowProps.opacity) {
		// A: REALITY-CHECK
		if (isNaN(Number(ShadowProps.opacity)) || ShadowProps.opacity < 0 || ShadowProps.opacity > 1) {
			console.warn('Warning: shadow.opacity can only be 0-1')
			ShadowProps.opacity = 0.75
		}

		// B: ROBUST: Cast any type of valid arg to int: '12', 12.3, etc. -> 12
		ShadowProps.opacity = Number(ShadowProps.opacity)
	}

	// OPT: `color`
	if (ShadowProps.color) {
		// INCORRECT FORMAT
		if (ShadowProps.color.startsWith('#')) {
			console.warn('Warning: shadow.color should not include hash (#) character, , e.g. "FF0000"')
			ShadowProps.color = ShadowProps.color.replace('#', '')
		}
	}

	return ShadowProps
}
