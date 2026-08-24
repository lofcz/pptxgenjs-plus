import { RAW_METRICS } from './metrics-data'

/** Points per inch - every `addX` option is in inches, every font size is in points */
const PT_PER_IN = 72

/**
 * Average advance width, in ems, used when nothing better is available.
 *
 * The core's own auto-paging constant works out to roughly this, and it is the reason
 * `autoPageCharWeight` exists: one number cannot describe a proportional font. Kept only so
 * `measureText` always returns something, and reported as `source: 'estimate'` so callers can
 * tell they are getting a guess.
 */
const ESTIMATED_CHAR_WIDTH = 0.5

/** Line height in ems when the font's own ascent/descent are unknown */
const ESTIMATED_LINE_HEIGHT = 1.2

export interface FontMetrics {
	/** Advance width per codepoint, in ems */
	widths: Map<number, number>
	/** Advance width for a codepoint the font does not cover, in ems */
	fallbackWidth: number
	/** Distance from baseline to the top of the line box, in ems */
	ascent: number
	/** Distance from baseline to the bottom of the line box, in ems (negative) */
	descent: number
	/** Extra leading between lines, in ems */
	lineGap: number
}

export interface MeasureProps {
	/** Typeface name, matched case-insensitively @default 'Calibri' */
	fontFace?: string
	/** Font size in points @default 12 */
	fontSize?: number
	bold?: boolean
	italic?: boolean
	/**
	 * Wrap width in inches. Omit to measure the text as a single unbroken line.
	 * Text boxes reduce this by their `margin`; pass the width you actually have for text.
	 */
	w?: number
	/** Line spacing as a multiple of the single-line height @default 1 */
	lineSpacingMultiple?: number
}

export interface Measurement {
	/** Width of the widest line, inches */
	w: number
	/** Total height of every line, inches */
	h: number
	/** The text as it wraps at the given width */
	lines: string[]
	/** Height of one line, inches */
	lineHeight: number
	/** Where the numbers came from - `'estimate'` means no metrics were available */
	source: 'canvas' | 'metrics' | 'estimate'
}

/** Parsed `RAW_METRICS` plus anything `registerFontMetrics` added, keyed by lowercased name */
const registry = new Map<string, FontMetrics>()

/** `codepoint:count:width` runs, per-mille of the em */
function decodeWidths (encoded: string): Map<number, number> {
	const widths = new Map<number, number>()
	for (const run of encoded.split(',')) {
		const [start, count, width] = run.split(':').map(Number)
		for (let i = 0; i < count; i++) widths.set(start + i, width / 1000)
	}
	return widths
}

for (const [name, raw] of Object.entries(RAW_METRICS)) {
	registry.set(name, {
		widths: decodeWidths(raw.widths),
		fallbackWidth: raw.fallbackWidth / 1000,
		ascent: raw.ascent / 1000,
		descent: raw.descent / 1000,
		lineGap: raw.lineGap / 1000,
	})
}

export interface RegisterFontMetricsProps {
	/** Advance width per character, in ems - `{ ' ': 0.226, M: 0.855 }` or codepoint keys */
	widths: Record<string, number> | Map<number, number>
	/** Advance width for uncovered characters, in ems @default the width of a space, else 0.5 */
	fallbackWidth?: number
	/** Baseline to line-box top, in ems @default 1 */
	ascent?: number
	/** Baseline to line-box bottom, in ems, negative @default -0.25 */
	descent?: number
	/** Extra leading, in ems @default 0 */
	lineGap?: number
}

/**
 * Teach `measureText` a font it does not ship metrics for.
 *
 * Bundled metrics cover Calibri (via the metric-compatible Carlito) in all four styles. Register
 * a style by its full name as used in text options plus the style suffix - `'Aptos Bold'` - so
 * `measureText('...', { fontFace: 'Aptos', bold: true })` finds it.
 *
 * @example
 * registerFontMetrics('Aptos', { widths: { ' ': 0.22, a: 0.5 }, ascent: 0.94, descent: -0.27 })
 */
export function registerFontMetrics (fontFace: string, props: RegisterFontMetricsProps): void {
	if (!fontFace || typeof fontFace !== 'string') throw new Error('registerFontMetrics: fontFace is required')

	const widths = new Map<number, number>()
	if (props.widths instanceof Map) {
		for (const [code, width] of props.widths) widths.set(code, width)
	} else {
		for (const [key, width] of Object.entries(props.widths)) {
			const code = key.length === 1 ? key.codePointAt(0) : Number(key)
			if (code === undefined || !Number.isInteger(code)) throw new Error(`registerFontMetrics: "${key}" is neither a single character nor a codepoint`)
			widths.set(code, width)
		}
	}
	if (widths.size === 0) throw new Error('registerFontMetrics: widths is empty')

	registry.set(fontFace.toLowerCase(), {
		widths,
		fallbackWidth: props.fallbackWidth ?? widths.get(0x20) ?? ESTIMATED_CHAR_WIDTH,
		ascent: props.ascent ?? 1,
		descent: props.descent ?? -0.25,
		lineGap: props.lineGap ?? 0,
	})
}

/** Metrics for the requested face+style, or undefined when none are registered */
function lookup (fontFace: string, bold: boolean, italic: boolean): FontMetrics | undefined {
	const suffix = `${bold ? ' bold' : ''}${italic ? ' italic' : ''}`
	// Fall back to the regular weight rather than to nothing: right glyph set, wrong weight beats a
	// flat 0.5em guess for every character.
	return registry.get(`${fontFace.toLowerCase()}${suffix}`) ?? registry.get(fontFace.toLowerCase())
}

/** A 2D canvas is the only accurate measurement available in a browser, and it needs no data */
function canvasContext (): { measureText: (text: string) => { width: number } } | undefined {
	if (typeof document === 'undefined' || typeof document.createElement !== 'function') return undefined
	try {
		const context = document.createElement('canvas').getContext('2d')
		return context ?? undefined
	} catch {
		return undefined
	}
}

interface Ruler {
	/** Width of a run of text, in ems */
	widthOf: (text: string) => number
	/** Height of one line, in ems */
	lineHeight: number
	source: Measurement['source']
}

function ruler (fontFace: string, fontSize: number, bold: boolean, italic: boolean): Ruler {
	const metrics = lookup(fontFace, bold, italic)
	const context = canvasContext()

	if (context) {
		// Measure at a fixed size and divide, so the ruler stays in ems like the other tiers
		const probe = 100
		const style = `${italic ? 'italic ' : ''}${bold ? 'bold ' : ''}${probe}px "${fontFace}"`
		// eslint-disable-next-line @typescript-eslint/no-explicit-any
		;(context as any).font = style
		return {
			widthOf: text => context.measureText(text).width / probe,
			lineHeight: metrics ? metrics.ascent - metrics.descent + metrics.lineGap : ESTIMATED_LINE_HEIGHT,
			source: 'canvas',
		}
	}

	if (metrics) {
		return {
			widthOf: text => {
				let total = 0
				for (const char of text) total += metrics.widths.get(char.codePointAt(0) ?? 0) ?? metrics.fallbackWidth
				return total
			},
			lineHeight: metrics.ascent - metrics.descent + metrics.lineGap,
			source: 'metrics',
		}
	}

	return {
		widthOf: text => [...text].length * ESTIMATED_CHAR_WIDTH,
		lineHeight: ESTIMATED_LINE_HEIGHT,
		source: 'estimate',
	}
}

/** Greedy word wrap: the same algorithm every renderer uses, over whichever ruler is available */
function wrap (text: string, maxEm: number, widthOf: (text: string) => number): string[] {
	const lines: string[] = []

	for (const paragraph of text.split('\n')) {
		if (paragraph === '') {
			lines.push('')
			continue
		}

		let line = ''
		for (const word of paragraph.split(' ')) {
			const candidate = line === '' ? word : `${line} ${word}`
			if (widthOf(candidate) <= maxEm || line === '') {
				// A word wider than the box still starts its own line, then gets broken below
				line = candidate
			} else {
				lines.push(line)
				line = word
			}

			// Break a word that cannot fit on a line of its own, so nothing silently overflows
			while (widthOf(line) > maxEm && [...line].length > 1) {
				const chars = [...line]
				let cut = chars.length - 1
				while (cut > 1 && widthOf(chars.slice(0, cut).join('')) > maxEm) cut--
				lines.push(chars.slice(0, cut).join(''))
				line = chars.slice(cut).join('')
			}
		}
		lines.push(line)
	}

	return lines
}

/**
 * Measure text the way a renderer would, in inches.
 *
 * Uses the best source available: a browser 2D canvas, then bundled advance widths (Calibri via
 * the metric-compatible Carlito), then a flat per-character estimate. Check `source` on the result
 * when it matters - `'estimate'` means the numbers are a guess, and
 * {@link registerFontMetrics} is how to fix that.
 *
 * PowerPoint substitutes a different font when the viewer's machine lacks yours and re-wraps with
 * its metrics, so treat the result as accurate-for-this-font rather than guaranteed.
 *
 * @example
 * const { h, lines } = measureText('a long paragraph...', { w: 4, fontSize: 14 })
 * slide.addText(text, { x: 1, y: 1, w: 4, h })
 */
export function measureText (text: string, props: MeasureProps = {}): Measurement {
	const { fontFace = 'Calibri', fontSize = 12, bold = false, italic = false, w, lineSpacingMultiple = 1 } = props

	if (fontSize <= 0) throw new Error(`measureText: fontSize must be > 0 (got ${fontSize})`)
	if (w !== undefined && !(w > 0)) throw new Error(`measureText: w must be > 0 when given (got ${w})`)
	if (!(lineSpacingMultiple > 0)) throw new Error(`measureText: lineSpacingMultiple must be > 0 (got ${lineSpacingMultiple})`)

	const { widthOf, lineHeight, source } = ruler(fontFace, fontSize, bold, italic)
	const emToInch = fontSize / PT_PER_IN

	const lines = w === undefined ? String(text).split('\n') : wrap(String(text), w / emToInch, widthOf)
	const widest = lines.reduce((max, line) => Math.max(max, widthOf(line)), 0)
	const lineHeightIn = lineHeight * emToInch * lineSpacingMultiple

	return {
		w: widest * emToInch,
		h: lines.length * lineHeightIn,
		lines,
		lineHeight: lineHeightIn,
		source,
	}
}
