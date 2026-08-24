/**
 * Slide transition XML generation (`<p:transition>`).
 *
 * ECMA-376 §19.3.1.50 `CT_SlideTransition` carries the base transition set.
 * MS-PPTX §2.2.1 adds the 2010+ modern transitions inside an `<mc:AlternateContent>` wrapper
 * (Choice in the `p14`/`p15` namespace, Fallback to a base transition) plus `morph` (§2.6.1.1,
 * `p16` namespace). Modern consumers (PowerPoint 2010+) read the Choice; older readers and other
 * tools use the Fallback, so the file stays valid everywhere.
 */
import type { TRANSITION_TYPE } from './core-enums'
import type { SlideTransitionProps } from './core-interfaces'

const P14_NS = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
const P15_NS = 'http://schemas.microsoft.com/office/powerpoint/2012/main'
const P16_NS = 'http://schemas.microsoft.com/office/powerpoint/2016/main'
const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'

/**
 * Base ECMA-376 transition types and the attribute their inner element takes.
 * Maps: type -> [element name, attribute kind]. The attribute kind selects which
 * `dir`/`orient`/`spokes`/`thruBlk` attribute (if any) to emit (ECMA-376 §19.3.1.50 + §A.4).
 */
const BASE: Record<string, [string, 'none' | 'side' | 'orient' | 'eight' | 'corner' | 'inout' | 'split' | 'wheel' | 'black']> = {
	blinds: ['blinds', 'orient'],
	checker: ['checker', 'orient'],
	circle: ['circle', 'none'],
	comb: ['comb', 'orient'],
	cover: ['cover', 'eight'],
	cut: ['cut', 'black'],
	diamond: ['diamond', 'none'],
	dissolve: ['dissolve', 'none'],
	fade: ['fade', 'black'],
	newsflash: ['newsflash', 'none'],
	none: ['none', 'none'],
	plus: ['plus', 'none'],
	pull: ['pull', 'eight'],
	push: ['push', 'side'],
	random: ['random', 'none'],
	randomBar: ['randomBar', 'orient'],
	split: ['split', 'split'],
	strips: ['strips', 'corner'],
	wedge: ['wedge', 'none'],
	wheel: ['wheel', 'wheel'],
	wipe: ['wipe', 'side'],
	zoom: ['zoom', 'inout'],
}

/**
 * Modern (MS-PPTX) transitions: type -> [namespace, attribute kind, fallback base type].
 * morph has no directional attribute (§2.6.3.1 CT_MorphTransition only takes `option`).
 */
const MODERN: Record<string, [string, 'none' | 'side' | 'orient' | 'eight' | 'inout', string]> = {
	conveyor: [P14_NS, 'eight', 'push'],
	doors: [P14_NS, 'orient', 'split'],
	ferris: [P14_NS, 'eight', 'push'],
	flash: [P14_NS, 'none', 'fade'],
	flip: [P14_NS, 'eight', 'push'],
	flythrough: [P14_NS, 'inout', 'zoom'],
	gallery: [P14_NS, 'eight', 'push'],
	glitter: [P14_NS, 'eight', 'dissolve'],
	honeycomb: [P14_NS, 'none', 'dissolve'],
	morph: [P16_NS, 'none', 'fade'],
	pan: [P14_NS, 'side', 'push'],
	prism: [P14_NS, 'eight', 'push'],
	reveal: [P14_NS, 'side', 'fade'],
	ripple: [P14_NS, 'none', 'dissolve'],
	shred: [P14_NS, 'none', 'dissolve'],
	switch: [P14_NS, 'eight', 'push'],
	vortex: [P14_NS, 'side', 'push'],
	warp: [P14_NS, 'inout', 'zoom'],
	wheelReverse: [P14_NS, 'none', 'wheel'],
	window: [P14_NS, 'orient', 'split'],
}

/**
 * Map friendly caller names to ECMA-376 transition tokens.
 * OOXML short codes (`l`/`horz`/`ld`) stay as-is. `ll`/`rl` are accepted as aliases
 * for the spec corner tokens `ld`/`rd` (ST_TransitionCornerDirectionType).
 */
const DIR_ALIASES: Record<string, string> = {
	l: 'l', r: 'r', u: 'u', d: 'd',
	left: 'l', right: 'r', up: 'u', down: 'd',
	horz: 'horz', vert: 'vert',
	horizontal: 'horz', vertical: 'vert',
	lu: 'lu', ru: 'ru', ld: 'ld', rd: 'rd',
	leftup: 'lu', rightup: 'ru', leftdown: 'ld', rightdown: 'rd',
	ll: 'ld', rl: 'rd',
	in: 'in', out: 'out',
}

function normalizeDir (value?: string): string | undefined {
	if (!value) return undefined
	const key = value.replace(/[-_\s]/g, '').toLowerCase()
	return DIR_ALIASES[key] ?? value
}

/** Build the inner element + its directional attribute for a base transition. */
function baseInner(type: string, kind: string, p: SlideTransitionProps): string {
	switch (kind) {
		case 'none':
			return `<p:${type}/>`
		case 'side':
			return `<p:${type}${dirAttr(['l', 'r', 'u', 'd'], p)}/>`
		case 'orient':
			return `<p:${type}${dirAttr(['horz', 'vert'], p)}/>`
		case 'eight':
			return `<p:${type}${dirAttr(['l', 'r', 'u', 'd', 'lu', 'ru', 'ld', 'rd'], p)}/>`
		case 'corner':
			return `<p:${type}${dirAttr(['lu', 'ru', 'ld', 'rd'], p)}/>`
		case 'inout':
			return `<p:${type}${dirAttr(['in', 'out'], p)}/>`
		case 'split': {
			// CT_SplitTransition: orient=ST_Direction (horz|vert), dir=ST_TransitionInOutDirectionType (in|out)
			const raw = normalizeDir(p.direction)
			const orient = normalizeDir(p.orient) ?? (raw === 'vert' || raw === 'horz' ? raw : 'horz')
			const splitDir = raw === 'in' || raw === 'out' ? raw : 'out'
			return `<p:${type} orient="${orient === 'vert' ? 'vert' : 'horz'}" dir="${splitDir}"/>`
		}
		case 'wheel':
			return `<p:${type}${p.spokes ? ` spokes="${p.spokes}"` : ''}/>`
		case 'black':
			return `<p:${type}${p.thruBlk ? ' thruBlk="1"' : ''}/>`
		default:
			return `<p:${type}/>`
	}
}

/** Emit a `dir` attribute when the supplied direction is valid for this transition. */
function dirAttr(valid: string[], p: SlideTransitionProps): string {
	const d = normalizeDir(p.direction)
	const token = d && valid.includes(d) ? d : undefined
	return token ? ` dir="${token}"` : ''
}

/**
 * Generates `<p:transition>...</p:transition>` for a slide, or '' if no transition is set.
 * Placed after `<p:clrMapOvr>` and before `<p:timing>` in the `sld` content model (ECMA-376 §19.3.1.38).
 */
export function genXmlTransition(slide: { transition?: SlideTransitionProps }): string {
	const t = slide.transition
	if (!t || !t.type) return ''

	const speed = t.speed && ['slow', 'med', 'fast'].includes(t.speed) ? t.speed : undefined
	const dur = typeof t.duration === 'number' && t.duration > 0 ? ` p14:dur="${Math.round(t.duration)}"` : ''
	// Fallback must stay in the ECMA-376 attribute set (`spd`); `p14:dur` is Choice-only.
	const fallbackAttrs =
		(speed ? ` spd="${speed}"` : '') +
		(t.advClick === false ? ' advClick="0"' : '') +
		(typeof t.advTm === 'number' && t.advTm >= 0 ? ` advTm="${Math.round(t.advTm)}"` : '')
	const attrs = fallbackAttrs + dur

	const type = t.type as TRANSITION_TYPE

	// Modern transition: wrap in mc:AlternateContent with a base-type Fallback.
	if (MODERN[type]) {
		const [ns, kind, fallbackType] = MODERN[type]
		const prefix = ns === P16_NS ? 'p16' : ns === P15_NS ? 'p15' : 'p14'
		const inner =
			kind === 'side' ? `<${prefix}:${type}${dirAttr(['l', 'r', 'u', 'd'], t)}/>`
				: kind === 'orient' ? `<${prefix}:${type}${dirAttr(['horz', 'vert'], t)}/>`
					: kind === 'eight' ? `<${prefix}:${type}${dirAttr(['l', 'r', 'u', 'd', 'lu', 'ru', 'ld', 'rd'], t)}/>`
						: kind === 'inout' ? `<${prefix}:${type}${dirAttr(['in', 'out'], t)}/>`
							: `<${prefix}:${type}/>`
		const [fbName, fbKind] = BASE[fallbackType]
		const fallback = baseInner(fbName, fbKind, { type: fallbackType as TRANSITION_TYPE })
		return (
			`<mc:AlternateContent xmlns:mc="${MC_NS}" xmlns:${prefix}="${ns}">` +
			`<mc:Choice Requires="${prefix}">` +
			`<p:transition${attrs}>${inner}</p:transition>` +
			'</mc:Choice>' +
			`<mc:Fallback><p:transition${fallbackAttrs}>${fallback}</p:transition></mc:Fallback>` +
			'</mc:AlternateContent>'
		)
	}

	// Base ECMA-376 transition.
	if (BASE[type]) {
		const [name, kind] = BASE[type]
		return `<p:transition${attrs}>${baseInner(name, kind, t)}</p:transition>`
	}

	return ''
}
