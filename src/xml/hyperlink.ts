/**
 * Hyperlinks and mouse-over actions.
 *
 * DrawingML gives the same concept two element names depending on where it appears:
 *
 * | host | click | mouse-over |
 * | :-- | :-- | :-- |
 * | `p:cNvPr` (shapes, pictures) | `a:hlinkClick` (20.1.2.2.22) | `a:hlinkHover` (20.1.2.2.23) |
 * | `a:rPr` (text runs) | `a:hlinkClick` | `a:hlinkMouseOver` (21.1.2.3.6) |
 *
 * Emitting `a:hlinkHover` inside `a:rPr` — or `a:hlinkMouseOver` inside `p:cNvPr` — produces markup
 * PowerPoint cannot parse, so the host decides the element name here rather than at each call site.
 */

import { HyperlinkProps } from '../core-interfaces'
import { encodeXmlEntities } from '../gen-utils'

/** Where the link lives, which decides the mouse-over element name */
export type HyperlinkHost = 'shape' | 'run'

/**
 * Build one `a:hlinkClick` / mouse-over element
 * @param {HyperlinkProps} link - link props
 * @param {'click' | 'hover'} trigger - which interaction this element describes
 * @param {HyperlinkHost} host - the element the link is attached to
 * @param {string} children - child elements (`a:snd`, `a:extLst`); self-closes when empty
 * @returns {string} XML string
 */
export function genXmlHyperlink (link: HyperlinkProps, trigger: 'click' | 'hover', host: HyperlinkHost, children = ''): string {
	const tag = trigger === 'click' ? 'a:hlinkClick' : host === 'run' ? 'a:hlinkMouseOver' : 'a:hlinkHover'

	let attrs = `r:id="rId${link._rId ?? 0}"`
	if (link.slide) {
		// an internal jump has no URL; the action attribute names the target kind
		attrs += ' action="ppaction://hlinksldjump"'
	} else {
		attrs += ' invalidUrl="" action="" tgtFrame=""'
	}
	attrs += ` tooltip="${link.tooltip ? encodeXmlEntities(link.tooltip) : ''}"`
	if (!link.slide) attrs += ' history="1"'
	// both default to false in the schema, so only write them when turned on
	if (link.highlightClick === true) attrs += ' highlightClick="1"'
	if (link.stopSoundsOnClick === true) attrs += ' endSnd="1"'

	const sound = genXmlHyperlinkSound(link)
	const inner = sound + children

	return inner ? `<${tag} ${attrs}>${inner}</${tag}>` : `<${tag} ${attrs}/>`
}

/**
 * Build the `a:snd` action sound (ECMA-376 20.1.2.2.32 `CT_EmbeddedWAVAudioFile`)
 * - the relationship is resolved when the object is created, so `_sndRId` is set by then
 * @param {HyperlinkProps} link - link props
 * @returns {string} XML string
 */
function genXmlHyperlinkSound (link: HyperlinkProps): string {
	if (!link.sound?._sndRId) return ''
	return `<a:snd r:embed="rId${link.sound._sndRId}" name="${encodeXmlEntities(link.sound.name ?? 'sound.wav')}"/>`
}
