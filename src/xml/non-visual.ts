/**
 * PptxGenJS: Non-visual drawing properties (`p:cNvPr`) and editing locks
 *
 * Every drawable object carries a `p:cNvPr` with its id, name, and alt text, and a sibling
 * `p:cNv*Pr` holding the lock element for its kind - `a:spLocks` for shapes, `a:picLocks` for
 * pictures, `a:graphicFrameLocks` for frames, `a:grpSpLocks` for groups, `a:cxnSpLocks` for
 * connectors. The lock attributes are shared, so one emitter serves all five.
 */

import { NonVisualProps, ShapeLockProps } from '../core-interfaces'
import { encodeXmlEntities } from '../gen-utils'

/** Lock attribute names, in the order the schema declares them */
const LOCK_ATTRS: Array<[keyof ShapeLockProps, string]> = [
	['noGroup', 'noGrp'],
	['noSelect', 'noSelect'],
	['noRotate', 'noRot'],
	['noChangeAspect', 'noChangeAspect'],
	['noMove', 'noMove'],
	['noResize', 'noResize'],
	['noEditPoints', 'noEditPoints'],
	['noAdjustHandles', 'noAdjustHandles'],
	['noChangeArrowheads', 'noChangeArrowheads'],
	['noChangeShapeType', 'noChangeShapeType'],
	['noTextEdit', 'noTextEdit'],
	['noCrop', 'noCrop'],
]

/** Which lock attributes each element actually permits */
const ALLOWED: Record<string, Set<string>> = {
	'a:spLocks': new Set(['noGrp', 'noSelect', 'noRot', 'noChangeAspect', 'noMove', 'noResize', 'noEditPoints', 'noAdjustHandles', 'noChangeArrowheads', 'noChangeShapeType', 'noTextEdit']),
	'a:picLocks': new Set(['noGrp', 'noSelect', 'noRot', 'noChangeAspect', 'noMove', 'noResize', 'noEditPoints', 'noAdjustHandles', 'noChangeArrowheads', 'noChangeShapeType', 'noCrop']),
	'a:graphicFrameLocks': new Set(['noGrp', 'noSelect', 'noChangeAspect', 'noMove', 'noResize', 'noDrilldown']),
	'a:grpSpLocks': new Set(['noGrp', 'noSelect', 'noRot', 'noChangeAspect', 'noMove', 'noResize', 'noUngrp']),
	'a:cxnSpLocks': new Set(['noGrp', 'noSelect', 'noRot', 'noChangeAspect', 'noMove', 'noResize', 'noEditPoints', 'noAdjustHandles', 'noChangeArrowheads', 'noChangeShapeType']),
}

/**
 * Create a lock element, or `''` when nothing is locked
 * - attributes the element does not permit are dropped: an unexpected one is a schema violation
 * @param {string} tag - lock element name for the object's kind
 * @param {ShapeLockProps} lock - lock props
 * @param {string} defaults - attributes the library already emits for this object
 * @returns {string} XML string
 */
export function genXmlLocks (tag: string, lock?: ShapeLockProps, defaults = ''): string {
	const allowed = ALLOWED[tag] ?? new Set<string>()
	let attrs = defaults

	LOCK_ATTRS.forEach(([prop, attr]) => {
		if (lock?.[prop] !== true) return
		if (!allowed.has(attr)) {
			console.warn(`[pptxgenjs] \`${String(prop)}\` does not apply to this object type - ignored`)
			return
		}
		// do not repeat an attribute the caller already gets by default
		if (defaults.includes(`${attr}="`)) return
		attrs += ` ${attr}="1"`
	})

	return attrs ? `<${tag}${attrs}/>` : ''
}

/**
 * Create the `p:cNvPr` element
 * @param {number} id - shape id, unique on the slide
 * @param {string} name - shape name shown in the selection pane (already XML-escaped by callers)
 * @param {NonVisualProps} props - non-visual options
 * @param {string} descr - alt-text description
 * @param {string} children - child elements (ex: hyperlinks); self-closes when empty
 * @returns {string} XML string
 */
export function genXmlCNvPr (id: number, name: string, props?: NonVisualProps, descr = '', children = ''): string {
	let attrs = `id="${id}" name="${name}"`
	// `descr` is the alt-text description; `title` is the separate alt-text title
	if (descr) attrs += ` descr="${encodeXmlEntities(descr)}"`
	if (props?.title) attrs += ` title="${encodeXmlEntities(props.title)}"`
	if (props?.hidden === true) attrs += ' hidden="1"'

	return children ? `<p:cNvPr ${attrs}>${children}</p:cNvPr>` : `<p:cNvPr ${attrs}/>`
}

/**
 * Create `p:cNvSpPr`, carrying the text-box flag and any shape locks
 * @param {object} options - shape options
 * @returns {string} XML string
 */
export function genXmlCNvSpPr (options: { isTextBox?: boolean, lock?: ShapeLockProps }): string {
	const txBox = options?.isTextBox ? ' txBox="1"' : ''
	const locks = genXmlLocks('a:spLocks', options?.lock)
	return locks ? `<p:cNvSpPr${txBox}>${locks}</p:cNvSpPr>` : `<p:cNvSpPr${txBox}/>`
}
