/**
 * Text-body, run, and placeholder rendering.
 */

import {
	BULLET_TYPES,
	CRLF,
	DEF_BULLET_MARGIN,
	PLACEHOLDER_TYPES,
	SLIDE_OBJECT_TYPES,
	TEXT_FIELD_TYPES,
} from '../core-enums'
import {
	ClassificationOutcome,
	Color,
	DesignerTag,
	ISlideObject,
	ObjectOptions,
	TableCell,
	TextProps,
	TextPropsOptions,
	TextShapeType,
} from '../core-interfaces'
import { genXmlHyperlink } from './hyperlink'

/** ECMA-376 §5.1.12.76 `ST_TextShapeType` — used to validate `a:prstTxWarp@prst`. */
const TEXT_SHAPE_TYPES: ReadonlySet<TextShapeType> = new Set([
	'textNoShape', 'textPlain', 'textStop', 'textTriangle', 'textTriangleInverted',
	'textChevron', 'textChevronInverted', 'textRingInside', 'textRingOutside',
	'textArchUp', 'textArchDown', 'textCircle', 'textButton',
	'textArchUpPour', 'textArchDownPour', 'textCirclePour', 'textButtonPour',
	'textCurveUp', 'textCurveDown', 'textCanUp', 'textCanDown',
	'textWave1', 'textWave2', 'textDoubleWave1', 'textWave4',
	'textInflate', 'textDeflate', 'textInflateBottom', 'textDeflateBottom',
	'textInflateTop', 'textDeflateTop', 'textDeflateInflate', 'textDeflateInflateDeflate',
	'textFadeRight', 'textFadeLeft', 'textFadeUp', 'textFadeDown',
	'textSlantUp', 'textSlantDown', 'textCascadeUp', 'textCascadeDown',
])
import {
	createColorElement,
	createGlowElement,
	encodeXmlEntities,
	genXmlColorSelection,
	inch2Emu,
	resolveGlowOptions,
	validateXmlFragment,
	valToPts,
	warnDeprecatedOnce,
} from '../gen-utils'

/**
 * Build the `a:buClr` / `a:buSzPct` / `a:buSzPts` / `a:buFont` prefix shared by every bullet kind
 * - the size defaults to 100% because that is what every branch emitted before
 */
function genXmlBulletPrefix (bullet: { color?: Color, size?: number, sizePts?: number, fontFace?: string }): string {
	let xml = ''
	if (bullet.color) xml += `<a:buClr>${createColorElement(bullet.color)}</a:buClr>`

	// `a:buSzPct` and `a:buSzPts` are the same choice, so only one may appear
	if (typeof bullet.sizePts === 'number' && isFinite(bullet.sizePts) && bullet.sizePts > 0) {
		xml += `<a:buSzPts val="${Math.round(bullet.sizePts * 100)}"/>`
	} else {
		const pct = typeof bullet.size === 'number' && isFinite(bullet.size) ? Math.min(400, Math.max(25, bullet.size)) : 100
		if (typeof bullet.size === 'number' && (bullet.size < 25 || bullet.size > 400)) {
			console.warn(`[pptxgenjs] bullet \`size\` must be between 25 and 400 percent - "${String(bullet.size)}" clamped`)
		}
		xml += `<a:buSzPct val="${Math.round(pct * 1000)}"/>`
	}

	if (bullet.fontFace) xml += `<a:buFont typeface="${encodeXmlEntities(bullet.fontFace)}"/>`
	return xml
}

/**
 * Generate XML Paragraph Properties
 * @param {ISlideObject|TextProps} textObj - text object
 * @param {boolean} isDefault - array of default relations
 * @return {string} XML
 */
function genXmlParagraphProperties (textObj: ISlideObject | TextProps, isDefault: boolean): string {
	let strXmlBullet = ''
	let strXmlLnSpc = ''
	let strXmlParaSpc = ''
	let strXmlTabStops = ''
	const tag = isDefault ? 'a:lvl1pPr' : 'a:pPr'
	let bulletMarL = valToPts(DEF_BULLET_MARGIN)
	const options: TextPropsOptions = textObj.options ?? {}

	let paragraphPropXml = `<${tag}${options.rtlMode ? ' rtl="1" ' : ''}`

	// A: Build paragraphProperties
	{
		// OPTION: align
		if (options.align) {
			switch (options.align) {
				case 'left':
					paragraphPropXml += ' algn="l"'
					break
				case 'right':
					paragraphPropXml += ' algn="r"'
					break
				case 'center':
					paragraphPropXml += ' algn="ctr"'
					break
				case 'justify':
					paragraphPropXml += ' algn="just"'
					break
				default:
					paragraphPropXml += ''
					break
			}
		}

		if (options.lineSpacing) {
			strXmlLnSpc = `<a:lnSpc><a:spcPts val="${Math.round(options.lineSpacing * 100)}"/></a:lnSpc>`
		} else if (options.lineSpacingMultiple) {
			strXmlLnSpc = `<a:lnSpc><a:spcPct val="${Math.round(options.lineSpacingMultiple * 100000)}"/></a:lnSpc>`
		}

		// OPTION: indent
		if (options.indentLevel && !isNaN(Number(options.indentLevel)) && options.indentLevel > 0) {
			paragraphPropXml += ` lvl="${options.indentLevel}"`
		}

		// OPTION: Paragraph Spacing: Before/After
		if (options.paraSpaceBefore && !isNaN(Number(options.paraSpaceBefore)) && options.paraSpaceBefore > 0) {
			strXmlParaSpc += `<a:spcBef><a:spcPts val="${Math.round(options.paraSpaceBefore * 100)}"/></a:spcBef>`
		}
		if (options.paraSpaceAfter && !isNaN(Number(options.paraSpaceAfter)) && options.paraSpaceAfter > 0) {
			strXmlParaSpc += `<a:spcAft><a:spcPts val="${Math.round(options.paraSpaceAfter * 100)}"/></a:spcAft>`
		}

		// OPTION: bullet
		// NOTE: OOXML uses the unicode character set for Bullets
		// EX: Unicode Character 'BULLET' (U+2022) ==> '<a:buChar char="&#x2022;"/>'
		if (typeof options.bullet === 'object') {
			if (options.bullet?.indent) bulletMarL = valToPts(options.bullet.indent)

			const buPrefix = genXmlBulletPrefix(options.bullet)
			const buIndent = ` marL="${options.indentLevel && options.indentLevel > 0 ? bulletMarL + bulletMarL * options.indentLevel : bulletMarL}" indent="-${bulletMarL}"`

			if (options.bullet.image && options.bullet._imageRId) {
				// `a:buBlip` replaces the character or number bullet entirely
				paragraphPropXml += buIndent
				strXmlBullet = `${buPrefix}<a:buBlip><a:blip r:embed="rId${options.bullet._imageRId}"/></a:buBlip>`
			} else if (options.bullet.type && options.bullet.type.toString().toLowerCase() === 'number') {
				// NOTE: only `type: 'number'` is a distinct branch; any other `type` (e.g. 'bullet') falls through to the char-bullet cases below (issue #1432)
				paragraphPropXml += buIndent
				// auto-numbered bullets fall back to the major latin face when none is given
				strXmlBullet = `${buPrefix}${options.bullet.fontFace ? '' : '<a:buFont typeface="+mj-lt"/>'}<a:buAutoNum type="${options.bullet.numberType || options.bullet.style || 'arabicPeriod'}" startAt="${options.bullet.numberStartAt || options.bullet.startAt || '1'
				}"/>`
			} else if (options.bullet.characterCode) {
				let bulletCode = `&#x${options.bullet.characterCode};`

				// Check value for hex-ness (s/b 4 char hex)
				if (!/^[0-9A-Fa-f]{4}$/.test(options.bullet.characterCode)) {
					console.warn('Warning: `bullet.characterCode should be a 4-digit unicode charatcer (ex: 22AB)`!')
					bulletCode = BULLET_TYPES.DEFAULT
				}

				paragraphPropXml += buIndent
				strXmlBullet = `${buPrefix}<a:buChar char="${bulletCode}"/>`
			} else if (options.bullet.code) {
				// @deprecated `bullet.code` v3.3.0
				let bulletCode = `&#x${options.bullet.code};`

				// Check value for hex-ness (s/b 4 char hex)
				if (!/^[0-9A-Fa-f]{4}$/.test(options.bullet.code)) {
					console.warn('Warning: `bullet.code should be a 4-digit hex code (ex: 22AB)`!')
					bulletCode = BULLET_TYPES.DEFAULT
				}

				paragraphPropXml += buIndent
				strXmlBullet = `${buPrefix}<a:buChar char="${bulletCode}"/>`
			} else {
				paragraphPropXml += buIndent
				strXmlBullet = `${buPrefix}<a:buChar char="${BULLET_TYPES.DEFAULT}"/>`
			}
		} else if (options.bullet) {
			paragraphPropXml += ` marL="${options.indentLevel && options.indentLevel > 0 ? bulletMarL + bulletMarL * options.indentLevel : bulletMarL
			}" indent="-${bulletMarL}"`
			strXmlBullet = `<a:buSzPct val="100000"/><a:buChar char="${BULLET_TYPES.DEFAULT}"/>`
		} else if (!options.bullet) {
			// We only add this when the user explicitely asks for no bullet, otherwise, it can override the master defaults!
			paragraphPropXml += ' indent="0" marL="0"' // FIX: ISSUE#589 - specify zero indent and marL or default will be hanging paragraph
			strXmlBullet = '<a:buNone/>'
		}

		// OPTION: tabStops
		if (options.tabStops && Array.isArray(options.tabStops)) {
			const tabStopsXml = options.tabStops.map(stop => `<a:tab pos="${inch2Emu(stop.position || 1)}" algn="${stop.alignment || 'l'}"/>`).join('')
			strXmlTabStops = `<a:tabLst>${tabStopsXml}</a:tabLst>`
		}

		// B: Close Paragraph-Properties
		// IMPORTANT: strXmlLnSpc, strXmlParaSpc, buClr, and strXmlBullet require strict ordering - anything out of order is ignored. (PPT-Online, PPT for Mac)
		paragraphPropXml += '>' + strXmlLnSpc + strXmlParaSpc + strXmlBullet + strXmlTabStops
		if (isDefault) paragraphPropXml += genXmlTextRunProperties(options, true)
		paragraphPropXml += '</' + tag + '>'
	}

	return paragraphPropXml
}

/**
 * Generate XML Text Run Properties (`a:rPr`)
 * @param {ObjectOptions|TextPropsOptions} opts - text options
 * @param {boolean} isDefault - whether these are the default text run properties
 * @return {string} XML
 */
function genXmlTextRunProperties (opts: ObjectOptions | TextPropsOptions, isDefault: boolean): string {
	let runProps = ''
	const runPropsTag = isDefault ? 'a:defRPr' : 'a:rPr'

	// BEGIN runProperties (ex: `<a:rPr lang="en-US" sz="1600" b="1" dirty="0">`)
	runProps += '<' + runPropsTag + ' lang="' + (opts.lang ? opts.lang : 'en-US') + '"' + (opts.lang ? ' altLang="en-US"' : '')
	runProps += opts.fontSize ? ` sz="${Math.round(opts.fontSize * 100)}"` : '' // NOTE: Use round so sizes like '7.5' wont cause corrupt presentations
	runProps += opts?.bold ? ` b="${opts.bold ? '1' : '0'}"` : ''
	runProps += opts?.italic ? ` i="${opts.italic ? '1' : '0'}"` : ''

	if (opts?.strike === true) warnDeprecatedOnce('strike-boolean', '`strike: true` is deprecated - use `strike: "sngStrike"` (or `"dblStrike"`)')
	runProps += opts?.strike ? ` strike="${typeof opts.strike === 'string' ? opts.strike : 'sngStrike'}"` : ''
	if (typeof opts.underline === 'object' && opts.underline?.style) {
		runProps += ` u="${opts.underline.style}"`
	} else if (typeof opts.underline === 'string') {
		// DEPRECATED: opts.underline is an object as of v3.5.0
		runProps += ` u="${String(opts.underline)}"`
	} else if (opts.hyperlink) {
		runProps += ' u="sng"'
	}
	if (opts.baseline) {
		runProps += ` baseline="${Math.round(opts.baseline * 50)}"`
	} else if (opts.subscript) {
		runProps += ' baseline="-40000"'
	} else if (opts.superscript) {
		runProps += ' baseline="30000"'
	}
	runProps += opts.charSpacing ? ` spc="${Math.round(opts.charSpacing * 100)}" kern="0"` : '' // IMPORTANT: Also disable kerning; otherwise text won't actually expand
	// ECMA-376 §5.1.12.64 ST_TextCapsType: render-only capitalization (issue #1312)
	runProps += opts.caps && ['none', 'small', 'all'].includes(opts.caps) ? ` cap="${opts.caps}"` : ''
	if (opts.kumimoji === true) runProps += ' kumimoji="1"'
	runProps += ' dirty="0">'
	// Color / Font / Highlight / Outline are children of <a:rPr>, so add them now before closing the runProperties tag
	if (opts.color || opts.gradient || opts.fontFace || opts.fontFaceEa || opts.fontFaceCs || opts.outline || (typeof opts.underline === 'object' && opts.underline.color)) {
		if (opts.outline && typeof opts.outline === 'object') {
			runProps += `<a:ln w="${valToPts(opts.outline.size || 0.75)}">${genXmlColorSelection({
				color: opts.outline.color || 'FFFFFF',
				transparency: opts.outline.transparency,
			})}</a:ln>`
		}
		// Run fill: gradient (`a:gradFill`) takes precedence over solid `color` (WordArt / text gradient)
		if (opts.gradient) runProps += genXmlColorSelection({ type: 'gradient', gradient: opts.gradient, transparency: opts.transparency })
		else if (opts.color) runProps += genXmlColorSelection({ color: opts.color, transparency: opts.transparency })
		if (opts.highlight) runProps += `<a:highlight>${createColorElement(opts.highlight)}</a:highlight>`
		if (typeof opts.underline === 'object' && opts.underline.color) runProps += `<a:uFill>${genXmlColorSelection(opts.underline.color)}</a:uFill>`
		const resolvedGlow = resolveGlowOptions(opts.glow)
		if (resolvedGlow) runProps += `<a:effectLst>${createGlowElement(resolvedGlow)}</a:effectLst>`
		const latin = opts.fontFace
		const ea = opts.fontFaceEa || opts.fontFace
		const cs = opts.fontFaceCs || opts.fontFace
		if (latin || ea || cs) {
			// NOTE: 'cs' = Complex Script, 'ea' = East Asian (use "-120" instead of "0" - per Issue #174)
			if (latin) runProps += `<a:latin typeface="${encodeXmlEntities(latin)}" pitchFamily="34" charset="0"/>`
			if (ea) runProps += `<a:ea typeface="${encodeXmlEntities(ea)}" pitchFamily="34" charset="-122"/>`
			if (cs) runProps += `<a:cs typeface="${encodeXmlEntities(cs)}" pitchFamily="34" charset="-120"/>`
		}
	}

	// Hyperlink support
	if (opts.hyperlink) {
		if (typeof opts.hyperlink !== 'object') throw new Error('ERROR: text `hyperlink` option should be an object. Ex: `hyperlink:{url:\'https://github.com\'}` ')
		else if (!opts.hyperlink.url && !opts.hyperlink.slide) throw new Error('ERROR: \'hyperlink requires either `url` or `slide`\'')
		// A colored link keeps its run color instead of the theme's hyperlink color (issue #74: an
		// `a:uFill` here breaks PPT2010, so the extension is used rather than an underline fill)
		const linkColorExt = opts.color
			? '<a:extLst><a:ext uri="{A12FA001-AC4F-418D-AE19-62706E023703}"><ahyp:hlinkClr xmlns:ahyp="http://schemas.microsoft.com/office/drawing/2018/hyperlinkcolor" val="tx"/></a:ext></a:extLst>'
			: ''
		runProps += genXmlHyperlink(opts.hyperlink, 'click', 'run', linkColorExt)
	}
	// `a:hlinkMouseOver` follows `a:hlinkClick` in CT_TextCharacterProperties
	if (opts.hyperlinkHover?._rId) runProps += genXmlHyperlink(opts.hyperlinkHover, 'hover', 'run')

	// END runProperties
	runProps += `</${runPropsTag}>`

	return runProps
}

/**
 * Build textBody text runs [`<a:r></a:r>`] for paragraphs [`<a:p>`]
 * @param {TextProps} textObj - Text object
 * @return {string} XML string
 */
/** DrawingML 2010 math zone wrapper — required for PowerPoint (bare `m:oMath` is stripped on open). */
export const A14_NS = 'http://schemas.microsoft.com/office/drawing/2010/main'
export const MATH_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/math'
const W_NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
export const MC_NS = 'http://schemas.openxmlformats.org/markup-compatibility/2006'
export const P14_NS = 'http://schemas.microsoft.com/office/powerpoint/2010/main'
export const P1710_NS = 'http://schemas.microsoft.com/office/powerpoint/2017/10/main'

/**
 * Normalize caller-supplied OMML for PowerPoint.
 * - Ensures an `m:oMath` / `m:oMathPara` root (inner fragments are wrapped).
 * - Wraps in `<a14:m>` — PowerPoint silently drops bare `m:oMath` siblings of `a:r`.
 * @see MS-ODRAWXML Math / mc:AlternateContent; proven via PowerPoint 16 COM round-trip.
 */
function normalizeOmml (omml: string): string {
	let trimmed = omml.trim()
	if (!trimmed) return ''

	// Malformed OMML would be embedded verbatim and corrupt the whole package
	// (PowerPoint repair prompt / open failure), so fail fast with a pointer.
	const problem = validateXmlFragment(trimmed)
	if (problem) {
		throw new Error(`ERROR: text run 'omml' option is not a well-formed XML fragment: ${problem}. Escape all '<', '&' in math text (e.g. <m:t>a&lt;b</m:t>).`)
	}

	// Already PowerPoint-wrapped
	if (/^<a14:m[\s/>]/i.test(trimmed)) return trimmed

	// Ensure oMath / oMathPara root
	if (!/^<m:oMath[\s/>]/i.test(trimmed) && !/^<m:oMathPara[\s/>]/i.test(trimmed)) {
		trimmed = `<m:oMath xmlns:m="${MATH_NS}" xmlns:w="${W_NS}">${trimmed}</m:oMath>`
	} else if (!/xmlns:m=/i.test(trimmed.slice(0, 180))) {
		trimmed = trimmed.replace(/^(<m:oMathPara|<m:oMath)/i, `$1 xmlns:m="${MATH_NS}" xmlns:w="${W_NS}"`)
	}

	return `<a14:m xmlns:a14="${A14_NS}">${trimmed}</a14:m>`
}

/**
 * A stable GUID for a field, derived from its type and cached value
 * - `a:fld@id` must be a GUID; deriving it keeps repeated exports byte-identical
 */
function fieldGuid (type: string, cached: string): string {
	let hash = 0
	for (const char of `${type}:${cached}`) hash = (hash * 31 + char.charCodeAt(0)) >>> 0
	const hex = hash.toString(16).padStart(8, '0')
	return `{${hex}-0000-0000-0000-${hex}0000}`
}

export function textRunsHaveOmml (text: TextProps[] | string | undefined): boolean {
	if (!Array.isArray(text)) return false
	return text.some(t => typeof t.options?.omml === 'string' && t.options.omml.trim().length > 0)
}

function genXmlTextRun (textObj: TextProps): string {
	// Native Office Math run — editable in PowerPoint (not an image).
	const omml = textObj.options?.omml
	if (typeof omml === 'string' && omml.trim()) {
		return normalizeOmml(omml)
	}

	// NOTE: Dont create full rPr runProps for empty [lineBreak] runs
	// Why? The size of the lineBreak wont match (eg: below it will be 18px instead of the correct 36px)
	// Do this:
	/*
		<a:p>
			<a:pPr algn="r"/>
			<a:endParaRPr lang="en-US" sz="3600" dirty="0"/>
		</a:p>
	*/
	// NOT this:
	/*
		<a:p>
			<a:pPr algn="r"/>
			<a:r>
				<a:rPr lang="en-US" sz="3600" dirty="0">
					<a:solidFill>
						<a:schemeClr val="accent5"/>
					</a:solidFill>
					<a:latin typeface="Times" pitchFamily="34" charset="0"/>
					<a:ea typeface="Times" pitchFamily="34" charset="-122"/>
					<a:cs typeface="Times" pitchFamily="34" charset="-120"/>
				</a:rPr>
				<a:t></a:t>
			</a:r>
			<a:endParaRPr lang="en-US" dirty="0"/>
		</a:p>
	*/

	/* A field is a run the consumer refreshes on open (ECMA-376 21.1.2.2.4). The `a:t` child holds
	 * the cached value so a consumer that does not refresh still renders something. `id` must be a
	 * GUID; it is derived from the field type and text so repeated exports stay reproducible. */
	const field = textObj.options?.field
	if (field) {
		if (!TEXT_FIELD_TYPES.has(field)) {
			console.warn(`[pptxgenjs] unknown text field "${String(field)}" - run emitted as plain text instead`)
		} else {
			return (
				`<a:fld id="${fieldGuid(field, textObj.text ?? '')}" type="${field}">` +
				genXmlTextRunProperties(textObj.options ?? {}, false) +
				`<a:t>${encodeXmlEntities(textObj.text ?? '')}</a:t>` +
				'</a:fld>'
			)
		}
	}

	// Return paragraph with text run
	return textObj.text ? `<a:r>${genXmlTextRunProperties(textObj.options ?? {}, false)}<a:t>${encodeXmlEntities(textObj.text)}</a:t></a:r>` : ''
}

/**
 * Builds `<a:bodyPr></a:bodyPr>` tag for "genXmlTextBody()"
 * @param {ISlideObject | TableCell} slideObject - various options
 * @return {string} XML string
 */
function genXmlBodyProperties (slideObject: ISlideObject | TableCell): string {
	let bodyProperties = '<a:bodyPr'

	// Placeholders need the same bodyPr (valign/margin/wrap) as text — mikemeerschaert/fix-placeholder-text-formatting-issues
	if (
		slideObject &&
		(slideObject._type === SLIDE_OBJECT_TYPES.text || slideObject._type === SLIDE_OBJECT_TYPES.placeholder) &&
		slideObject.options?._bodyProp
	) {
		// PPT-2019 EX: <a:bodyPr wrap="square" lIns="1270" tIns="1270" rIns="1270" bIns="1270" rtlCol="0" anchor="ctr"/>

		// A: Enable or disable textwrapping none or square
		bodyProperties += slideObject.options._bodyProp.wrap ? ' wrap="square"' : ' wrap="none"'
		if (slideObject.options._bodyProp.vertOverflow) bodyProperties += ` vertOverflow="${slideObject.options._bodyProp.vertOverflow}"`
		if (slideObject.options._bodyProp.horzOverflow) bodyProperties += ` horzOverflow="${slideObject.options._bodyProp.horzOverflow}"`

		// B: Textbox margins [padding]
		if (slideObject.options._bodyProp.lIns || slideObject.options._bodyProp.lIns === 0) bodyProperties += ` lIns="${slideObject.options._bodyProp.lIns}"`
		if (slideObject.options._bodyProp.tIns || slideObject.options._bodyProp.tIns === 0) bodyProperties += ` tIns="${slideObject.options._bodyProp.tIns}"`
		if (slideObject.options._bodyProp.rIns || slideObject.options._bodyProp.rIns === 0) bodyProperties += ` rIns="${slideObject.options._bodyProp.rIns}"`
		if (slideObject.options._bodyProp.bIns || slideObject.options._bodyProp.bIns === 0) bodyProperties += ` bIns="${slideObject.options._bodyProp.bIns}"`

		// Text columns — ECMA-376 §5.1.5.1.4 CT_TextBodyProperties@numCol/@spcCol (issue #1320)
		if (slideObject.options._bodyProp.numCol && slideObject.options._bodyProp.numCol > 1) {
			bodyProperties += ` numCol="${slideObject.options._bodyProp.numCol}"`
			if (slideObject.options._bodyProp.spcCol) bodyProperties += ` spcCol="${slideObject.options._bodyProp.spcCol}"`
		}

		// `rtlCol` follows the text box's RTL mode unless overridden, so RTL columns flow correctly
		const rtlCols = typeof slideObject.options.rtlColumns === 'boolean' ? slideObject.options.rtlColumns : slideObject.options.rtlMode === true
		bodyProperties += ` rtlCol="${rtlCols ? '1' : '0'}"`

		// D: Add anchorPoints
		if (slideObject.options._bodyProp.anchor) bodyProperties += ' anchor="' + slideObject.options._bodyProp.anchor + '"' // VALS: [t,ctr,b]
		if (slideObject.options._bodyProp.vert) bodyProperties += ' vert="' + slideObject.options._bodyProp.vert + '"' // VALS: [eaVert,horz,mongolianVert,vert,vert270,wordArtVert,wordArtVertRtl]

		// E: Close <a:bodyPr element
		bodyProperties += '>'

		// E.1: WordArt text warp (`a:prstTxWarp`) — first child of CT_TextBodyProperties, before autofit
		if (slideObject.options.presetShape && TEXT_SHAPE_TYPES.has(slideObject.options.presetShape)) {
			bodyProperties += `<a:prstTxWarp prst="${slideObject.options.presetShape}"><a:avLst/></a:prstTxWarp>`
		}

		/**
		 * F: Text Fit/AutoFit/Shrink option
		 * @see: http://officeopenxml.com/drwSp-text-bodyPr-fit.php
		 * @see: http://www.datypic.com/sc/ooxml/g-a_EG_TextAutofit.html
		 */
		if (slideObject.options.fit) {
			// NOTE: Use of '<a:noAutofit/>' instead of '' causes issues in PPT-2013!
			if (slideObject.options.fit === 'none') bodyProperties += ''
			// NOTE: Shrink does not work automatically - PowerPoint calculates the `fontScale` value dynamically upon resize
			else if (slideObject.options.fit === 'shrink') bodyProperties += '<a:normAutofit/>'
			// Object form: explicit fontScale/lnSpcReduction percentages (ECMA-376 §5.1.5.1.3, issue #1199).
			// Values are 0-100%; normAutofit stores 1000ths of a percent (100000 = 100%).
			else if (typeof slideObject.options.fit === 'object' && slideObject.options.fit?.type === 'shrink') {
				const fs = slideObject.options.fit.fontScale
				const ls = slideObject.options.fit.lnSpcReduction
				bodyProperties += `<a:normAutofit${typeof fs === 'number' ? ` fontScale="${Math.round(fs * 1000)}"` : ''}${typeof ls === 'number' ? ` lnSpcReduction="${Math.round(ls * 1000)}"` : ''}/>`
			}
			else if (slideObject.options.fit === 'resize') bodyProperties += '<a:spAutoFit/>'
		}
		//
		// DEPRECATED: below (@deprecated v3.3.0)
		if (slideObject.options.shrinkText) bodyProperties += '<a:normAutofit/>' // MS-PPT > Format shape > Text Options: "Shrink text on overflow"
		/* DEPRECATED: below (@deprecated v3.3.0)
		 * MS-PPT > Format shape > Text Options: "Resize shape to fit text" [spAutoFit]
		 * NOTE: Use of '<a:noAutofit/>' in lieu of '' below causes issues in PPT-2013
		 */
		bodyProperties += slideObject.options._bodyProp.autoFit ? '<a:spAutoFit/>' : ''

		// LAST: Close _bodyProp
		bodyProperties += '</a:bodyPr>'
	} else {
		// DEFAULT:
		const opts = slideObject.options as ObjectOptions | undefined
		const rtlCols = typeof opts?.rtlColumns === 'boolean' ? opts.rtlColumns : opts?.rtlMode === true
		bodyProperties += ` wrap="square" rtlCol="${rtlCols ? '1' : '0'}">`
		bodyProperties += '</a:bodyPr>'
	}

	// LAST: Return Close _bodyProp
	return slideObject._type === SLIDE_OBJECT_TYPES.tablecell ? '<a:bodyPr/>' : bodyProperties
}

/**
 * Generate the XML for text and its options (bold, bullet, etc) including text runs (word-level formatting)
 * @param {ISlideObject|TableCell} slideObj - slideObj or tableCell
 * @note PPT text lines [lines followed by line-breaks] are created using <p>-aragraph's
 * @note Bullets are a paragragh-level formatting device
 * @template
 *    <p:txBody>
 *        <a:bodyPr wrap="square" rtlCol="0">
 *            <a:spAutoFit/>
 *        </a:bodyPr>
 *        <a:lstStyle/>
 *        <a:p>
 *            <a:pPr algn="ctr"/>
 *            <a:r>
 *                <a:rPr lang="en-US" dirty="0" err="1"/>
 *                <a:t>textbox text</a:t>
 *            </a:r>
 *            <a:endParaRPr lang="en-US" dirty="0"/>
 *        </a:p>
 *    </p:txBody>
 * @returns XML containing the param object's text and formatting
 */
export function genXmlTextBody (slideObj: ISlideObject | TableCell): string {
	const opts: ObjectOptions = slideObj.options || {}
	let tmpTextObjects: TextProps[] = []
	const arrTextObjects: TextProps[] = []

	// FIRST: Shapes without text still require a `<p:txBody>` child on `<p:sp>` per OOXML; returning '' omits it and triggers PowerPoint repair
	if (opts && slideObj._type !== SLIDE_OBJECT_TYPES.tablecell && (typeof slideObj.text === 'undefined' || slideObj.text === null)) {
		return `<p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:endParaRPr lang="${opts.lang || 'en-US'}"/></a:p></p:txBody>`
	}

	// STEP 1: Start textBody
	let strSlideXml = slideObj._type === SLIDE_OBJECT_TYPES.tablecell ? '<a:txBody>' : '<p:txBody>'

	// STEP 2: Add bodyProperties
	{
		// A: 'bodyPr'
		strSlideXml += genXmlBodyProperties(slideObj)

		// B: 'lstStyle'
		// NOTE: shape type 'LINE' has different text align needs (a lstStyle.lvl1pPr between bodyPr and p)
		// FIXME: LINE horiz-align doesnt work (text is always to the left inside line) (FYI: the PPT code diff is substantial!)
		if (opts.h === 0 && opts.line && opts.align) strSlideXml += '<a:lstStyle><a:lvl1pPr algn="l"/></a:lstStyle>'
		else if (slideObj._type === 'placeholder') strSlideXml += `<a:lstStyle>${genXmlParagraphProperties(slideObj, true)}</a:lstStyle>`
		else strSlideXml += '<a:lstStyle/>'
	}

	/* STEP 3: Modify slideObj.text to array
		CASES:
		addText( 'string' ) // string
		addText( 'line1\n line2' ) // string with lineBreak
		addText( {text:'word1'} ) // TextProps object
		addText( ['barry','allen'] ) // array of strings
		addText( [{text:'word1'}, {text:'word2'}] ) // TextProps object array
		addText( [{text:'line1\n line2'}, {text:'end word'}] ) // TextProps object array with lineBreak
	*/
	if (typeof slideObj.text === 'string' || typeof slideObj.text === 'number') {
		// Handle cases 1,2
		tmpTextObjects.push({ text: slideObj.text.toString(), options: opts || {} })
	} else if (slideObj.text && !Array.isArray(slideObj.text) && typeof slideObj.text === 'object' && Object.keys(slideObj.text).includes('text')) {
		// } else if (!Array.isArray(slideObj.text) && slideObj.text!.hasOwnProperty('text')) { // 20210706: replaced with below as ts compiler rejected it
		// Handle case 3
		tmpTextObjects.push({ text: slideObj.text || '', options: slideObj.options || {} })
	} else if (Array.isArray(slideObj.text)) {
		// Handle cases 4,5,6
		tmpTextObjects = slideObj.text.map(item => ({ text: item.text, options: item.options }))
	}

	// STEP 4: Iterate over text objects, set text/options, break into pieces if '\n'/breakLine found
	tmpTextObjects.forEach((itext, idx) => {
		if (!itext.text) itext.text = ''

		// A: Set options. `addText(string, opts)` and unstyled runs share the shape `opts`
		// object; glow/softEdge/reflection/shadow belong on the shape `effectLst` (issue #84).
		if (!itext.options || itext.options === opts) {
			const runOpts = { ...opts }
			delete runOpts.glow
			delete runOpts.softEdge
			delete runOpts.reflection
			delete runOpts.shadow
			delete runOpts.fillOverlay
			delete runOpts.effectDag
			itext.options = runOpts
		}
		if (idx === 0 && itext.options && !itext.options.bullet && opts.bullet) itext.options.bullet = opts.bullet

		// B: Cast to text-object and fix line-breaks (if needed)
		if (typeof itext.text === 'string' || typeof itext.text === 'number') {
			// 1: Convert "\n" or any variation into CRLF
			itext.text = itext.text.toString().replace(/\r*\n/g, CRLF)
		}

		// C: If text string has line-breaks, then create a separate text-object for each (much easier than dealing with split inside a loop below)
		// NOTE: Filter for trailing lineBreak prevents the creation of an empty textObj as the last item
		// Martin-N: do not mutate shared `options` when splitting CRLF — only non-final lines get breakLine
		if (itext.text.includes(CRLF) && itext.text.match(/\n$/g) === null) {
			const lines = itext.text.split(CRLF)
			lines.forEach((line, index) => {
				if (index === lines.length - 1) {
					arrTextObjects.push({ text: line, options: itext.options })
				} else {
					arrTextObjects.push({ text: line, options: { ...itext.options, breakLine: true } })
				}
			})
		} else {
			arrTextObjects.push(itext)
		}
	})

	// STEP 5: Group textObj into lines by checking for lineBreak, bullets, alignment change, etc.
	const arrLines: TextProps[][] = []
	let arrTexts: TextProps[] = []
	arrTextObjects.forEach((textObj, idx) => {
		// OMML-only runs use empty `text` but must still emit content
		if (!textObj.text && textObj.options?.omml) textObj.text = ''

		// A: Align or Bullet trigger new line
		if (arrTexts.length > 0 && (textObj.options?.align || opts.align)) {
			// Only start a new paragraph when align *changes*
			if (textObj.options?.align !== arrTextObjects[idx - 1].options?.align) {
				arrLines.push(arrTexts)
				arrTexts = []
			}
		} else if (arrTexts.length > 0 && textObj.options?.bullet && arrTexts.length > 0) {
			arrLines.push(arrTexts)
			arrTexts = []
			if (textObj.options) textObj.options.breakLine = false // For cases with both `bullet` and `brekaLine` - prevent double lineBreak
		}

		// B: Add this text to current line
		arrTexts.push(textObj)

		// C: BreakLine begins new line **after** adding current text
		if (arrTexts.length > 0 && textObj.options?.breakLine) {
			// Avoid starting a para right as loop is exhausted
			if (idx + 1 < arrTextObjects.length) {
				arrLines.push(arrTexts)
				arrTexts = []
			}
		}

		// D: Flush buffer
		if (idx + 1 === arrTextObjects.length) arrLines.push(arrTexts)
	})

	// STEP 6: Loop over each line and create paragraph props, text run, etc.
	arrLines.forEach(line => {
		let reqsClosingFontSize = false

		// A: Start paragraph
		strSlideXml += '<a:p>'

		// A.1: Emit paragraph properties ONCE, derived from the first run.
		// OOXML (ECMA-376 CT_TextParagraph) allows at most one `a:pPr` and it must be the
		// first child of `a:p`. Emitting a `pPr` before every run (legacy behavior) produces
		// schema-invalid XML that PowerPoint silently tolerated for text-only paragraphs,
		// but REPAIRS DESTRUCTIVELY (dropping runs/equations) once `m:oMath` is present.
		{
			const firstOpts = line[0].options ?? (line[0].options = {})
			firstOpts._lineIdx = 0
			firstOpts.align = firstOpts.align || opts.align
			firstOpts.lineSpacing = firstOpts.lineSpacing || opts.lineSpacing
			firstOpts.lineSpacingMultiple = firstOpts.lineSpacingMultiple || opts.lineSpacingMultiple
			firstOpts.indentLevel = firstOpts.indentLevel || opts.indentLevel
			firstOpts.paraSpaceBefore = firstOpts.paraSpaceBefore || opts.paraSpaceBefore
			firstOpts.paraSpaceAfter = firstOpts.paraSpaceAfter || opts.paraSpaceAfter
			const paragraphPropXml = genXmlParagraphProperties(line[0], false)
			strSlideXml += paragraphPropXml.replace('<a:pPr></a:pPr>', '') // IMPORTANT: Empty "pPr" blocks will generate needs-repair/corrupt msg
		}

		// B: Loop over line runs and add text runs
		line.forEach((textObj, idx) => {
			// NOTE: `options` is always populated by the time text is serialized; bind the real object so mutations below persist
			const textOpts = textObj.options ?? (textObj.options = {})
			// A: Set line index
			textOpts._lineIdx = idx

			// A.1: Add soft break if not the first run of the line.
			if (idx > 0 && textOpts.softBreakBefore) {
				strSlideXml += '<a:br/>'
			}

			// B: Inherit pPr-type options from parent shape's `options`
			textOpts.align = textOpts.align || opts.align
			textOpts.lineSpacing = textOpts.lineSpacing || opts.lineSpacing
			textOpts.lineSpacingMultiple = textOpts.lineSpacingMultiple || opts.lineSpacingMultiple
			textOpts.indentLevel = textOpts.indentLevel || opts.indentLevel
			textOpts.paraSpaceBefore = textOpts.paraSpaceBefore || opts.paraSpaceBefore
			textOpts.paraSpaceAfter = textOpts.paraSpaceAfter || opts.paraSpaceAfter

			// C: Inherit any main options (color, fontSize, etc.)
			// NOTE: We only pass the text.options to genXmlTextRun (not the Slide.options),
			// so the run building function cant just fallback to Slide.color, therefore, we need to do that here before passing options below.
			// FILTER RULE: Hyperlinks should not inherit `color` from main options (let PPT default to local color, eg: blue on MacOS)
			Object.entries(opts).filter(([key]) => !(textOpts.hyperlink && key === 'color')).forEach(([key, val]) => {
				// if (textOpts.hyperlink && key === 'color') null
				// NOTE: This loop will pick up unecessary keys (`x`, etc.), but it doesnt hurt anything
				// `glow` / softEdge / reflection belong on the shape effectLst (not per-run), unless set on the run itself
				// `omml` is a run-level math payload — inheriting it would replace sibling plain-text runs
				if (key === 'glow' || key === 'softEdge' || key === 'reflection' || key === 'shadow' || key === 'blur' || key === 'fillOverlay' || key === 'effectDag' || key === 'omml') return
				if (key !== 'bullet' && !textOpts[key]) textOpts[key] = val
			})

			// D: Add formatted textrun
			strSlideXml += genXmlTextRun(textObj)

			// E: Flag close fontSize for empty [lineBreak] elements
			if ((!textObj.text && opts.fontSize) || textOpts.fontSize) {
				reqsClosingFontSize = true
				opts.fontSize = opts.fontSize || textOpts.fontSize
			}
		})

		/* C: Append 'endParaRPr' (when needed) and close current open paragraph
		 * NOTE: (ISSUE#20, ISSUE#193): Add 'endParaRPr' with font/size props or PPT default (Arial/18pt en-us) is used making row "too tall"/not honoring options
		 */
		if (slideObj._type === SLIDE_OBJECT_TYPES.tablecell && (opts.fontSize || opts.fontFace || opts.fontFaceEa || opts.fontFaceCs)) {
			const latin = opts.fontFace
			const ea = opts.fontFaceEa || opts.fontFace
			const cs = opts.fontFaceCs || opts.fontFace
			if (latin || ea || cs) {
				strSlideXml += `<a:endParaRPr lang="${opts.lang || 'en-US'}"` + (opts.fontSize ? ` sz="${Math.round(opts.fontSize * 100)}"` : '') + ' dirty="0">'
				if (latin) strSlideXml += `<a:latin typeface="${encodeXmlEntities(latin)}" charset="0"/>`
				if (ea) strSlideXml += `<a:ea typeface="${encodeXmlEntities(ea)}" charset="0"/>`
				if (cs) strSlideXml += `<a:cs typeface="${encodeXmlEntities(cs)}" charset="0"/>`
				strSlideXml += '</a:endParaRPr>'
			} else {
				strSlideXml += `<a:endParaRPr lang="${opts.lang || 'en-US'}"` + (opts.fontSize ? ` sz="${Math.round(opts.fontSize * 100)}"` : '') + ' dirty="0"/>'
			}
		} else if (reqsClosingFontSize) {
			// Empty [lineBreak] lines should not contain runProp, however, they need to specify fontSize in `endParaRPr`
			strSlideXml += `<a:endParaRPr lang="${opts.lang || 'en-US'}"` + (opts.fontSize ? ` sz="${Math.round(opts.fontSize * 100)}"` : '') + ' dirty="0"/>'
		} else {
			strSlideXml += `<a:endParaRPr lang="${opts.lang || 'en-US'}" dirty="0"/>` // Added 20180101 to address PPT-2007 issues
		}

		// D: End paragraph
		strSlideXml += '</a:p>'
	})

	// IMPORTANT: An empty txBody will cause "needs repair" error! Add <p> content if missing.
	// [FIXED in v3.13.0]: This fixes issue with table auto-paging where some cells w/b empty on subsequent pages.
	/*
		<a:txBody>
			<a:bodyPr/>
			<a:lstStyle/>
		</a:txBody>
	*/
	if (strSlideXml.indexOf('<a:p>') === -1) {
		strSlideXml += '<a:p><a:endParaRPr/></a:p>'
	}

	// STEP 7: Close the textBody
	strSlideXml += slideObj._type === SLIDE_OBJECT_TYPES.tablecell ? '</a:txBody>' : '</p:txBody>'

	// LAST: Return XML
	return strSlideXml
}

const CLASSIFICATION_OUTCOMES = new Set<ClassificationOutcome>(['none', 'hdr', 'ftr', 'watermark'])

const URI_DESIGN_ELEM = '{386F3935-93C4-4BCD-93E2-E3B085C9AB24}'
const URI_CLASSIFICATION = '{1162E1C5-73C7-4A58-AE30-91384D911F3F}'
const URI_DESIGN_PR = '{E7BDC344-281C-4309-B0C6-D0EE65EED2A8}'
export const URI_DESIGN_TAG_LST = '{E3EDB536-0D56-4F60-86BA-61A60CA02DAB}'
const URI_PH_TYPE_EXT = '{56F484CC-4922-43CF-B6FB-B326C6A72FC8}'

const NS_P16_2015 = 'http://schemas.microsoft.com/office/powerpoint/2015/main'
const NS_P184 = 'http://schemas.microsoft.com/office/powerpoint/2018/4/main'
const NS_P202 = 'http://schemas.microsoft.com/office/powerpoint/2020/02/main'
const NS_P232 = 'http://schemas.microsoft.com/office/powerpoint/2023/02/main'

/** Designer Service tag list (`CT_DesignerTagList`). Empty when no tags. */
export function genXmlDesignTagLst (tags?: DesignerTag[], withNs = false): string {
	if (!tags || tags.length === 0) return ''
	const items = tags
		.filter(t => t && typeof t.name === 'string' && typeof t.val === 'string')
		.map(t => `<p202:designTag name="${encodeXmlEntities(t.name)}" val="${encodeXmlEntities(t.val)}"/>`)
		.join('')
	if (!items) return ''
	const ns = withNs ? ` xmlns:p202="${NS_P202}"` : ''
	return `<p202:designTagLst${ns}>${items}</p202:designTagLst>`
}

/**
 * Opt-in MS-PPTX nvPr extensions: designElem, classification, designPr.
 * Classification is emitted only when explicitly set to a valid outcome token.
 */
export function genXmlNvPrDesignExts (options?: ObjectOptions): string {
	if (!options) return ''
	const exts: string[] = []

	if (options.designElem === true) {
		exts.push(
			`<p:ext uri="${URI_DESIGN_ELEM}"><p16:designElem xmlns:p16="${NS_P16_2015}" val="1"/></p:ext>`
		)
	}

	if (typeof options.classification === 'string' && CLASSIFICATION_OUTCOMES.has(options.classification)) {
		exts.push(
			`<p:ext uri="${URI_CLASSIFICATION}"><p184:classification xmlns:p184="${NS_P184}" val="${options.classification}"/></p:ext>`
		)
	}

	if (options.designPr) {
		const edt = options.designPr.edtDesignElem === true ? ' edtDesignElem="1"' : ''
		const tags = genXmlDesignTagLst(options.designPr.tags)
		exts.push(
			`<p:ext uri="${URI_DESIGN_PR}"><p202:designPr xmlns:p202="${NS_P202}"${edt}>${tags}</p202:designPr></p:ext>`
		)
	}

	return exts.join('')
}

/** Wrap design + extra `p:ext` children in `p:extLst`, or return empty. */
export function genXmlNvPrExtLst (options?: ObjectOptions, extraExts: string[] = []): string {
	const body = extraExts.join('') + genXmlNvPrDesignExts(options)
	return body ? `<p:extLst>${body}</p:extLst>` : ''
}

/**
 * Resolve a placeholder type to the ST_PlaceholderType token written on `p:ph@type`.
 * ECMA-376 §4.4.1.33 / CT_Placeholder: `type` is `use="optional" default="obj"`
 * (`standards/ecma/part-20_presentationml-reference-material-slides.txt`).
 * Omitting the attribute is therefore the generic object placeholder, not a picture/table slot.
 *
 * `PlaceholderProps.type` is the OOXML token ('pic', 'tbl', …) while PLACEHOLDER_TYPES is keyed
 * by friendly name ('image', 'table'). Accept both so `type: 'pic'` and `PLACEHOLDER_TYPES.image`
 * emit `type="pic"`. Port of gitbrent/PptxGenJS#1526; the prior double lookup dropped pic/tbl.
 * @param {string} type placeholder type as provided by the user
 * @returns OOXML placeholder type, or an empty string when unrecognized
 */
function resolvePlaceholderType (type: string | undefined): string {
	if (!type) return ''

	// A) OOXML value ('pic'): the documented form of `PlaceholderProps.type`
	const ooxmlTypes: string[] = Object.values(PLACEHOLDER_TYPES)
	if (ooxmlTypes.includes(type)) return type

	// B) Enum key ('image'): tolerated so untyped (JS) callers keep working
	return PLACEHOLDER_TYPES[type as keyof typeof PLACEHOLDER_TYPES]?.toString() ?? ''
}

/**
 * Generate an XML Placeholder
 * @param {ISlideObject} placeholderObj
 * @param {ObjectOptions} extraOpts slide-object options (phTypeExt override)
 * @returns XML
 */
export function genXmlPlaceholder (placeholderObj: ISlideObject | undefined, extraOpts?: ObjectOptions): string {
	if (!placeholderObj) return ''

	const placeholderIdx = placeholderObj.options?._placeholderIdx ? placeholderObj.options._placeholderIdx : ''
	const placeholderType = resolvePlaceholderType(placeholderObj.options?._placeholderType)

	// `orient` and `sz` default to horz/full in the schema, so only a non-default is written
	const placeholderOrient = placeholderObj.options?.orient === 'vert' ? ' orient="vert"' : ''
	const placeholderSize = placeholderObj.options?.sz === 'half' || placeholderObj.options?.sz === 'quarter' ? ` sz="${placeholderObj.options.sz}"` : ''

	const attrs =
		`${placeholderIdx ? ' idx="' + placeholderIdx.toString() + '"' : ''}` +
		`${placeholderType ? ` type="${placeholderType}"` : ''}` +
		placeholderOrient +
		placeholderSize +
		`${placeholderObj.text && placeholderObj.text.length > 0 ? ' hasCustomPrompt="1"' : ''}`

	const phTypeExt = extraOpts?.phTypeExt ?? placeholderObj.options?.phTypeExt
	if (phTypeExt === 'cameo' || phTypeExt === 'unknown') {
		return (
			`<p:ph${attrs}>` +
			`<p:extLst><p:ext uri="${URI_PH_TYPE_EXT}">` +
			`<p232:phTypeExt xmlns:p232="${NS_P232}"><p232:type><p232:${phTypeExt}/></p232:type></p232:phTypeExt>` +
			'</p:ext></p:extLst></p:ph>'
		)
	}

	return `<p:ph${attrs}/>`
}
