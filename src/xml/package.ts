/**
 * OOXML package-part rendering.
 */

import { CHART_STYLE, CRLF, DEF_COLOR_MAP, EMU, LAYOUT_IDX_SERIES_BASE, OOXML_CHARTEX, SLDNUMFLDID, SLIDE_OBJECT_TYPES, isChartexType } from '../core-enums'
import {
	AnimationConfig,
	AnimationType,
	ColorMapOverrideProps,
	DocumentProps,
	GuideProps,
	IPresentationProps,
	ISlideRelChart,
	PresSlide,
	SectionProps,
	SlideLayout,
	SlideObjectAnimation,
	SlideShowEvent,
	SlideShowProps,
	TableStyleBorderProps,
	TableStylePartProps,
	TableStyleProps,
	TextProps,
} from '../core-interfaces'
import { chartColorsPartName, chartStylePartName, wantsChartStyleParts } from '../charts/style'
import { createTimingXml, MediaPlaybackEntry } from '../gen-animations'
import { genXmlTransition } from '../gen-transition'
import { AUTHOR_PART_CONTENT_TYPE, AUTHOR_REL_TYPE, COMMENT_PART_CONTENT_TYPE, COMMENT_REL_URI, P188_NS } from '../gen-comments'
import { createColorElement, encodeXmlEntities, genXmlColorSelection, getUuid, inch2Emu, resolveThemeColors } from '../gen-utils'
import { extPartPackagePath } from './content-parts'
import { slideCommentsRelId } from './relationships'
import { genXmlLine } from './line'
import { resolveZoomSections, slideObjectToXml } from './slide'
import { A14_NS, genXmlDesignTagLst, MATH_NS, MC_NS, P14_NS, P1710_NS, URI_DESIGN_TAG_LST } from './text'
import {
	CHANGES_INFO_CONTENT_TYPE,
	CHANGES_INFO_REL_TYPE,
	REVISION_INFO_CONTENT_TYPE,
	REVISION_INFO_REL_TYPE,
} from '../gen-revision'

/**
 * Content-type overrides a chart needs: the chart itself plus, when requested, its style and colour-style parts.
 * - one helper because charts appear on slides, layouts and the master, and a part with no declared
 *   content type is a repair-dialog cause
 */
function chartContentTypes (rel: ISlideRelChart): string {
	const partContentType = isChartexType(rel.opts._type) ? OOXML_CHARTEX.partContentType : 'application/vnd.openxmlformats-officedocument.drawingml.chart+xml'
	let xml = `<Override PartName="${rel.Target}" ContentType="${partContentType}"/>`
	if (wantsChartStyleParts(rel.opts)) {
		xml += `<Override PartName="/${chartColorsPartName(rel.globalId)}" ContentType="${CHART_STYLE.colorsContentType}"/>`
		xml += `<Override PartName="/${chartStylePartName(rel.globalId)}" ContentType="${CHART_STYLE.styleContentType}"/>`
	}
	return xml
}

/** Opt-in MS-PPTX presentation parts (zero-or-one each). */
export type PresentationTrackingParts = {
	revisionInfo?: boolean
	changesInfo?: boolean
}

/** MS-PPTX §2.2.6 showPr browseMode */
const URI_BROWSE_MODE = '{F99C55AA-B7CB-42B0-86F8-08522FDF87E8}'
/** MS-PPTX §2.2.6 showPr laserClr */
const URI_LASER_CLR = '{EC167BDD-8182-4AB7-AECC-EB403E3ABB37}'
/** MS-PPTX §2.2.6 sld laserTraceLst */
const URI_LASER_TRACE_LST = '{3A86A75C-4F4B-4683-9AE1-C65F6400EC91}'
/** MS-PPTX §2.2.6 sld showEvtLst */
const URI_SHOW_EVT_LST = '{E180D4A7-C9FB-4DFB-919C-405C955672EB}'
/** MS-PPTX §2.2.7 presentationPr defaultImageDpi */
const URI_DEFAULT_IMAGE_DPI = '{D31A062A-798A-4329-ABDD-BBA856620510}'
/** MS-PPTX §2.2.7 presentationPr discardImageEditData */
const URI_DISCARD_IMAGE_EDIT_DATA = '{E76CE94A-603C-4142-B9EB-6D1370010A27}'
/** MS-PPTX §2.2.16 presentationPr readonlyRecommended */
const URI_READONLY_RECOMMENDED = '{1BD7E111-0CB8-44D6-8891-C1BB2F81B7CC}'

function showEventXml (evt: SlideShowEvent): string {
	const time = Math.round(evt.time)
	const objId = Math.round(evt.objId)
	switch (evt.type) {
		case 'trigger':
			return `<p14:triggerEvt type="${evt.trigger ?? 'onClick'}" time="${time}" objId="${objId}"/>`
		case 'play':
			return `<p14:playEvt time="${time}" objId="${objId}"/>`
		case 'stop':
			return `<p14:stopEvt time="${time}" objId="${objId}"/>`
		case 'pause':
			return `<p14:pauseEvt time="${time}" objId="${objId}"/>`
		case 'resume':
			return `<p14:resumeEvt time="${time}" objId="${objId}"/>`
		case 'seek':
			return `<p14:seekEvt time="${time}" objId="${objId}" seek="${Math.round(evt.seek ?? 0)}"/>`
		case 'null':
			return `<p14:nullEvt time="${time}" objId="${objId}"/>`
	}
}

function slideShowExtLst (slide: PresSlide): string {
	const exts: string[] = []
	const traces = slide.laserTraces
	if (traces && traces.length > 0) {
		const body = traces
			.map(pts => `<p14:tracePtLst>${pts.map(p => `<p14:tracePt t="${Math.round(p.t)}" x="${Math.round(p.x)}" y="${Math.round(p.y)}"/>`).join('')}</p14:tracePtLst>`)
			.join('')
		exts.push(`<p:ext uri="${URI_LASER_TRACE_LST}"><p14:laserTraceLst>${body}</p14:laserTraceLst></p:ext>`)
	}
	const evts = slide.showEvents
	if (evts && evts.length > 0) {
		exts.push(`<p:ext uri="${URI_SHOW_EVT_LST}"><p14:showEvtLst>${evts.map(showEventXml).join('')}</p14:showEvtLst></p:ext>`)
	}
	return exts.length > 0 ? `<p:extLst>${exts.join('')}</p:extLst>` : ''
}

// XML-GEN: First 6 functions create the base /ppt files

/**
 * Generate XML ContentType
 * @param {PresSlide[]} slides - slides
 * @param {SlideLayout[]} slideLayouts - slide layouts
 * @param {PresSlide} masterSlide - master slide
 * @returns XML
 */
export function makeXmlContTypes (slides: PresSlide[], slideLayouts: SlideLayout[], masterSlide?: PresSlide, tracking?: PresentationTrackingParts): string {
	let strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF
	strXml += '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
	strXml += '<Default Extension="xml" ContentType="application/xml"/>'
	strXml += '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
	strXml += '<Default Extension="jpeg" ContentType="image/jpeg"/>'
	strXml += '<Default Extension="jpg" ContentType="image/jpeg"/>'
	strXml += '<Default Extension="svg" ContentType="image/svg+xml"/>'

	// STEP 1: Add standard/any media types used in Presentation
	strXml += '<Default Extension="png" ContentType="image/png"/>'
	strXml += '<Default Extension="gif" ContentType="image/gif"/>'
	strXml += '<Default Extension="m4v" ContentType="video/mp4"/>' // NOTE: Hard-Code this extension as it wont be created in loop below (as extn !== type)
	strXml += '<Default Extension="mp4" ContentType="video/mp4"/>' // NOTE: Hard-Code this extension as it wont be created in loop below (as extn !== type)
	slides.forEach(slide => {
		(slide._relsMedia || []).forEach(rel => {
			if (!rel.isLinked && rel.type !== 'image' && rel.type !== 'online' && rel.type !== 'chart' && rel.extn !== 'm4v' && !strXml.includes(rel.type)) {
				strXml += '<Default Extension="' + rel.extn + '" ContentType="' + rel.type + '"/>'
			}
		})
	})
	strXml += '<Default Extension="vml" ContentType="application/vnd.openxmlformats-officedocument.vmlDrawing"/>'
	strXml += '<Default Extension="xlsx" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"/>'

	// STEP 2: Add presentation and slide master(s)/slide(s)
	strXml += '<Override PartName="/ppt/presentation.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presentation.main+xml"/>'
	strXml += '<Override PartName="/ppt/notesMasters/notesMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesMaster+xml"/>'
	
	// NOTE: Only one slideMaster (slideMaster1.xml) is ever written (see pptxgen.ts), so emit its Override once — NOT once per slide, which referenced phantom slideMaster parts and triggered the PowerPoint repair dialog (#1444)
	strXml += '<Override PartName="/ppt/slideMasters/slideMaster1.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideMaster+xml"/>'

	// ECMA-376 Part 2 §9.1.2.2 [M2.4]: every part needs Default and/or Override (Override wins).
	// Default Extension="xml" is application/xml, which is not the Slide part content type
	// (Part 1 §13.3.8: application/vnd.openxmlformats-officedocument.presentationml.slide+xml).
	// Override is required because that Default is "not consistent with" the part (§9.1.2.2).
	// Omitting it makes PowerPoint treat the package as corrupt (repair dialog / dropped content).
	slides.forEach((slide, idx) => {
		strXml += `<Override PartName="/ppt/slides/slide${idx + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/>`
		// Add charts if any
		slide._relsChart.forEach(rel => {
			strXml += chartContentTypes(rel)
		})
	})

	// STEP 3: Core PPT
	strXml += '<Override PartName="/ppt/presProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.presProps+xml"/>'
	strXml += '<Override PartName="/ppt/viewProps.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.viewProps+xml"/>'
	strXml += '<Override PartName="/ppt/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
	strXml += '<Override PartName="/ppt/theme/theme2.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml"/>'
	strXml += '<Override PartName="/ppt/tableStyles.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.tableStyles+xml"/>'

	// STEP 4: Add Slide Layouts
	slideLayouts.forEach((layout, idx) => {
		strXml += `<Override PartName="/ppt/slideLayouts/slideLayout${idx + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slideLayout+xml"/>`
		; (layout._relsChart || []).forEach(rel => {
			strXml += chartContentTypes(rel)
		})
	})

	// STEP 5: Add notes slide(s)
	slides.forEach((_slide, idx) => {
		strXml += `<Override PartName="/ppt/notesSlides/notesSlide${idx + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.notesSlide+xml"/>`
	})

	// STEP 5b: Content-part / ink / Office App parts (MS-PPTX §2.2.3 / §2.2.13) — only when requested.
	slides.forEach(slide => {
		(slide._rels || []).forEach(rel => {
			if ((rel.type !== SLIDE_OBJECT_TYPES.contentPart && rel.type !== SLIDE_OBJECT_TYPES.officeApp) || !rel.contentType) return
			strXml += `<Override PartName="/${extPartPackagePath(rel.Target)}" ContentType="${rel.contentType}"/>`
		})
	})

	// STEP 5c: Modern threaded comments (MS-PPTX §2.16) — only when a slide has comments.
	if (slides.some(s => (s.comments ?? []).length > 0)) {
		strXml += `<Override PartName="/ppt/authors.xml" ContentType="${AUTHOR_PART_CONTENT_TYPE}"/>`
		slides.forEach((slide, idx) => {
			if ((slide.comments ?? []).length > 0)
				strXml += `<Override PartName="/ppt/comments/commentSlide${idx + 1}.xml" ContentType="${COMMENT_PART_CONTENT_TYPE}"/>`
		})
	}

	// STEP 5c: Revision / Changes Information parts (MS-PPTX §2.1.2 / §2.1.4) — opt-in, zero-or-one.
	if (tracking?.revisionInfo)
		strXml += `<Override PartName="/ppt/revisionInfo.xml" ContentType="${REVISION_INFO_CONTENT_TYPE}"/>`
	if (tracking?.changesInfo)
		strXml += `<Override PartName="/ppt/changesInfo.xml" ContentType="${CHANGES_INFO_CONTENT_TYPE}"/>`

	// STEP 6: Add rels
	; (masterSlide?._relsChart ?? []).forEach(rel => {
		strXml += chartContentTypes(rel)
	})
	; (masterSlide?._relsMedia ?? []).forEach(rel => {
		if (!rel.isLinked && rel.type !== 'image' && rel.type !== 'online' && rel.type !== 'chart' && rel.extn !== 'm4v' && !strXml.includes(rel.type)) { strXml += ' <Default Extension="' + rel.extn + '" ContentType="' + rel.type + '"/>' }
	})

	// LAST: Finish XML (Resume core)
	strXml += ' <Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>'
	strXml += ' <Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>'
	strXml += '</Types>'

	return strXml
}

/**
 * Creates `_rels/.rels`
 * @returns XML
 */
export function makeXmlRootRels (): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
		<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>
		<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>
		<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="ppt/presentation.xml"/>
		</Relationships>`
}

/**
 * Count text paragraphs across the deck for `app.xml`
 * - derived rather than exposed: the value is a fact about the deck, not a caller preference
 * @param {PresSlide[]} slides - presentation slides
 * @returns {number} paragraph count
 */
function countParagraphs (slides: PresSlide[]): number {
	return slides.reduce((total, slide) => total + (slide._slideObjects ?? []).reduce((count, obj) => {
		if (obj._type !== SLIDE_OBJECT_TYPES.text && obj._type !== SLIDE_OBJECT_TYPES.placeholder) return count
		return count + Math.max(1, (obj.text ?? []).length)
	}, 0), 0)
}

/**
 * Creates `docProps/app.xml`
 * @param {PresSlide[]} slides - Presenation Slides
 * @param {string} company - "Company" metadata
 * @param {DocumentProps} [props] - optional extended properties
 * @returns XML
 */
export function makeXmlApp (slides: PresSlide[], company: string, props?: DocumentProps): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">
	${props?.template ? `<Template>${encodeXmlEntities(props.template)}</Template>` : ''}
	${props?.manager ? `<Manager>${encodeXmlEntities(props.manager)}</Manager>` : ''}
	<TotalTime>${typeof props?.totalEditTime === 'number' && isFinite(props.totalEditTime) && props.totalEditTime >= 0 ? Math.round(props.totalEditTime) : 0}</TotalTime>
	<Words>0</Words>
	<Application>Microsoft Office PowerPoint</Application>
	<PresentationFormat>On-screen Show (16:9)</PresentationFormat>
	<Paragraphs>${countParagraphs(slides)}</Paragraphs>
	<Slides>${slides.length}</Slides>
	<Notes>${slides.filter(slide => (slide._slideObjects ?? []).some(obj => obj._type === SLIDE_OBJECT_TYPES.notes)).length}</Notes>
	<HiddenSlides>${slides.filter(slide => slide.hidden).length}</HiddenSlides>
	<MMClips>${slides.reduce((sum, slide) => sum + (slide._slideObjects ?? []).filter(obj => obj._type === SLIDE_OBJECT_TYPES.media).length, 0)}</MMClips>
	<ScaleCrop>false</ScaleCrop>
	<HeadingPairs>
		<vt:vector size="6" baseType="variant">
			<vt:variant><vt:lpstr>Fonts Used</vt:lpstr></vt:variant>
			<vt:variant><vt:i4>2</vt:i4></vt:variant>
			<vt:variant><vt:lpstr>Theme</vt:lpstr></vt:variant>
			<vt:variant><vt:i4>1</vt:i4></vt:variant>
			<vt:variant><vt:lpstr>Slide Titles</vt:lpstr></vt:variant>
			<vt:variant><vt:i4>${slides.length}</vt:i4></vt:variant>
		</vt:vector>
	</HeadingPairs>
	<TitlesOfParts>
		<vt:vector size="${slides.length + 1 + 2}" baseType="lpstr">
			<vt:lpstr>Arial</vt:lpstr>
			<vt:lpstr>Calibri</vt:lpstr>
			<vt:lpstr>Office Theme</vt:lpstr>
			${slides.map((_slideObj, idx) => `<vt:lpstr>Slide ${idx + 1}</vt:lpstr>`).join('')}
		</vt:vector>
	</TitlesOfParts>
	<Company>${encodeXmlEntities(company)}</Company>
	<LinksUpToDate>false</LinksUpToDate>
	<SharedDoc>false</SharedDoc>
	${props?.hyperlinkBase ? `<HyperlinkBase>${encodeXmlEntities(props.hyperlinkBase)}</HyperlinkBase>` : ''}
	<HyperlinksChanged>false</HyperlinksChanged>
	<AppVersion>16.0000</AppVersion>
	</Properties>`
}

/**
 * Creates `docProps/core.xml`
 * @param {string} title - metadata data
 * @param {string} subject - metadata data
 * @param {string} author - metadata value
 * @param {string} revision - metadata value
 * @param {Date} [created] - OPC `dcterms:created` (defaults to now)
 * @param {Date} [modified] - OPC `dcterms:modified` (defaults to now)
 * @param {DocumentProps} [props] - optional extra core properties
 * @returns XML
 */
export function makeXmlCore (title: string, subject: string, author: string, revision: string, created?: Date, modified?: Date, props?: DocumentProps): string {
	const createdAt = (created ?? new Date()).toISOString().replace(/\.\d\d\dZ/, 'Z')
	const modifiedAt = (modified ?? new Date()).toISOString().replace(/\.\d\d\dZ/, 'Z')
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
	<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:dcmitype="http://purl.org/dc/dcmitype/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">
		<dc:title>${encodeXmlEntities(title)}</dc:title>
		<dc:subject>${encodeXmlEntities(subject)}</dc:subject>
		${props?.description ? `<dc:description>${encodeXmlEntities(props.description)}</dc:description>` : ''}
		${props?.keywords ? `<cp:keywords>${encodeXmlEntities(props.keywords)}</cp:keywords>` : ''}
		${props?.category ? `<cp:category>${encodeXmlEntities(props.category)}</cp:category>` : ''}
		${props?.contentStatus ? `<cp:contentStatus>${encodeXmlEntities(props.contentStatus)}</cp:contentStatus>` : ''}
		${props?.version ? `<cp:version>${encodeXmlEntities(props.version)}</cp:version>` : ''}
		${props?.language ? `<dc:language>${encodeXmlEntities(props.language)}</dc:language>` : ''}
		${props?.identifier ? `<dc:identifier>${encodeXmlEntities(props.identifier)}</dc:identifier>` : ''}
		${props?.lastPrinted ? `<cp:lastPrinted>${encodeXmlEntities(props.lastPrinted)}</cp:lastPrinted>` : ''}
		<dc:creator>${encodeXmlEntities(author)}</dc:creator>
		<cp:lastModifiedBy>${encodeXmlEntities(author)}</cp:lastModifiedBy>
		<cp:revision>${revision}</cp:revision>
		<dcterms:created xsi:type="dcterms:W3CDTF">${createdAt}</dcterms:created>
		<dcterms:modified xsi:type="dcterms:W3CDTF">${modifiedAt}</dcterms:modified>
	</cp:coreProperties>`
}

/**
 * Creates `ppt/_rels/presentation.xml.rels`
 * @param {PresSlide[]} slides - Presenation Slides
 * @returns XML
 */
export function makeXmlPresentationRels (slides: PresSlide[], tracking?: PresentationTrackingParts): string {
	let intRelNum = 1
	let strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF
	strXml += '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
	strXml += '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster" Target="slideMasters/slideMaster1.xml"/>'
	for (let idx = 1; idx <= slides.length; idx++) {
		strXml += `<Relationship Id="rId${++intRelNum}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide${idx}.xml"/>`
	}
	intRelNum++
	strXml +=
		`<Relationship Id="rId${intRelNum + 0}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="notesMasters/notesMaster1.xml"/>` +
		`<Relationship Id="rId${intRelNum + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps" Target="presProps.xml"/>` +
		`<Relationship Id="rId${intRelNum + 2}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps" Target="viewProps.xml"/>` +
		`<Relationship Id="rId${intRelNum + 3}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="theme/theme1.xml"/>` +
		`<Relationship Id="rId${intRelNum + 4}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles" Target="tableStyles.xml"/>`
	// MS-PPTX §2.1.6: implicit authors rel from presentation (only when comments exist).
	let extraRel = intRelNum + 5
	if (slides.some(s => (s.comments ?? []).length > 0))
		strXml += `<Relationship Id="rId${extraRel++}" Type="${AUTHOR_REL_TYPE}" Target="authors.xml"/>`
	// MS-PPTX §2.1.2 / §2.1.4: implicit Internal rels from Presentation; parts MUST NOT have outbound rels.
	if (tracking?.revisionInfo)
		strXml += `<Relationship Id="rId${extraRel++}" Type="${REVISION_INFO_REL_TYPE}" Target="revisionInfo.xml" TargetMode="Internal"/>`
	if (tracking?.changesInfo)
		strXml += `<Relationship Id="rId${extraRel}" Type="${CHANGES_INFO_REL_TYPE}" Target="changesInfo.xml" TargetMode="Internal"/>`
	strXml += '</Relationships>'

	return strXml
}

// XML-GEN: Functions that run 1-N times (once for each Slide)

/**
 * Collect animations from all objects on a slide
 * @param {PresSlide} slide - slide object
 * @returns {SlideObjectAnimation[]} array of object animations with their indices
 */
/**
 * Collect media shapes that need a playback entry in the slide timing tree.
 * The shape id targeted by `<p:spTgt spid>` is the media `<p:cNvPr id>`, which the
 * media emitter computes as `mediaRid + 2`. Online (linked) videos are excluded —
 * they don't support the embedded playback timing tree.
 */
function collectMediaPlayback (slide: PresSlide): MediaPlaybackEntry[] {
	const entries: MediaPlaybackEntry[] = []
	;(slide._slideObjects ?? []).forEach(obj => {
		if (obj._type !== SLIDE_OBJECT_TYPES.media || (obj.mtype !== 'audio' && obj.mtype !== 'video')) return
		const o = obj.options
		// Defaults are all false: no timing node unless at least one playback flag is set.
		if (!o || !(o.autoplay || o.loop || o.fullScreen || o.mute)) return
		entries.push({
			spid: (obj.mediaRid ?? 0) + 2,
			kind: obj.mtype === 'video' ? 'video' : 'audio',
			autoplay: o.autoplay,
			loop: o.loop,
			fullScreen: o.fullScreen,
			mute: o.mute,
			isNarration: o.isNarration,
		})
	})
	return entries
}

function collectSlideAnimations (slide: PresSlide): SlideObjectAnimation[] {
	const animations: SlideObjectAnimation[] = []

	if (!slide._slideObjects) return animations

	slide._slideObjects.forEach((slideObj, index) => {
		// Warn when animations are incorrectly attached to individual text runs
		if (slideObj.text && slideObj.text.length > 1) {
			slideObj.text.forEach((textObj: TextProps) => {
				if (textObj.options?.animation) {
					const preview = typeof textObj.text === 'string' ? textObj.text.substring(0, 30) : ''
					console.warn(
						'Warning: Animations on individual text pieces within an array are not supported. ' +
						'Please apply animation to the container options instead.\n' +
						`Text: "${preview}..."`
					)
				}
			})
		}

		// MelleB/feat/appear-on-click: `appearOnClick` → appear entrance on click (ignored if `animation` is set)
		let animConfig = slideObj.options?.animation
		if (!animConfig && slideObj.options?.appearOnClick) {
			animConfig = { type: 'appear', trigger: 'onClick', duration: 1, delay: 0 }
		}
		if (!animConfig) return

		let animation: AnimationConfig
		if (typeof animConfig === 'string') {
			animation = {
				type: animConfig as AnimationType,
				trigger: 'onClick',
				duration: 1000,
				delay: 0,
			}
		} else {
			animation = {
				trigger: 'onClick',
				duration: 1000,
				delay: 0,
				...animConfig,
			}
		}

		// Media cNvPr id is mediaRid+2 (not the auto shape index). Charts keep index+2 / sId.
		const shapeId = slideObj._type === SLIDE_OBJECT_TYPES.media && slideObj.mediaRid != null
			? slideObj.mediaRid + 2
			: slideObj.options?.sId ?? index + 2

		animations.push({
			objectIndex: index,
			shapeId,
			animation,
			buildKind: slideObj._type === SLIDE_OBJECT_TYPES.chart ? 'chart' : 'shape',
		})
	})

	return animations
}

/**
 * Generates XML for the slide file (`ppt/slides/slide1.xml`)
 * @param {PresSlide} slide - the slide object to transform into XML
 * @return {string} XML
 */
export function makeXmlSlide (slide: PresSlide, sections?: SectionProps[]): string {
	// Pre-assign stable section GUIDs and resolve any section/summary zoom anchors on this slide
	// (MS-PPTX §2.9/§2.11) before slideObjectToXml runs, so zoom XML can reference them.
	if (sections) sections.forEach(s => { if (!s._id) s._id = getUuid('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx') })
	resolveZoomSections(slide, sections)

	const animations = collectSlideAnimations(slide)
	const mediaPlayback = collectMediaPlayback(slide)
	const timingXml = animations.length > 0 || mediaPlayback.length > 0 ? createTimingXml(animations, mediaPlayback) : ''
	const transitionXml = genXmlTransition(slide)
	const showExtXml = slideShowExtLst(slide)
	const hasTrans = transitionXml.length > 0
	const hasP14 = hasTrans || showExtXml.length > 0
	const hasComments = (slide.comments ?? []).length > 0
	const commentRelExt = hasComments
		? `<p:ext uri="${COMMENT_REL_URI}"><p188:commentRel xmlns:p188="${P188_NS}" r:id="rId${slideCommentsRelId(slide)}"/></p:ext>`
		: ''
	const sldExtLst = [showExtXml.replace(/^<p:extLst>|<\/p:extLst>$/g, ''), commentRelExt].filter(Boolean)
	const sldExtXml = sldExtLst.length > 0 ? `<p:extLst>${sldExtLst.join('')}</p:extLst>` : ''

	return (
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}` +
		'<p:sld xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
		'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" ' +
		`xmlns:m="${MATH_NS}" ` +
		`xmlns:a14="${A14_NS}" ` +
		(hasP14 ? `xmlns:p14="${P14_NS}" ` : '') +
		`xmlns:mc="${MC_NS}" mc:Ignorable="a14${hasP14 ? ' p14' : ''}"` +
		`${slide?.hidden ? ' show="0"' : ''}>` +
		`${slideObjectToXml(slide)}` +
		'<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>' +
		`${transitionXml}` +
		`${timingXml}` +
		`${sldExtXml}` +
		'</p:sld>'
	)
}

/**
 * Get text content of Notes from Slide
 * @param {PresSlide} slide - the slide object to transform into XML
 * @return {string} notes text
 */
export function getNotesFromSlide (slide: PresSlide): string {
	let notesText = ''

	slide._slideObjects.forEach(data => {
		if (data._type === SLIDE_OBJECT_TYPES.notes) notesText += data?.text && data.text[0] ? data.text[0].text : ''
	})

	return notesText.replace(/\r*\n/g, CRLF)
}

/**
 * Generate XML for Notes Master (notesMaster1.xml)
 * @returns {string} XML
 */
export function makeXmlNotesMaster (): string {
	// ECMA-376 §4.4.1.24 CT_NotesMaster = cSld + EG_TopLevelSlide(clrMap) + hf? + notesStyle? + extLst?
	// No placeholder shapes are required. PowerPoint's repair strips <p:ph> placeholder <p:sp> shapes from a
	// notesMaster (issue #1443), so emit a spec-compliant empty spTree: bg + clrMap + notesStyle only.
	const notesStyle = '<p:notesStyle><a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr></p:notesStyle>'
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<p:notesMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:bg><p:bgRef idx="1001"><a:schemeClr val="bg1"/></p:bgRef></p:bg><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr></p:spTree></p:cSld><p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>${notesStyle}</p:notesMaster>`
}

/**
 * Creates Notes Slide (`ppt/notesSlides/notesSlide1.xml`)
 * @param {PresSlide} slide - the slide object to transform into XML
 * @return {string} XML
 */
export function makeXmlNotesSlide (slide: PresSlide): string {
	const notesText = getNotesFromSlide(slide)
	// ECMA-376 §5.1.5.2.6 `CT_RegularTextRun` / `a:t` (`standards/ecma/part-22_drawingml-reference-material-drawingml-main.txt`):
	// an empty run is invalid. PowerPoint repair flags `<a:t></a:t>` on notesSlide (Juliussssssss 3260b6e).
	// Empty notes keep a valid paragraph with only `a:endParaRPr` (`CT_TextParagraph`).
	const notesRun = notesText
		? `<a:r><a:rPr lang="en-US" dirty="0"/><a:t>${encodeXmlEntities(notesText)}</a:t></a:r>`
		: ''
	return (
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<p:notes xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:cSld><p:spTree><p:nvGrpSpPr><p:cNvPr id="1" name=""/><p:cNvGrpSpPr/><p:nvPr/></p:nvGrpSpPr><p:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="0" cy="0"/><a:chOff x="0" y="0"/><a:chExt cx="0" cy="0"/></a:xfrm></p:grpSpPr><p:sp><p:nvSpPr><p:cNvPr id="2" name="Slide Image Placeholder 1"/><p:cNvSpPr><a:spLocks noGrp="1" noRot="1" noChangeAspect="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldImg"/></p:nvPr></p:nvSpPr><p:spPr/></p:sp><p:sp><p:nvSpPr><p:cNvPr id="3" name="Notes Placeholder 2"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="body" idx="1"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p>${notesRun}<a:endParaRPr lang="en-US" dirty="0"/></a:p></p:txBody></p:sp><p:sp><p:nvSpPr><p:cNvPr id="4" name="Slide Number Placeholder 3"/><p:cNvSpPr><a:spLocks noGrp="1"/></p:cNvSpPr><p:nvPr><p:ph type="sldNum" sz="quarter" idx="10"/></p:nvPr></p:nvSpPr><p:spPr/><p:txBody><a:bodyPr/><a:lstStyle/><a:p><a:fld id="${SLDNUMFLDID}" type="slidenum"><a:rPr lang="en-US"/><a:t>${slide._slideNum}</a:t></a:fld><a:endParaRPr lang="en-US"/></a:p></p:txBody></p:sp></p:spTree><p:extLst><p:ext uri="{BB962C8B-B14F-4D97-AF65-F5344CB8AC3E}"><p14:creationId xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main" val="1024086991"/></p:ext></p:extLst></p:cSld><p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr></p:notes>`
	)
}

/**
 * Generates the XML layout resource from a layout object
 * @param {SlideLayout} layout - slide layout (master)
 * @return {string} XML
 */
export function makeXmlLayout (layout: SlideLayout): string {
	// `preserve` has always been written as "1", so it stays on unless the caller turns it off
	let attrs = layout.preserve === false ? '' : ' preserve="1"'
	if (layout.layoutType) attrs += ` type="${layout.layoutType}"`
	if (layout.matchingName) attrs += ` matchingName="${encodeXmlEntities(layout.matchingName)}"`
	// both default to true in the schema, so only the "off" case is written
	if (layout.showMasterShapes === false) attrs += ' showMasterSp="0"'
	if (layout.showMasterPlaceholderAnimation === false) attrs += ' showMasterPhAnim="0"'
	if (layout.userDrawn === true) attrs += ' userDrawn="1"'

	// CT_ColorMapping requires all twelve attributes, so a partial override is filled from the
	// identity map - which is exactly what inheriting `a:masterClrMapping` means
	const clrMapOvr = layout.colorMapOverride
		? `<p:clrMapOvr><a:overrideClrMapping${Object.entries(DEF_COLOR_MAP).map(([slot, fallback]) => ` ${slot}="${layout.colorMapOverride?.[slot as keyof ColorMapOverrideProps] ?? fallback}"`).join('')}/></p:clrMapOvr>`
		: '<p:clrMapOvr><a:masterClrMapping/></p:clrMapOvr>'

	const transitionXml = genXmlTransition(layout)
	const hasP14 = transitionXml.includes('p14:') || transitionXml.includes('xmlns:p14')
	const ns = hasP14 ? ` xmlns:p14="${P14_NS}"` : ''

	// CT_SlideLayout sequence: cSld, clrMapOvr, transition, timing, hf, extLst
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		<p:sldLayout xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"${ns}${attrs}>
		${slideObjectToXml(layout)}
		${clrMapOvr}${transitionXml}</p:sldLayout>`
}

/**
 * Creates Slide Master 1 (`ppt/slideMasters/slideMaster1.xml`)
 * @param {PresSlide} slide - slide object that represents master slide layout
 * @param {SlideLayout[]} layouts - slide layouts
 * @return {string} XML
 */
export function makeXmlMaster (slide: PresSlide, layouts: SlideLayout[]): string {
	// NOTE: Pass layouts as static rels because they are not referenced any time
	const layoutDefs = layouts.map((_layoutDef, idx) => `<p:sldLayoutId id="${LAYOUT_IDX_SERIES_BASE + idx}" r:id="rId${slide._rels.length + idx + 1}"/>`)

	let strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF
	strXml +=
		'<p:sldMaster xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">'
	strXml += slideObjectToXml(slide)
	strXml +=
		'<p:clrMap bg1="lt1" tx1="dk1" bg2="lt2" tx2="dk2" accent1="accent1" accent2="accent2" accent3="accent3" accent4="accent4" accent5="accent5" accent6="accent6" hlink="hlink" folHlink="folHlink"/>'
	strXml += '<p:sldLayoutIdLst>' + layoutDefs.join('') + '</p:sldLayoutIdLst>'
	// ECMA-376 §4.4.1.22 `p:hf@sldNum` (default true) enables the slide-number placeholder.
	// Honour it only when the user set slideNumber — otherwise keep it off so unused masters
	// do not surface a default sldNum slot.
	strXml += `<p:hf sldNum="${slide._slideNumberProps ? '1' : '0'}" hdr="0" ftr="0" dt="0"/>`
	strXml +=
		'<p:txStyles>' +
		' <p:titleStyle>' +
		'  <a:lvl1pPr algn="ctr" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="0"/></a:spcBef><a:buNone/><a:defRPr sz="4400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mj-lt"/><a:ea typeface="+mj-ea"/><a:cs typeface="+mj-cs"/></a:defRPr></a:lvl1pPr>' +
		' </p:titleStyle>' +
		' <p:bodyStyle>' +
		'  <a:lvl1pPr marL="342900" indent="-342900" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="3200" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>' +
		'  <a:lvl2pPr marL="742950" indent="-285750" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="–"/><a:defRPr sz="2800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>' +
		'  <a:lvl3pPr marL="1143000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2400" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>' +
		'  <a:lvl4pPr marL="1600200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="–"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>' +
		'  <a:lvl5pPr marL="2057400" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="»"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>' +
		'  <a:lvl6pPr marL="2514600" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr>' +
		'  <a:lvl7pPr marL="2971800" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr>' +
		'  <a:lvl8pPr marL="3429000" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr>' +
		'  <a:lvl9pPr marL="3886200" indent="-228600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:spcBef><a:spcPct val="20000"/></a:spcBef><a:buFont typeface="Arial" pitchFamily="34" charset="0"/><a:buChar char="•"/><a:defRPr sz="2000" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr>' +
		' </p:bodyStyle>' +
		' <p:otherStyle>' +
		'  <a:defPPr><a:defRPr lang="en-US"/></a:defPPr>' +
		'  <a:lvl1pPr marL="0" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl1pPr>' +
		'  <a:lvl2pPr marL="457200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl2pPr>' +
		'  <a:lvl3pPr marL="914400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl3pPr>' +
		'  <a:lvl4pPr marL="1371600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl4pPr>' +
		'  <a:lvl5pPr marL="1828800" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl5pPr>' +
		'  <a:lvl6pPr marL="2286000" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl6pPr>' +
		'  <a:lvl7pPr marL="2743200" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl7pPr>' +
		'  <a:lvl8pPr marL="3200400" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl8pPr>' +
		'  <a:lvl9pPr marL="3657600" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1"><a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/></a:defRPr></a:lvl9pPr>' +
		' </p:otherStyle>' +
		'</p:txStyles>'
	strXml += '</p:sldMaster>'

	return strXml
}

// XML-GEN: Last 5 functions create root /ppt files

/**
 * Creates `ppt/theme/theme1.xml` (slide master) or `ppt/theme/theme2.xml` (notes master).
 * Distinct `thm15:themeFamily` ids keep the notes theme from being a byte-identical twin
 * (Office repair creates theme2 when notesMaster shares theme1; Juliussssssss 9bdfe09 / SCV-Soft 1a1998e).
 * @return {string} XML
 */
export function makeXmlTheme (pres: IPresentationProps, variant: 'slide' | 'notes' = 'slide'): string {
	const majorFont = pres.theme?.headFontFace ? `<a:latin typeface="${pres.theme?.headFontFace}"/>` : '<a:latin typeface="Calibri Light" panose="020F0302020204030204"/>'
	const minorFont = pres.theme?.bodyFontFace ? `<a:latin typeface="${pres.theme?.bodyFontFace}"/>` : '<a:latin typeface="Calibri" panose="020F0502020204030204"/>'
	const eaFace = pres.theme?.eaFontFace ?? ''
	const csFace = pres.theme?.csFontFace ?? ''
	const themeFamilyId = variant === 'notes' ? '{2E142A2C-CD16-42D6-873A-C26D2A0506FA}' : '{62F939B6-93AF-4DB8-9C6B-D6C7DFDC589F}'
	const themeFamilyVid = variant === 'notes' ? '{1BDDFF52-6CD6-40A5-AB3C-68EB2F1E4D0A}' : '{4A3C46E8-61CC-4603-A589-7422A47A8E4A}'
	// Exactly 12 hex colors required (upstream only checked truthy `.length`, which allowed short arrays → undefined attrs)
	const isCustomColors = Array.isArray(pres.theme?.themeColors) && pres.theme!.themeColors!.length === 12
	const c = resolveThemeColors(pres.theme)
	const themeName = isCustomColors ? 'Custom Theme' : 'Office Theme'
	const schemeName = isCustomColors ? 'Custom' : 'Office'
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?><a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="${themeName}"><a:themeElements><a:clrScheme name="${schemeName}"><a:dk1><a:sysClr val="windowText" lastClr="${c[0]}"/></a:dk1><a:lt1><a:sysClr val="window" lastClr="${c[1]}"/></a:lt1><a:dk2><a:srgbClr val="${c[2]}"/></a:dk2><a:lt2><a:srgbClr val="${c[3]}"/></a:lt2><a:accent1><a:srgbClr val="${c[4]}"/></a:accent1><a:accent2><a:srgbClr val="${c[5]}"/></a:accent2><a:accent3><a:srgbClr val="${c[6]}"/></a:accent3><a:accent4><a:srgbClr val="${c[7]}"/></a:accent4><a:accent5><a:srgbClr val="${c[8]}"/></a:accent5><a:accent6><a:srgbClr val="${c[9]}"/></a:accent6><a:hlink><a:srgbClr val="${c[10]}"/></a:hlink><a:folHlink><a:srgbClr val="${c[11]}"/></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont>${majorFont}<a:ea typeface="${eaFace}"/><a:cs typeface="${csFace}"/><a:font script="Jpan" typeface="游ゴシック Light"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线 Light"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Times New Roman"/><a:font script="Hebr" typeface="Times New Roman"/><a:font script="Thai" typeface="Angsana New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="MoolBoran"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Times New Roman"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/><a:font script="Armn" typeface="Arial"/><a:font script="Bugi" typeface="Leelawadee UI"/><a:font script="Bopo" typeface="Microsoft JhengHei"/><a:font script="Java" typeface="Javanese Text"/><a:font script="Lisu" typeface="Segoe UI"/><a:font script="Mymr" typeface="Myanmar Text"/><a:font script="Nkoo" typeface="Ebrima"/><a:font script="Olck" typeface="Nirmala UI"/><a:font script="Osma" typeface="Ebrima"/><a:font script="Phag" typeface="Phagspa"/><a:font script="Syrn" typeface="Estrangelo Edessa"/><a:font script="Syrj" typeface="Estrangelo Edessa"/><a:font script="Syre" typeface="Estrangelo Edessa"/><a:font script="Sora" typeface="Nirmala UI"/><a:font script="Tale" typeface="Microsoft Tai Le"/><a:font script="Talu" typeface="Microsoft New Tai Lue"/><a:font script="Tfng" typeface="Ebrima"/></a:majorFont><a:minorFont>${minorFont}<a:ea typeface="${eaFace}"/><a:cs typeface="${csFace}"/><a:font script="Jpan" typeface="游ゴシック"/><a:font script="Hang" typeface="맑은 고딕"/><a:font script="Hans" typeface="等线"/><a:font script="Hant" typeface="新細明體"/><a:font script="Arab" typeface="Arial"/><a:font script="Hebr" typeface="Arial"/><a:font script="Thai" typeface="Cordia New"/><a:font script="Ethi" typeface="Nyala"/><a:font script="Beng" typeface="Vrinda"/><a:font script="Gujr" typeface="Shruti"/><a:font script="Khmr" typeface="DaunPenh"/><a:font script="Knda" typeface="Tunga"/><a:font script="Guru" typeface="Raavi"/><a:font script="Cans" typeface="Euphemia"/><a:font script="Cher" typeface="Plantagenet Cherokee"/><a:font script="Yiii" typeface="Microsoft Yi Baiti"/><a:font script="Tibt" typeface="Microsoft Himalaya"/><a:font script="Thaa" typeface="MV Boli"/><a:font script="Deva" typeface="Mangal"/><a:font script="Telu" typeface="Gautami"/><a:font script="Taml" typeface="Latha"/><a:font script="Syrc" typeface="Estrangelo Edessa"/><a:font script="Orya" typeface="Kalinga"/><a:font script="Mlym" typeface="Kartika"/><a:font script="Laoo" typeface="DokChampa"/><a:font script="Sinh" typeface="Iskoola Pota"/><a:font script="Mong" typeface="Mongolian Baiti"/><a:font script="Viet" typeface="Arial"/><a:font script="Uigh" typeface="Microsoft Uighur"/><a:font script="Geor" typeface="Sylfaen"/><a:font script="Armn" typeface="Arial"/><a:font script="Bugi" typeface="Leelawadee UI"/><a:font script="Bopo" typeface="Microsoft JhengHei"/><a:font script="Java" typeface="Javanese Text"/><a:font script="Lisu" typeface="Segoe UI"/><a:font script="Mymr" typeface="Myanmar Text"/><a:font script="Nkoo" typeface="Ebrima"/><a:font script="Olck" typeface="Nirmala UI"/><a:font script="Osma" typeface="Ebrima"/><a:font script="Phag" typeface="Phagspa"/><a:font script="Syrn" typeface="Estrangelo Edessa"/><a:font script="Syrj" typeface="Estrangelo Edessa"/><a:font script="Syre" typeface="Estrangelo Edessa"/><a:font script="Sora" typeface="Nirmala UI"/><a:font script="Tale" typeface="Microsoft Tai Le"/><a:font script="Talu" typeface="Microsoft New Tai Lue"/><a:font script="Tfng" typeface="Ebrima"/></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:lumMod val="110000"/><a:satMod val="105000"/><a:tint val="67000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="103000"/><a:tint val="73000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="105000"/><a:satMod val="109000"/><a:tint val="81000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:satMod val="103000"/><a:lumMod val="102000"/><a:tint val="94000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:satMod val="110000"/><a:lumMod val="100000"/><a:shade val="100000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:lumMod val="99000"/><a:satMod val="120000"/><a:shade val="78000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="6350" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="12700" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln><a:ln w="19050" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:prstDash val="solid"/><a:miter lim="800000"/></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst/></a:effectStyle><a:effectStyle><a:effectLst><a:outerShdw blurRad="57150" dist="19050" dir="5400000" algn="ctr" rotWithShape="0"><a:srgbClr val="000000"><a:alpha val="63000"/></a:srgbClr></a:outerShdw></a:effectLst></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr"/></a:solidFill><a:solidFill><a:schemeClr val="phClr"><a:tint val="95000"/><a:satMod val="170000"/></a:schemeClr></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="93000"/><a:satMod val="150000"/><a:shade val="98000"/><a:lumMod val="102000"/></a:schemeClr></a:gs><a:gs pos="50000"><a:schemeClr val="phClr"><a:tint val="98000"/><a:satMod val="130000"/><a:shade val="90000"/><a:lumMod val="103000"/></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:shade val="63000"/><a:satMod val="120000"/></a:schemeClr></a:gs></a:gsLst><a:lin ang="5400000" scaled="0"/></a:gradFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults/><a:extraClrSchemeLst/><a:extLst><a:ext uri="{05A4C25C-085E-4340-85A3-A5531E510DB2}"><thm15:themeFamily xmlns:thm15="http://schemas.microsoft.com/office/thememl/2012/main" name="Office Theme" id="${themeFamilyId}" vid="${themeFamilyVid}"/></a:ext></a:extLst></a:theme>`
}

const P15_NS = 'http://schemas.microsoft.com/office/powerpoint/2012/main'
const SLD_GUIDE_URI = '{EFAFB233-063F-42B5-8137-9DF3F51BA10A}'
const NOTES_GUIDE_URI = '{2D200454-40CA-4A62-9FC3-DE9A4176ACB9}'

/**
 * MS-PPTX §2.2.11 / §2.4.3.4: `p15:sldGuideLst` or `p15:notesGuideLst` under presentation extLst.
 * `pos` is EMU relative to the left (vert) / top (horz) edge; `clr` child is required (§2.4.3.3).
 */
function genXmlGuideLst (guides: GuideProps[], element: 'sldGuideLst' | 'notesGuideLst'): string {
	const uri = element === 'sldGuideLst' ? SLD_GUIDE_URI : NOTES_GUIDE_URI
	let xml = `<p:ext uri="${uri}"><p15:${element} xmlns:p15="${P15_NS}">`
	guides.forEach((g, i) => {
		const pos = Math.round((g.pos ?? 0) * EMU)
		const clr = g.color ?? 'A0A0A0'
		const id = g.id ?? i + 1
		xml +=
			`<p15:guide id="${id}" orient="${g.orient === 'horz' ? 'horz' : 'vert'}" pos="${pos}"` +
			`${g.name ? ` name="${encodeXmlEntities(g.name)}"` : ''}${g.userDrawn === false ? ' userDrawn="0"' : ' userDrawn="1"'}>` +
			`<p15:clr><a:srgbClr val="${clr}"/></p15:clr></p15:guide>`
	})
	xml += `</p15:${element}></p:ext>`
	return xml
}

/**
 * Create presentation file (`ppt/presentation.xml`)
 * @see https://docs.microsoft.com/en-us/office/open-xml/structure-of-a-presentationml-document
 * @see http://www.datypic.com/sc/ooxml/t-p_CT_Presentation.html
 * @param {IPresentationProps} pres - presentation
 * @return {string} XML
 */
export function makeXmlPresentation (pres: IPresentationProps): string {
	let strXml =
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}` +
		'<p:presentation xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
		`xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main" firstSlideNum="${pres.firstSlideNum}" ${pres.rtlMode ? 'rtl="1"' : ''} saveSubsetFonts="1" autoCompressPictures="0">`

	// STEP 1: Add slide master (SPEC: tag 1 under <presentation>)
	strXml += '<p:sldMasterIdLst><p:sldMasterId id="2147483648" r:id="rId1"/></p:sldMasterIdLst>'

	// ECMA-376 §4.3.1.30: sldIdLst is the presentation's slide list; child order is viewing order.
	// §4.3.1.29: @id is ST_SlideId (unique in the presentation); @r:id resolves the Slide part
	// via the Presentation relationship of type .../relationships/slide (§13.3.8).
	// MS-PPTX section lists (p14:sldIdLst) reference those ST_SlideId values, not r:id.
	strXml += '<p:sldIdLst>'
	pres.slides.forEach(slide => {
		const tagLst = genXmlDesignTagLst(slide.designTags, true)
		if (tagLst) {
			strXml += `<p:sldId id="${slide._slideId}" r:id="rId${slide._rId}"><p:extLst><p:ext uri="${URI_DESIGN_TAG_LST}">${tagLst}</p:ext></p:extLst></p:sldId>`
		} else {
			strXml += `<p:sldId id="${slide._slideId}" r:id="rId${slide._rId}"/>`
		}
	})
	strXml += '</p:sldIdLst>'

	// CT_Presentation sequence is sldMasterIdLst, notesMasterIdLst, handoutMasterIdLst, sldIdLst.
	// Office warns if notesMasterIdLst is placed in that schema position, so emit it after sldIdLst.
	// (NOTE: length+2 is from `presentation.xml.rels` func (since we have to match this rId, we just use same logic))
	// Presentations open without warning without this line, but then they aren't previewed in Finder or viewable on iOS.
	strXml += `<p:notesMasterIdLst><p:notesMasterId r:id="rId${pres.slides.length + 2}"/></p:notesMasterIdLst>`

	// STEP 4: Add sizes
	// `@type` records which preset the dimensions match; omitted when the caller does not say
	strXml += `<p:sldSz cx="${pres.presLayout.width}" cy="${pres.presLayout.height}"${pres.slideSizeType ? ` type="${pres.slideSizeType}"` : ''}/>`
	strXml += `<p:notesSz cx="${pres.presLayout.height}" cy="${pres.presLayout.width}"/>`

	// CT_Presentation sequence puts these between `notesSz` and `defaultTextStyle` (photoAlbum, kinsoku)
	if (pres.photoAlbum) {
		const album = pres.photoAlbum
		let attrs = ''
		if (album.blackWhite === true) attrs += ' bw="1"'
		if (album.showCaptions === true) attrs += ' showCaptions="1"'
		if (album.layout) attrs += ` layout="${album.layout}"`
		if (album.frame) attrs += ` frame="${album.frame}"`
		strXml += `<p:photoAlbum${attrs}/>`
	}
	// `invalStChars` and `invalEndChars` are required on CT_Kinsoku, so a partial value is dropped
	// rather than emitted as an element PowerPoint would refuse to open
	if (pres.kinsoku?.invalidStartChars && pres.kinsoku.invalidEndChars) {
		const kin = pres.kinsoku
		const lang = kin.lang ? ` lang="${encodeXmlEntities(kin.lang)}"` : ''
		strXml += `<p:kinsoku${lang} invalStChars="${encodeXmlEntities(kin.invalidStartChars)}" invalEndChars="${encodeXmlEntities(kin.invalidEndChars)}"/>`
	}

	// STEP 5: Add text styles
	strXml += '<p:defaultTextStyle>'
	for (let idy = 1; idy < 10; idy++) {
		strXml +=
			`<a:lvl${idy}pPr marL="${(idy - 1) * 457200}" algn="l" defTabSz="914400" rtl="0" eaLnBrk="1" latinLnBrk="0" hangingPunct="1">` +
			'<a:defRPr sz="1800" kern="1200"><a:solidFill><a:schemeClr val="tx1"/></a:solidFill><a:latin typeface="+mn-lt"/><a:ea typeface="+mn-ea"/><a:cs typeface="+mn-cs"/>' +
			`</a:defRPr></a:lvl${idy}pPr>`
	}
	strXml += '</p:defaultTextStyle>'

	// STEP 6: Add Sections and/or Guides (any of them forces an extLst)
	const hasSections = pres.sections && pres.sections.length > 0
	const hasGuides = !!(pres.guides && pres.guides.length > 0)
	const hasNotesGuides = !!(pres.notesGuides && pres.notesGuides.length > 0)
	if (hasSections || hasGuides || hasNotesGuides) {
		strXml += '<p:extLst>'
		if (hasSections) {
			strXml += '<p:ext uri="{521415D9-36F7-43E2-AB2F-B90AF26B5E84}">'
			strXml += '<p14:sectionLst xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main">'
			pres.sections.forEach(sect => {
				// Stable GUID so section/summary zoom objects can anchor to it (MS-PPTX §2.9/§2.11).
				if (!sect._id) sect._id = getUuid('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx')
				strXml += `<p14:section name="${encodeXmlEntities(sect.title)}" id="{${sect._id}}"><p14:sldIdLst>`
				;(sect._slides ?? []).forEach(slide => (strXml += `<p14:sldId id="${slide._slideId}"/>`))
				strXml += '</p14:sldIdLst></p14:section>'
			})
			strXml += '</p14:sectionLst></p:ext>'
		}
		// MS-PPTX §2.2.11: emit each guide list only when the caller opted in (no empty stubs).
		if (hasGuides) strXml += genXmlGuideLst(pres.guides ?? [], 'sldGuideLst')
		if (hasNotesGuides) strXml += genXmlGuideLst(pres.notesGuides ?? [], 'notesGuideLst')
		strXml += '</p:extLst>'
	}

	// Done
	strXml += '</p:presentation>'
	return strXml
}

/**
 * Create `ppt/presProps.xml`
 * @return {string} XML
 */
export function makeXmlPresProps (pres?: IPresentationProps): string {
	// CT_PresentationProperties sequence: htmlPubPr, webPr, prnPr, showPr, clrMru, extLst.
	// Opt-in: emit nothing extra unless set. prnPr/clrMru live here (not on CT_Presentation).
	const prnPr = makeXmlPrnPr(pres)
	const showPr = makeXmlShowPr(pres)
	const clrMru = makeXmlClrMru(pres)
	const extLst = makeXmlPresPropsExtLst(pres)
	return (
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}` +
		'<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">' +
		`${prnPr}${showPr}${clrMru}${extLst}</p:presentationPr>`
	)
}

function makeXmlPrnPr (pres?: IPresentationProps): string {
	if (!pres?.printProps) return ''
	const prn = pres.printProps
	let attrs = ''
	if (prn.what) attrs += ` prnWhat="${prn.what}"`
	if (prn.colorMode) attrs += ` clrMode="${prn.colorMode}"`
	if (prn.hiddenSlides === true) attrs += ' hiddenSlides="1"'
	if (prn.scaleToFitPaper === true) attrs += ' scaleToFitPaper="1"'
	if (prn.frameSlides === true) attrs += ' frameSlides="1"'
	return `<p:prnPr${attrs}/>`
}

function makeXmlClrMru (pres?: IPresentationProps): string {
	if (!Array.isArray(pres?.recentColors) || pres.recentColors.length === 0) return ''
	return `<p:clrMru>${pres.recentColors.map(color => createColorElement(color)).join('')}</p:clrMru>`
}

/**
 * `p:showPr` for slide-show options, or '' when none are set.
 * ECMA-376 §19.2.1.30 requires a present/browse/kiosk choice whenever `showPr` is written.
 */
function makeXmlShowPr (pres?: IPresentationProps): string {
	const show: SlideShowProps = pres?.slideShow ?? {}
	const browseMode = show.browseMode ?? pres?.browseMode
	const laserColor = show.laserColor ?? pres?.laserColor
	const hasShow = !!pres?.slideShow && Object.keys(show).length > 0
	const hasExt = typeof browseMode === 'boolean' || !!laserColor
	if (!hasShow && !hasExt) return ''

	const mode = show.mode ?? 'present'
	if (show.mode && !['present', 'browse', 'kiosk'].includes(show.mode)) {
		console.warn(`[pptxgenjs] slideShow.mode must be 'present' | 'browse' | 'kiosk' - "${String(show.mode)}" ignored, 'present' used`)
	}
	let attrs = ''
	if (show.loop === true) attrs += ' loop="1"'
	if (show.showNarration === false) attrs += ' showNarration="0"'
	if (show.showAnimation === false) attrs += ' showAnimation="0"'
	if (show.useTimings === false) attrs += ' useTimings="0"'

	const resolvedMode = ['present', 'browse', 'kiosk'].includes(mode) ? mode : 'present'
	const choice = resolvedMode === 'browse'
		? `<p:browse${show.showScrollbar === true ? ' showScrollbar="1"' : ''}/>`
		: resolvedMode === 'kiosk' ? '<p:kiosk/>' : '<p:present/>'

	const showExts: string[] = []
	if (typeof browseMode === 'boolean')
		showExts.push(`<p:ext uri="${URI_BROWSE_MODE}"><p14:browseMode xmlns:p14="${P14_NS}" showStatus="${browseMode ? '1' : '0'}"/></p:ext>`)
	if (laserColor)
		showExts.push(`<p:ext uri="${URI_LASER_CLR}"><p14:laserClr xmlns:p14="${P14_NS}">${createColorElement(laserColor)}</p14:laserClr></p:ext>`)

	return `<p:showPr${attrs}>${choice}${showExts.length > 0 ? `<p:extLst>${showExts.join('')}</p:extLst>` : ''}</p:showPr>`
}

function makeXmlPresPropsExtLst (pres?: IPresentationProps): string {
	const exts: string[] = []
	if (typeof pres?.defaultImageDpi === 'number' && Number.isFinite(pres.defaultImageDpi) && pres.defaultImageDpi >= 0)
		exts.push(`<p:ext uri="${URI_DEFAULT_IMAGE_DPI}"><p14:defaultImageDpi xmlns:p14="${P14_NS}" val="${Math.round(pres.defaultImageDpi)}"/></p:ext>`)
	if (pres?.discardImageEditData)
		exts.push(`<p:ext uri="${URI_DISCARD_IMAGE_EDIT_DATA}"><p14:discardImageEditData xmlns:p14="${P14_NS}" val="1"/></p:ext>`)
	if (pres?.readonlyRecommended)
		exts.push(`<p:ext uri="${URI_READONLY_RECOMMENDED}"><p1710:readonlyRecommended xmlns:p1710="${P1710_NS}" val="1"/></p:ext>`)
	if (pres?.chartTrackingRefBased)
		exts.push('<p:ext uri="{FD5EFAAD-0ECE-453E-9831-46B23BE46B34}"><p15:chartTrackingRefBased xmlns:p15="http://schemas.microsoft.com/office/powerpoint/2012/main" val="1"/></p:ext>')
	return exts.length > 0 ? `<p:extLst>${exts.join('')}</p:extLst>` : ''
}

/**
 * Create `ppt/tableStyles.xml`
 * @see: http://openxmldeveloper.org/discussions/formats/f/13/p/2398/8107.aspx
 * @return {string} XML
 */
/**
 * The GUID `a:tblStyleLst@def` has always carried - PowerPoint's "Medium Style 2 - Accent 1"
 */
const DEF_TABLE_STYLE_ID = '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}'

/**
 * CT_TableStyle fixes this child order, and it is neither alphabetical nor the intuitive one:
 * `lastCol` precedes `firstCol`, and `firstRow` sits between `swCell` and `neCell`.
 */
const TABLE_STYLE_PARTS: Array<[keyof TableStyleProps, string]> = [
	['wholeTable', 'wholeTbl'],
	['band1H', 'band1H'],
	['band2H', 'band2H'],
	['band1V', 'band1V'],
	['band2V', 'band2V'],
	['lastCol', 'lastCol'],
	['firstCol', 'firstCol'],
	['lastRow', 'lastRow'],
	['seCell', 'seCell'],
	['swCell', 'swCell'],
	['firstRow', 'firstRow'],
	['neCell', 'neCell'],
	['nwCell', 'nwCell'],
]

/** CT_TableCellBorderStyle child order */
const TABLE_STYLE_BORDERS: Array<keyof TableStyleBorderProps> = ['left', 'right', 'top', 'bottom', 'insideH', 'insideV']

/**
 * One part of a table style (`a:wholeTbl`, `a:band1H`, ...).
 * - `a:tcTxStyle` precedes `a:tcStyle`; inside the cell style, `a:tcBdr` precedes the fill
 */
function genXmlTableStylePart (tag: string, part: TableStylePartProps): string {
	// `b`/`i` are ST_OnOffStyleType: an unset property means "def", i.e. leave it to the theme
	const bold = typeof part.bold === 'boolean' ? ` b="${part.bold ? 'on' : 'off'}"` : ''
	const italic = typeof part.italic === 'boolean' ? ` i="${part.italic ? 'on' : 'off'}"` : ''
	const txStyle = bold || italic || part.color
		? `<a:tcTxStyle${bold}${italic}>${part.color ? createColorElement(part.color) : ''}</a:tcTxStyle>`
		: ''

	const borders = TABLE_STYLE_BORDERS
		.filter(edge => part.borders?.[edge])
		.map(edge => {
			const line = part.borders?.[edge]
			return line ? `<a:${edge}>${genXmlLine(line)}</a:${edge}>` : ''
		})
		.join('')
	const fill = part.fill ? `<a:fill>${genXmlColorSelection(part.fill)}</a:fill>` : ''
	const cellStyle = borders || fill ? `<a:tcStyle>${borders ? `<a:tcBdr>${borders}</a:tcBdr>` : ''}${fill}</a:tcStyle>` : ''

	return txStyle || cellStyle ? `<a:${tag}>${txStyle}${cellStyle}</a:${tag}>` : ''
}

/**
 * Creates `ppt/tableStyles.xml`
 * - with no custom styles this is the same self-closing stub earlier versions wrote
 * @param {TableStyleProps[]} [styles] - custom table style definitions
 * @return {string} XML
 */
export function makeXmlTableStyles (styles?: TableStyleProps[]): string {
	const valid = (styles ?? []).filter(style => {
		// `@styleId` is an ST_Guid and `@styleName` is required; a malformed id would be rejected on open
		if (!/^\{[0-9A-Fa-f]{8}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{4}-[0-9A-Fa-f]{12}\}$/.test(style?.id ?? '')) {
			console.warn(`[pptxgenjs] table style \`id\` must be a braced GUID, e.g. '{5C22544A-7EE6-4342-B048-85BDC9FD1C3A}' - "${String(style?.id)}" ignored`)
			return false
		}
		if (!style.name) {
			console.warn(`[pptxgenjs] table style ${style.id} has no \`name\` - ignored`)
			return false
		}
		return true
	})

	const head = `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<a:tblStyleLst xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" def="${DEF_TABLE_STYLE_ID}"`
	if (valid.length === 0) return `${head}/>`

	const body = valid.map(style => {
		const parts = TABLE_STYLE_PARTS
			.map(([prop, tag]) => {
				const part = style[prop]
				return part && typeof part === 'object' ? genXmlTableStylePart(tag, part) : ''
			})
			.join('')
		return `<a:tblStyle styleId="${style.id}" styleName="${encodeXmlEntities(style.name)}">${parts}</a:tblStyle>`
	}).join('')

	return `${head}>${body}</a:tblStyleLst>`
}

/**
 * Creates `ppt/viewProps.xml`
 * - the hardcoded literal is kept when the caller sets no `viewProps`, so default packages stay byte-identical
 * @return {string} XML
 */
export function makeXmlViewProps (pres?: IPresentationProps): string {
	if (!pres?.viewProps) {
		return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"><p:normalViewPr horzBarState="maximized"><p:restoredLeft sz="15611"/><p:restoredTop sz="94610"/></p:normalViewPr><p:slideViewPr><p:cSldViewPr snapToGrid="0" snapToObjects="1"><p:cViewPr varScale="1"><p:scale><a:sx n="136" d="100"/><a:sy n="136" d="100"/></p:scale><p:origin x="216" y="312"/></p:cViewPr><p:guideLst/></p:cSldViewPr></p:slideViewPr><p:notesTextViewPr><p:cViewPr><p:scale><a:sx n="1" d="1"/><a:sy n="1" d="1"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:notesTextViewPr><p:gridSpacing cx="76200" cy="76200"/></p:viewPr>`
	}

	const view = pres.viewProps
	const zoom = typeof view.zoom === 'number' && isFinite(view.zoom) && view.zoom > 0 ? Math.round(view.zoom) : 136
	const grid = typeof view.gridSpacing === 'number' && isFinite(view.gridSpacing) && view.gridSpacing > 0 ? inch2Emu(view.gridSpacing) : 76200
	const snapGrid = view.snapToGrid === true ? '1' : '0'
	const snapObj = view.snapToObjects === false ? '0' : '1'
	const showGuides = view.showGuides === true ? ' showGuides="1"' : ''
	const showComments = view.showComments === false ? ' showComments="0"' : ''
	const lastView = view.lastView ? ` lastView="${view.lastView}"` : ''

	// `p:guideLst` is the classic guide list, distinct from the MS-PPTX `p15:sldGuideLst` extension
	const guides = (view.guides ?? [])
		.filter((guide): guide is GuideProps & { pos: number } => typeof guide.pos === 'number' && isFinite(guide.pos) && guide.pos >= 0)
		.map(guide => `<p:guide${guide.orient === 'vert' ? ' orient="vert"' : ''} pos="${Math.round(guide.pos * 96)}"/>`)
		.join('')

	return (
		`<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}` +
		`<p:viewPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"${lastView}${showComments}>` +
		'<p:normalViewPr horzBarState="maximized"><p:restoredLeft sz="15611"/><p:restoredTop sz="94610"/></p:normalViewPr>' +
		`<p:slideViewPr><p:cSldViewPr snapToGrid="${snapGrid}" snapToObjects="${snapObj}"${showGuides}>` +
		`<p:cViewPr varScale="1"><p:scale><a:sx n="${zoom}" d="100"/><a:sy n="${zoom}" d="100"/></p:scale><p:origin x="216" y="312"/></p:cViewPr>` +
		(guides ? `<p:guideLst>${guides}</p:guideLst>` : '<p:guideLst/>') +
		'</p:cSldViewPr></p:slideViewPr>' +
		'<p:notesTextViewPr><p:cViewPr><p:scale><a:sx n="1" d="1"/><a:sy n="1" d="1"/></p:scale><p:origin x="0" y="0"/></p:cViewPr></p:notesTextViewPr>' +
		`<p:gridSpacing cx="${grid}" cy="${grid}"/>` +
		'</p:viewPr>'
	)
}
