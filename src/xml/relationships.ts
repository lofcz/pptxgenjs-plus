/**
 * Relationship-part rendering.
 */

import { CRLF, OOXML_CHARTEX, SLIDE_OBJECT_TYPES, isChartexType } from '../core-enums'
import { COMMENT_REL_TYPE } from '../gen-comments'
import { ISlideRel, ISlideRelChart, ISlideRelMedia, PresSlide, SlideLayout } from '../core-interfaces'
import { REL_TYPE_CUSTOM_XML, REL_TYPE_WEBEXTENSION } from './content-parts'

/**
 * Transforms slide relations to XML string.
 * Extra relations that are not dynamic can be passed using the 2nd arg (e.g. theme relation in master file).
 * These relations use rId series that starts with 1-increased maximum of rIds used for dynamic relations.
 * @param {PresSlide | SlideLayout} slide - slide object whose relations are being transformed
 * @param {{ target: string; type: string }[]} defaultRels - array of default relations
 * @return {string} XML
 */
function slideObjectRelationsToXml (slide: PresSlide | SlideLayout, defaultRels: Array<{ target: string, type: string }>): string {
	let lastRid = 0 // stores maximum rId used for dynamic relations
	let strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' + CRLF + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'

	// STEP 1: Add all rels for this Slide
	slide._rels.forEach((rel: ISlideRel) => {
		lastRid = Math.max(lastRid, rel.rId)
		if (rel.type.toLowerCase().includes('hyperlink')) {
			if (rel.data === 'slide') {
				strXml += `<Relationship Id="rId${rel.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slide${rel.Target}.xml"/>`
			} else {
				strXml += `<Relationship Id="rId${rel.rId}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="${rel.Target}" TargetMode="External"/>`
			}
		} else if (rel.type.toLowerCase().includes('notesSlide')) {
			strXml += `<Relationship Id="rId${rel.rId}" Target="${rel.Target}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"/>`
		} else if (rel.type === SLIDE_OBJECT_TYPES.contentPart) {
			strXml += `<Relationship Id="rId${rel.rId}" Type="${REL_TYPE_CUSTOM_XML}" Target="${rel.Target}"/>`
		} else if (rel.type === SLIDE_OBJECT_TYPES.officeApp) {
			strXml += `<Relationship Id="rId${rel.rId}" Type="${REL_TYPE_WEBEXTENSION}" Target="${rel.Target}"/>`
		}
	})
	; (slide._relsChart || []).forEach((rel: ISlideRelChart) => {
		lastRid = Math.max(lastRid, rel.rId)
		// OPC Part 2: relationship Target is resolved from the source part. Stored `rel.Target`
		// is the absolute `/ppt/charts/chartN.xml` form reused as Content_Types PartName.
		// PowerPoint writes the relative `../charts/chartN.xml`; absolute targets are valid but
		// non-idiomatic and break stricter consumers (PR 1465 / eliasaronson).
		const relType = isChartexType(rel.opts._type) ? OOXML_CHARTEX.relType : 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart'
		strXml += `<Relationship Id="rId${rel.rId}" Type="${relType}" Target="${rel.Target.replace(/^\/ppt\//, '../')}"/>`
	})
	; (slide._relsMedia || []).forEach((rel: ISlideRelMedia) => {
		const relRid = rel.rId.toString()
		lastRid = Math.max(lastRid, rel.rId)
		if (rel.type.toLowerCase().includes('image')) {
			strXml += '<Relationship Id="rId' + relRid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="' + rel.Target + '"/>'
		} else if (rel.type.toLowerCase().includes('audio')) {
			// As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
			const externalAttr = rel.isLinked ? ' TargetMode="External"' : ''
			if (strXml.includes(' Target="' + rel.Target + '"')) {
				strXml += '<Relationship Id="rId' + relRid + '" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="' + rel.Target + '"' + externalAttr + '/>'
			} else {
				strXml += '<Relationship Id="rId' + relRid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/audio" Target="' + rel.Target + '"' + externalAttr + '/>'
			}
		} else if (rel.type.toLowerCase().includes('video')) {
			// As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
			const externalAttr = rel.isLinked ? ' TargetMode="External"' : ''
			if (strXml.includes(' Target="' + rel.Target + '"')) {
				strXml += '<Relationship Id="rId' + relRid + '" Type="http://schemas.microsoft.com/office/2007/relationships/media" Target="' + rel.Target + '"' + externalAttr + '/>'
			} else {
				strXml += '<Relationship Id="rId' + relRid + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video" Target="' + rel.Target + '"' + externalAttr + '/>'
			}
		} else if (rel.type.toLowerCase().includes('online')) {
			// As media has *TWO* rel entries per item, check for first one, if found add second rel with alt style
			if (strXml.includes(' Target="' + rel.Target + '"')) {
				strXml += '<Relationship Id="rId' + relRid + '" Type="http://schemas.microsoft.com/office/2007/relationships/image" Target="' + rel.Target + '"/>'
			} else {
				strXml += '<Relationship Id="rId' + relRid + '" Target="' + rel.Target + '" TargetMode="External" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>'
			}
		}
	})

	// STEP 2: Add default rels
	defaultRels.forEach((rel, idx) => {
		strXml += `<Relationship Id="rId${lastRid + idx + 1}" Type="${rel.type}" Target="${rel.target}"/>`
	})

	strXml += '</Relationships>'
	return strXml
}

/** Highest rId already claimed by dynamic slide rels (hyperlinks, charts, media). */
export function lastSlideDynamicRelId (slide: PresSlide): number {
	let lastRid = 0
	slide._rels.forEach(rel => { lastRid = Math.max(lastRid, rel.rId) })
	;(slide._relsChart || []).forEach(rel => { lastRid = Math.max(lastRid, rel.rId) })
	;(slide._relsMedia || []).forEach(rel => { lastRid = Math.max(lastRid, rel.rId) })
	return lastRid
}

/**
 * rId of the §2.1.5 comments relationship.
 * Default slide rels are layout (+1), notes (+2), then comments (+3).
 */
export function slideCommentsRelId (slide: PresSlide): number {
	return lastSlideDynamicRelId(slide) + 3
}

/**
 * Generates XML string for a slide layout relation file
 * @param {number} layoutNumber - 1-indexed number of a layout that relations are generated for
 * @param {SlideLayout[]} slideLayouts - Slide Layouts
 * @return {string} XML
 */
export function makeXmlSlideLayoutRel (layoutNumber: number, slideLayouts: SlideLayout[]): string {
	return slideObjectRelationsToXml(slideLayouts[layoutNumber - 1], [
		{
			target: '../slideMasters/slideMaster1.xml',
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster',
		},
	])
}

/**
 * Creates `ppt/_rels/slide*.xml.rels`
 * @param {PresSlide[]} slides
 * @param {SlideLayout[]} slideLayouts - Slide Layout(s)
 * @param {number} `slideNumber` 1-indexed number of a layout that relations are generated for
 * @return {string} XML
 */
export function makeXmlSlideRel (slides: PresSlide[], slideLayouts: SlideLayout[], slideNumber: number): string {
	const slide = slides[slideNumber - 1]
	const rels = [
		{
			target: `../slideLayouts/slideLayout${getLayoutIdxForSlide(slides, slideLayouts, slideNumber)}.xml`,
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
		},
		{
			target: `../notesSlides/notesSlide${slideNumber}.xml`,
			type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide',
		},
	]
	// MS-PPTX §2.1.5: explicit comment rel from a slide that has threaded comments.
	if ((slide.comments ?? []).length > 0)
		rels.push({ target: `../comments/commentSlide${slideNumber}.xml`, type: COMMENT_REL_TYPE })
	return slideObjectRelationsToXml(slide, rels)
}

/**
 * Generates XML string for a slide relation file.
 * @param {number} slideNumber - 1-indexed number of a layout that relations are generated for
 * @return {string} XML
 */
export function makeXmlNotesSlideRel (slideNumber: number): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
		<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
			<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster" Target="../notesMasters/notesMaster1.xml"/>
			<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="../slides/slide${slideNumber}.xml"/>
		</Relationships>`
}

/**
 * Creates `ppt/slideMasters/_rels/slideMaster1.xml.rels`
 * @param {PresSlide} masterSlide - Slide object
 * @param {SlideLayout[]} slideLayouts - Slide Layouts
 * @return {string} XML
 */
export function makeXmlMasterRel (masterSlide: PresSlide, slideLayouts: SlideLayout[]): string {
	const defaultRels = slideLayouts.map((_layoutDef, idx) => ({
		target: `../slideLayouts/slideLayout${idx + 1}.xml`,
		type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout',
	}))
	defaultRels.push({ target: '../theme/theme1.xml', type: 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme' })

	return slideObjectRelationsToXml(masterSlide, defaultRels)
}

/**
 * Creates `ppt/notesMasters/_rels/notesMaster1.xml.rels`
 * @return {string} XML
 */
export function makeXmlNotesMasterRel (): string {
	return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>${CRLF}<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
		<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="../theme/theme2.xml"/>
		</Relationships>`
}

/**
 * For the passed slide number, resolves name of a layout that is used for.
 * @param {PresSlide[]} slides - srray of slides
 * @param {SlideLayout[]} slideLayouts - array of slideLayouts
 * @param {number} slideNumber
 * @return {number} slide number
 */
function getLayoutIdxForSlide (slides: PresSlide[], slideLayouts: SlideLayout[], slideNumber: number): number {
	for (let i = 0; i < slideLayouts.length; i++) {
		if (slideLayouts[i]._name === slides[slideNumber - 1]._slideLayout._name) {
			return i + 1
		}
	}

	// IMPORTANT: Return 1 (for `slideLayout1.xml`) when no def is found
	// So all objects are in Layout1 and every slide that references it uses this layout.
	return 1
}
