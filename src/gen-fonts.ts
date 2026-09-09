/**
 * PptxGenJS: Opt-in embedded TrueType fonts (OOXML `ppt/fonts/*.fntdata`)
 *
 * `embedFontsIntoZip` is a no-op when `fonts` is empty — default export never
 * writes Font parts, fntdata content types, or Presentation font relationships.
 * Callers supply their own font bytes via `addFont` and must have a license
 * that permits embedding those files. This library ships no font files.
 *
 * Font-embed orchestration absorbed from MIT `pptx-embed-fonts`
 * (https://github.com/liyao1520/pptx-embed-fonts). TTF/OTF/WOFF→EOT conversion is a
 * vendored subset of MIT `fonteditor-core` under `src/vendor/fonteditor-core`
 * (https://github.com/kekee000/fonteditor-core). WOFF inflate uses pako 3.
 */

import type { JSZip } from '@node-projects/jszip'
import { addZipFile } from './zip'
import type { AddFontOptions, EmbedFontType } from './core-interfaces'
import { convertToEot } from './vendor/fonteditor-core/convert-to-eot.js'

/** Pending font registration prior to zip export */
export interface PendingEmbedFont {
	fontFace: string
	fontFile: ArrayBuffer
	fontType: EmbedFontType
}

interface PreparedFont {
	name: string
	data: ArrayBuffer
	rid: number
}

/** High rId base so font relationships do not collide with slide/master rIds */
const FONT_RID_START = 201314

const FONT_REL_TYPE = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/font'

/**
 * Convert TTF / OTF / WOFF bytes to EOT (Office `fntdata` payload).
 * Lazy-loads pako only for WOFF.
 */
async function fontToEot (type: 'ttf' | 'woff' | 'otf', fontBuffer: ArrayBuffer): Promise<ArrayBuffer> {
	const options: {
		hinting: boolean
		inflate?: (deflatedData: number[]) => number[]
	} = { hinting: true }

	if (type === 'woff') {
		// pako 3: named exports only; InflateInput is Uint8Array | ArrayBuffer (not number[])
		const { inflate } = await import('pako')
		options.inflate = (deflatedData: number[]) =>
			Array.from(inflate(Uint8Array.from(deflatedData)))
	}

	return convertToEot(type, fontBuffer, options)
}

async function toFntdata (fontType: EmbedFontType, fontFile: ArrayBuffer): Promise<ArrayBuffer> {
	switch (fontType) {
		case 'eot':
			// EOT bytes are stored as-is under ppt/fonts/*.fntdata
			return fontFile
		case 'ttf':
		case 'otf':
		case 'woff':
			return await fontToEot(fontType, fontFile)
		default:
			throw new Error(`Invalid font type "${String(fontType as AddFontOptions['fontType'])}" - use ttf | otf | woff | eot`)
	}
}

function ensureFntdataContentType (contentTypesXml: string): string {
	if (/Extension="fntdata"/i.test(contentTypesXml)) return contentTypesXml
	const entry = '<Default Extension="fntdata" ContentType="application/x-fontdata"/>'
	// Prefer inserting among other Defaults near the top of <Types>
	if (/<Default\b/i.test(contentTypesXml)) {
		return contentTypesXml.replace(/<Default\b/i, `${entry}<Default`)
	}
	return contentTypesXml.replace(/<Types([^>]*)>/i, `<Types$1>${entry}`)
}

function ensurePresentationFontAttrs (presentationXml: string): string {
	let xml = presentationXml
	if (/embedTrueTypeFonts=/i.test(xml)) {
		xml = xml.replace(/embedTrueTypeFonts="[^"]*"/i, 'embedTrueTypeFonts="true"')
	} else {
		xml = xml.replace(/<p:presentation\b([^>]*)>/i, '<p:presentation$1 embedTrueTypeFonts="true">')
	}
	if (/saveSubsetFonts=/i.test(xml)) {
		xml = xml.replace(/saveSubsetFonts="[^"]*"/i, 'saveSubsetFonts="true"')
	} else {
		xml = xml.replace(/<p:presentation\b([^>]*)>/i, '<p:presentation$1 saveSubsetFonts="true">')
	}
	return xml
}

function buildEmbeddedFontLst (fonts: PreparedFont[]): string {
	let lst = '<p:embeddedFontLst>'
	for (const font of fonts) {
		lst +=
			'<p:embeddedFont>' +
			`<p:font typeface="${escapeXmlAttr(font.name)}"/>` +
			`<p:regular r:id="rId${font.rid}"/>` +
			'</p:embeddedFont>'
	}
	lst += '</p:embeddedFontLst>'
	return lst
}

function escapeXmlAttr (value: string): string {
	return value
		.replace(/&/g, '&amp;')
		.replace(/</g, '&lt;')
		.replace(/>/g, '&gt;')
		.replace(/"/g, '&quot;')
}

function ensureEmbeddedFontLst (presentationXml: string, fonts: PreparedFont[]): string {
	const lst = buildEmbeddedFontLst(fonts)
	if (/<p:embeddedFontLst[\s>]/i.test(presentationXml)) {
		// Replace existing list wholesale (export always owns the registered set)
		return presentationXml.replace(/<p:embeddedFontLst[\s\S]*?<\/p:embeddedFontLst>/i, lst)
	}
	// ISO 29500: embeddedFontLst precedes defaultTextStyle.
	// Insert after the full opening tag (match consumes `>` or attrs so it's preserved).
	const dtsOpen = presentationXml.match(/<p:defaultTextStyle(?:\s[^>]*)?>/i)
	if (dtsOpen) {
		return presentationXml.replace(dtsOpen[0], `${lst}${dtsOpen[0]}`)
	}
	// Fallback: before closing presentation
	return presentationXml.replace(/<\/p:presentation>/i, `${lst}</p:presentation>`)
}

function appendFontRelationships (relsXml: string, fonts: PreparedFont[]): string {
	const nodes = fonts
		.map(
			font =>
				`<Relationship Id="rId${font.rid}" ` +
				`Type="${FONT_REL_TYPE}" ` +
				`Target="fonts/${font.rid}.fntdata"/>`,
		)
		.join('')
	if (!nodes) return relsXml
	return relsXml.replace(/<\/Relationships>/i, `${nodes}</Relationships>`)
}

/**
 * Convert registered fonts into OOXML `ppt/fonts/*.fntdata` + presentation.xml entries.
 * Mutates `zip` in place (before generateAsync).
 */
export async function embedFontsIntoZip (zip: JSZip, fonts: PendingEmbedFont[]): Promise<void> {
	if (fonts.length === 0) return

	const prepared: PreparedFont[] = []
	let nextRid = FONT_RID_START
	for (const font of fonts) {
		prepared.push({
			name: font.fontFace,
			data: await toFntdata(font.fontType, font.fontFile),
			rid: nextRid++,
		})
	}

	const contentTypesFile = zip.file('[Content_Types].xml')
	if (!contentTypesFile) throw new Error('[Content_Types].xml not found')
	addZipFile(zip, '[Content_Types].xml', ensureFntdataContentType(await contentTypesFile.async('string')))

	const presentationFile = zip.file('ppt/presentation.xml')
	if (!presentationFile) throw new Error('ppt/presentation.xml not found')
	let presentationXml = await presentationFile.async('string')
	presentationXml = ensurePresentationFontAttrs(presentationXml)
	presentationXml = ensureEmbeddedFontLst(presentationXml, prepared)
	addZipFile(zip, 'ppt/presentation.xml', presentationXml)

	const relsFile = zip.file('ppt/_rels/presentation.xml.rels')
	if (!relsFile) throw new Error('ppt/_rels/presentation.xml.rels not found')
	addZipFile(zip, 'ppt/_rels/presentation.xml.rels', appendFontRelationships(await relsFile.async('string'), prepared))

	for (const font of prepared) {
		addZipFile(zip, `ppt/fonts/${font.rid}.fntdata`, font.data, {
			binary: true,
			compression: 'DEFLATE',
		})
	}
}
