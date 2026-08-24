/**
 *  :: pptxgen.ts ::
 *
 *  JavaScript framework that creates PowerPoint (pptx) presentations
 *  https://github.com/gitbrent/PptxGenJS
 *
 *  This framework is released under the MIT Public License (MIT)
 *
 *  PptxGenJS (C) 2015-present Brent Ely -- https://github.com/gitbrent
 *  Portions (C) 2026-present lofcz -- https://github.com/lofcz
 *
 *  Some code derived from the OfficeGen project:
 *  github.com/Ziv-Barber/officegen/ (Copyright 2013 Ziv Barber)
 *
 *  Permission is hereby granted, free of charge, to any person obtaining a copy
 *  of this software and associated documentation files (the "Software"), to deal
 *  in the Software without restriction, including without limitation the rights
 *  to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
 *  copies of the Software, and to permit persons to whom the Software is
 *  furnished to do so, subject to the following conditions:
 *
 *  The above copyright notice and this permission notice shall be included in all
 *  copies or substantial portions of the Software.
 *
 *  THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
 *  IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
 *  FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
 *  AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
 *  LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
 *  OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
 *  SOFTWARE.
 */

/**
 * Units of Measure used in PowerPoint documents
 *
 * PowerPoint units are in `DXA` (except for font sizing)
 * - 1 inch is 1440 DXA
 * - 1 inch is 72 points
 * -  1 DXA is 1/20th's of a point
 * - 20 DXA is 1 point
 *
 * Another form of measurement using is an `EMU`
 * - 914400 EMUs is 1 inch
 * -  12700 EMUs is 1 point
 *
 * @see https://startbigthinksmall.wordpress.com/2010/01/04/points-inches-and-emus-measuring-units-in-office-open-xml/
 */

/**
 * Object Layouts
 *
 * - 16x9 (10" x 5.625")
 * - 16x10 (10" x 6.25")
 * - 4x3 (10" x 7.5")
 * - Wide (13.33" x 7.5")
 * - [custom] (any size)
 *
 * @see https://docs.microsoft.com/en-us/office/open-xml/structure-of-a-presentationml-document
 * @see https://docs.microsoft.com/en-us/previous-versions/office/developer/office-2010/hh273476(v=office.14)
 */

import { JSZip } from '@node-projects/jszip'
import Slide from './slide'
import {
	AlignH,
	AlignV,
	AnimationPreset,
	CHART_TYPE,
	ChartType,
	DEF_PRES_LAYOUT,
	DEF_PRES_LAYOUT_NAME,
	DEF_SLIDE_MARGIN_IN,
	EMU,
	OutputType,
	SCHEME_COLOR_NAMES,
	ANCHOR,
	SHAPE_TYPE,
	SLIDE_OBJECT_TYPES,
	SchemeColor,
	ShapeType,
	TransitionType,
	WRITE_OUTPUT_TYPE,
} from './core-enums'
import {
	AddFontOptions,
	AddSlideProps,
	CompressionLevel,
	ChangesInfoProps,
	Color,
	CommentAuthorProps,
	DefineLayoutProps,
	DocumentProps,
	KinsokuProps,
	PhotoAlbumProps,
	PrintProps,
	SlideSizeType,
	ViewProps,
	GuideProps,
	IPresentationProps,
	MediaOnError,
	PresLayout,
	PresSlide,
	RevisionInfoProps,
	SectionProps,
	SlideLayout,
	SlideMasterProps,
	SlideShowProps,
	SlideNumberProps,
	TableStyleProps,
	TableToSlidesProps,
	ThemeProps,
	WriteBaseProps,
	WriteFileProps,
	WriteProps,
} from './core-interfaces'
import * as genCharts from './charts'
import { embedFontsIntoZip, PendingEmbedFont } from './gen-fonts'
import * as genObj from './gen-objects'
import * as genMedia from './gen-media'
import * as genTable from './gen-tables'
import * as genXml from './xml'
import * as genComments from './gen-comments'
import * as genRevision from './gen-revision'
import { importNodeBuiltin, isNodeRuntime, warnDeprecatedOnce } from './gen-utils'
import { VERSION } from './version.generated'

export default class PptxGenJS implements IPresentationProps {
	// Property getters/setters

	/**
	 * Presentation layout name
	 * Standard layouts:
	 * - 'LAYOUT_4x3'   (10"    x 7.5")
	 * - 'LAYOUT_16x9'  (10"    x 5.625")
	 * - 'LAYOUT_16x10' (10"    x 6.25")
	 * - 'LAYOUT_WIDE'  (13.33" x 7.5")
	 * Custom layouts:
	 * Use `pptx.defineLayout()` to create custom layouts (e.g.: 'A4')
	 * @type {string}
	 * @see https://support.office.com/en-us/article/Change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e
	 */
	private _layout: string
	public set layout(value: string) {
		const newLayout: PresLayout = this.LAYOUTS[value]

		if (newLayout) {
			this._layout = value
			this._presLayout = newLayout
		} else {
			throw new Error('UNKNOWN-LAYOUT')
		}
	}

	public get layout(): string {
		return this._layout
	}

	/**
	 * PptxGenJS Library Version
	 */
	private readonly _version: string = VERSION
	public get version(): string {
		return this._version
	}

	/**
	 * @type {string}
	 */
	private _author: string
	public set author(value: string) {
		this._author = value
	}

	public get author(): string {
		return this._author
	}

	/**
	 * OPC `dcterms:created` written to `docProps/core.xml`.
	 */
	private _created?: Date
	public set created(value: Date) {
		this._created = value
	}

	public get created(): Date | undefined {
		return this._created
	}

	/**
	 * OPC `dcterms:modified` written to `docProps/core.xml`.
	 */
	private _modified?: Date
	public set modified(value: Date) {
		this._modified = value
	}

	public get modified(): Date | undefined {
		return this._modified
	}

	/**
	 * Failed image/media load policy. Default throws so invalid input still fails the export.
	 */
	private _mediaOnError: MediaOnError = 'throw'
	public set mediaOnError(value: MediaOnError) {
		this._mediaOnError = value === 'placeholder' ? 'placeholder' : 'throw'
	}

	public get mediaOnError(): MediaOnError {
		return this._mediaOnError
	}

	/**
	 * @type {string}
	 */
	private _company: string
	public set company(value: string) {
		this._company = value
	}

	public get company(): string {
		return this._company
	}

	/**
	 * @type {string}
	 * @note the `revision` value must be a whole number only (without "." or "," - otherwise, PPT will throw errors upon opening!)
	 */
	private _revision: string
	public set revision(value: string) {
		this._revision = value
	}

	public get revision(): string {
		return this._revision
	}

	/**
	 * @type {string}
	 */
	private _subject: string
	public set subject(value: string) {
		this._subject = value
	}

	public get subject(): string {
		return this._subject
	}

	/**
	 * @type {ThemeProps}
	 */
	private _theme?: ThemeProps
	public set theme(value: ThemeProps) {
		this._theme = value
	}

	public get theme(): ThemeProps | undefined {
		return this._theme
	}

	/**
	 * @type {string}
	 */
	private _title: string
	public set title(value: string) {
		this._title = value
	}

	public get title(): string {
		return this._title
	}

	/**
	 * Whether Right-to-Left (RTL) mode is enabled
	 * @type {boolean}
	 */
	private _rtlMode: boolean
	public set rtlMode(value: boolean) {
		this._rtlMode = value
	}

	public get rtlMode(): boolean {
		return this._rtlMode
	}

	/** tableStyles - custom table style definitions written into `ppt/tableStyles.xml` */
	private _tableStyles?: TableStyleProps[]
	public set tableStyles(value: TableStyleProps[] | undefined) {
		this._tableStyles = value
	}

	public get tableStyles(): TableStyleProps[] | undefined {
		return this._tableStyles
	}

	/**
	 * Starting slide number (`p:presentation@firstSlideNum`)
	 * @default 1
	 * @example pptx.firstSlideNum = 0
	 */
	private _firstSlideNum = 1
	public get firstSlideNum (): number {
		return this._firstSlideNum
	}

	public set firstSlideNum (value: number) {
		const n = Number(value)
		if (!Number.isInteger(n) || n < 0) {
			console.warn(`[pptxgenjs] firstSlideNum must be an integer >= 0 (got ${String(value)}); keeping ${this._firstSlideNum}`)
			return
		}
		this._firstSlideNum = n
	}

	/**
	 * Zip compression for exported files - document config applied by every export method
	 * @default 'none'
	 * @since v4.1.0
	 */
	private _compression: CompressionLevel
	public set compression(value: CompressionLevel) {
		// Guard plain-JS callers: anything but a valid level would otherwise silently select DEFLATE
		if (value !== 'none' && value !== 'fast' && value !== 'best') {
			console.warn(`[pptxgenjs] invalid compression "${String(value)}" - valid values are 'none' | 'fast' | 'best'; using 'none'`)
			value = 'none'
		}
		this._compression = value
	}

	public get compression(): CompressionLevel {
		return this._compression
	}

	/** master slide layout object */
	private readonly _masterSlide: PresSlide
	public get masterSlide(): PresSlide {
		return this._masterSlide
	}

	/** this Presentation's Slide objects */
	private readonly _slides: PresSlide[]
	public get slides(): PresSlide[] {
		return this._slides
	}

	/** this Presentation's sections */
	private readonly _sections: SectionProps[]
	public get sections(): SectionProps[] {
		return this._sections
	}

	/**
	 * Editor alignment guides (MS-PPTX §2.4.3.3). Emitted as `p15:sldGuideLst`.
	 * @example pptx.guides = [{ orient: 'vert', pos: 3.5 }, { orient: 'horz', pos: 2 }]
	 */
	public guides: GuideProps[] = []
	/** Default image DPI. MS-PPTX §2.3.1.5 `p14:defaultImageDpi` on `presentationPr`. */
	public defaultImageDpi?: number
	/** Discard image edit data on save. MS-PPTX §2.3.1.6 `p14:discardImageEditData`. */
	public discardImageEditData?: boolean
	/** Recommend opening read-only. MS-PPTX §2.14.1.1 `p1710:readonlyRecommended`. */
	public readonlyRecommended?: boolean
	/** Browse-mode status bar. MS-PPTX §2.3.1.2 `p14:browseMode` on `showPr`. */
	public browseMode?: boolean
	/** Laser-pointer color. MS-PPTX §2.3.1.16 `p14:laserClr` on `showPr`. */
	public laserColor?: string
	/**
	 * Slide-show options (`p:showPr`). Opt-in — omitted unless set.
	 * @example pptx.slideShow = { mode: 'kiosk', loop: true }
	 */
	public slideShow?: SlideShowProps

	/**
	 * Chart data-point properties and datalabels follow their cell references.
	 * MS-PPTX §2.4.1.1 `chartTrackingRefBased` on `presentationPr` extLst.
	 * @example pptx.chartTrackingRefBased = true
	 */
	public chartTrackingRefBased?: boolean

	/**
	 * Notes-page alignment guides (MS-PPTX §2.4.1.3). Emitted as `p15:notesGuideLst`.
	 * @example pptx.notesGuides = [{ orient: 'horz', pos: 2 }]
	 */
	public notesGuides: GuideProps[] = []
	/** Additional document properties written to `docProps`. */
	public documentProps?: DocumentProps
	/** Slide-size preset (`p:sldSz@type`). */
	public slideSizeType?: SlideSizeType
	/** Photo-album mode (`p:photoAlbum`). */
	public photoAlbum?: PhotoAlbumProps
	/** East Asian line-breaking rules (`p:kinsoku`). */
	public kinsoku?: KinsokuProps
	/** Print defaults (`p:prnPr` on `presentationPr`). */
	public printProps?: PrintProps
	/** Recently-used colours (`p:clrMru` on `presentationPr`). */
	public recentColors?: Color[]
	/** View properties (`ppt/viewProps.xml`). Unset keeps the previous hardcoded viewPr. */
	public viewProps?: ViewProps

	/**
	 * Modern comment authors (MS-PPTX §2.16). Emitted to `ppt/authors.xml`.
	 * Auto-populated from slide `addComment` author names if left empty.
	 * @example pptx.commentAuthors = [{ name: 'Ada Lovelace', initials: 'AL' }]
	 */
	public commentAuthors: CommentAuthorProps[] = []

	/**
	 * Revision Information part (MS-PPTX §2.1.2). Opt-in — not emitted unless set.
	 * @example pptx.revisionInfo = { clients: [{ id: 'app-1', v: 1, dt: '2024-08-15T00:00:00Z' }] }
	 */
	public revisionInfo?: boolean | RevisionInfoProps

	/**
	 * Changes Information part (MS-PPTX §2.1.4). Opt-in — not emitted unless set.
	 * @example pptx.changesInfo = true
	 */
	public changesInfo?: boolean | ChangesInfoProps

	/** slide layout definition objects, used for generating slide layout files */
	private readonly _slideLayouts: SlideLayout[]
	public get slideLayouts(): SlideLayout[] {
		return this._slideLayouts
	}

	private LAYOUTS: { [key: string]: PresLayout }

	/** Fonts registered via addFont() — embedded into the PPTX on export */
	private readonly _embedFonts: PendingEmbedFont[] = []

	// Exposed class props
	private readonly _alignH = AlignH
	public get AlignH(): typeof AlignH {
		return this._alignH
	}

	private readonly _alignV = AlignV
	public get AlignV(): typeof AlignV {
		return this._alignV
	}

	private readonly _chartType = ChartType
	public get ChartType(): typeof ChartType {
		return this._chartType
	}

	private readonly _outputType = OutputType
	public get OutputType(): typeof OutputType {
		return this._outputType
	}

	private _presLayout: PresLayout
	public get presLayout(): PresLayout {
		return this._presLayout
	}

	private readonly _schemeColor = SchemeColor
	public get SchemeColor(): typeof SchemeColor {
		return this._schemeColor
	}

	private readonly _shapeType = ShapeType
	public get ShapeType(): typeof ShapeType {
		return this._shapeType
	}

	private readonly _transitionType = TransitionType
	public get TransitionType(): typeof TransitionType {
		return this._transitionType
	}

	private readonly _animationPreset = AnimationPreset
	public get AnimationPreset(): typeof AnimationPreset {
		return this._animationPreset
	}

	/**
	 * @depricated use `ChartType`
	 */
	private readonly _charts = CHART_TYPE
	public get charts(): typeof CHART_TYPE {
		return this._charts
	}

	/**
	 * @depricated use `SchemeColor`
	 */
	private readonly _colors = SCHEME_COLOR_NAMES
	public get colors(): typeof SCHEME_COLOR_NAMES {
		return this._colors
	}

	/**
	 * @depricated use `ShapeType`
	 */
	private readonly _shapes = SHAPE_TYPE
	public get shapes(): typeof SHAPE_TYPE {
		return this._shapes
	}

	/** Connection-site helpers for connectors (`line.sourceAnchorPos` / `targetAnchorPos`) — ZentoSoft */
	private readonly _anchor = ANCHOR
	public get anchor(): typeof ANCHOR {
		return this._anchor
	}
	/** Alias of `anchor` (PascalCase, matches `ShapeType`) */
	public get Anchor(): typeof ANCHOR {
		return this._anchor
	}

	private _chartCounter = 0
	private readonly allocChartId = (): number => {
		this._chartCounter += 1
		return this._chartCounter
	}

	constructor() {
		const layout4x3: PresLayout = { name: 'screen4x3', width: 9144000, height: 6858000 }
		const layout16x9: PresLayout = { name: 'screen16x9', width: 9144000, height: 5143500 }
		const layout16x10: PresLayout = { name: 'screen16x10', width: 9144000, height: 5715000 }
		const layoutWide: PresLayout = { name: 'custom', width: 12192000, height: 6858000 }
		// Set available layouts
		this.LAYOUTS = {
			LAYOUT_4x3: layout4x3,
			LAYOUT_16x9: layout16x9,
			LAYOUT_16x10: layout16x10,
			LAYOUT_WIDE: layoutWide,
		}

		// Core
		this._author = 'PptxGenJS'
		this._company = 'PptxGenJS'
		this._revision = '1' // Note: Must be a whole number
		this._subject = 'PptxGenJS Presentation'
		this._title = 'PptxGenJS Presentation'
		// PptxGenJS props
		this._layout = DEF_PRES_LAYOUT
		this._presLayout = {
			name: this.LAYOUTS[DEF_PRES_LAYOUT].name,
			_sizeW: this.LAYOUTS[DEF_PRES_LAYOUT].width,
			_sizeH: this.LAYOUTS[DEF_PRES_LAYOUT].height,
			width: this.LAYOUTS[DEF_PRES_LAYOUT].width,
			height: this.LAYOUTS[DEF_PRES_LAYOUT].height,
		}
		this._rtlMode = false
		this._compression = 'none'
		//
		this._slideLayouts = [
			{
				_allocChartId: this.allocChartId,
				_margin: DEF_SLIDE_MARGIN_IN,
				_name: DEF_PRES_LAYOUT_NAME,
				_presLayout: this._presLayout,
				_rels: [],
				_relsChart: [],
				_relsMedia: [],
				_slide: undefined,
				_slideNum: 1000,
				_slideNumberProps: undefined,
				_slideObjects: [],
			},
		]
		this._slides = []
		this._sections = []
		// The master is a container for master-level objects/rels; its add* methods and
		// slide identity are never used, so add* throw and identity fields hold neutral values.
		const notOnMaster = (): never => {
			throw new Error('add* methods are not available on the master slide')
		}
		this._masterSlide = {
			_allocChartId: this.allocChartId,
			addChart: notOnMaster,
			addImage: notOnMaster,
			addMedia: notOnMaster,
			addNotes: notOnMaster,
			addShape: notOnMaster,
			addTable: notOnMaster,
			addText: notOnMaster,
			addGroup: notOnMaster,
			addConnector: notOnMaster,
			addWordArt: notOnMaster,
			addTransition: notOnMaster,
			addAnimation: notOnMaster,
			addComment: notOnMaster,
			addZoom: notOnMaster,
			addSectionZoom: notOnMaster,
			addSummaryZoom: notOnMaster,
			addContentPart: notOnMaster,
			addInk: notOnMaster,
			addOfficeApp: notOnMaster,
			//
			_name: '',
			_presLayout: this._presLayout,
			_rId: 0,
			_rels: [],
			_relsChart: [],
			_relsMedia: [],
			_slideId: 0,
			_slideLayout: { _allocChartId: this.allocChartId, _name: '', _presLayout: this._presLayout, _rels: [], _relsChart: [], _relsMedia: [], _slideNum: null, _slideObjects: [] },
			_slideNum: null,
			_slideObjects: [],
		}
	}

	/**
	 * Provides an API for `addTableDefinition` to create slides as needed for auto-paging
	 * @param {AddSlideProps} options - slide masterName and/or sectionTitle
	 * @return {PresSlide} new Slide
	 */
	private readonly addNewSlide = (options?: AddSlideProps): PresSlide => {
		// Continue using sections if the slide being paged (the last slide) already belongs to a Section.
		// Find the section that actually contains the parent slide - NOT only the last section - so that
		// autopaged slides follow their parent even when a later section exists (issue #1405).
		const parentSlideNum = this.slides.length > 0 ? this.slides[this.slides.length - 1]._slideNum : undefined
		const parentSect = this.sections.filter(sect => (sect._slides ?? []).some(slide => slide._slideNum === parentSlideNum))[0]

		const opts: AddSlideProps = options ?? {}
		opts.sectionTitle = parentSect ? parentSect.title : undefined

		return this.addSlide(opts)
	}

	/**
	 * Provides an API for `addTableDefinition` to get slide reference by number
	 * @param {number} slideNum - slide number
	 * @return {PresSlide} Slide
	 * @since 3.0.0
	 */
	private readonly getSlide = (slideNum: number): PresSlide => this.slides.filter(slide => slide._slideNum === slideNum)[0]

	/**
	 * Enables the `Slide` class to set PptxGenJS [Presentation] master/layout slidenumbers
	 * @param {SlideNumberProps} slideNum - slide number config
	 */
	private readonly setSlideNumber = (slideNum: SlideNumberProps): void => {
		// 1: Add slideNumber to slideMaster1.xml
		this.masterSlide._slideNumberProps = slideNum

		// 2: Add slideNumber to DEF_PRES_LAYOUT_NAME layout
		this.slideLayouts.filter(layout => layout._name === DEF_PRES_LAYOUT_NAME)[0]._slideNumberProps = slideNum
	}

	/**
	 * Create all chart and media rels for this Presentation
	 * @param {PresSlide | SlideLayout} slide - slide with rels
	 * @param {JSZip} zip - JSZip instance
	 * @param {Promise<string>[]} chartPromises - promise array
	 */
	private readonly createChartMediaRels = (slide: PresSlide | SlideLayout, zip: JSZip, chartPromises: Promise<string>[], mediaPaths: Set<string>): void => {
		slide._relsChart.forEach(rel => chartPromises.push(genCharts.createExcelWorksheet(rel, zip)))
		slide._relsMedia.forEach(rel => {
			if (!rel.isLinked && rel.type !== 'online' && rel.type !== 'hyperlink') {
				// A: Loop vars
				let data: string = rel.data && typeof rel.data === 'string' ? rel.data : ''

				// B: Users will undoubtedly pass various string formats, so correct prefixes as needed
				if (!data.includes(',') && !data.includes(';')) data = 'image/png;base64,' + data
				else if (!data.includes(',')) data = 'image/png;base64,' + data
				else if (!data.includes(';')) data = 'image/png;' + data

				// C: Add media (skip if another rel already wrote this OPC part)
				const mediaPath = rel.Target.replace(/\.\./g, 'ppt')
				if (mediaPaths.has(mediaPath)) return
				mediaPaths.add(mediaPath)
				zip.file(mediaPath, data.split(',').pop() ?? '', { base64: true })
			}
		})
	}

	/**
	 * Create and export the .pptx file
	 * @param {string} exportName - output file type
	 * @param {Blob} blobContent - Blob content
	 * @return {Promise<string>} Promise with file name
	 */
	private readonly writeFileToBrowser = async (exportName: string, blobContent: Blob): Promise<string> => {
		// STEP 1: Create element
		const eleLink = document.createElement('a')
		eleLink.setAttribute('style', 'display:none;')
		eleLink.dataset.interception = 'off' // @see https://docs.microsoft.com/en-us/sharepoint/dev/spfx/hyperlinking
		document.body.appendChild(eleLink)

		// STEP 2: Download file to browser
		// DESIGN: Use `createObjectURL()` to D/L files in client browsers (FYI: synchronously executed)
		if (window.URL.createObjectURL) {
			const url = window.URL.createObjectURL(new Blob([blobContent], { type: 'application/vnd.openxmlformats-officedocument.presentationml.presentation' }))
			eleLink.href = url
			eleLink.download = exportName
			eleLink.click()

			// Clean-up (NOTE: Add a slight delay before removing to avoid 'blob:null' error in Firefox Issue#81)
			setTimeout(() => {
				window.URL.revokeObjectURL(url)
				document.body.removeChild(eleLink)
			}, 100)

			// Done
			return await Promise.resolve(exportName)
		}

		return await Promise.resolve(exportName)
	}

	/**
	 * Create and export the .pptx file
	 * @param {WRITE_OUTPUT_TYPE} outputType - output file type
	 * @return {Promise<string | ArrayBuffer | Blob | Uint8Array>} Promise with data or stream (node) or filename (browser)
	 */
	private readonly exportPresentation = async (props: WriteProps): Promise<string | ArrayBuffer | Blob | Uint8Array> => {
		const arrChartPromises: Promise<string>[] = []
		let arrMediaPromises: Promise<string>[] = []
		const zip = new JSZip()

		// STEP 1: Read/Encode all Media before zip as base64 content, etc. is required
		this.slides.forEach(slide => {
			arrMediaPromises = arrMediaPromises.concat(genMedia.encodeSlideMediaRels(slide, { onError: this._mediaOnError }))
		})
		this.slideLayouts.forEach(layout => {
			arrMediaPromises = arrMediaPromises.concat(genMedia.encodeSlideMediaRels(layout, { onError: this._mediaOnError }))
		})
		arrMediaPromises = arrMediaPromises.concat(genMedia.encodeSlideMediaRels(this.masterSlide, { onError: this._mediaOnError }))

		// STEP 2: Wait for Promises (if any) then generate the PPTX file
		return await Promise.all(arrMediaPromises).then(async () => {
			// A: Size images added without `w`/`h` to their natural dimensions (media bytes are loaded by now)
			this.slides.forEach(slide => genMedia.applyNaturalImageSizes(slide))
			this.slideLayouts.forEach(layout => genMedia.applyNaturalImageSizes(layout))
			genMedia.applyNaturalImageSizes(this.masterSlide)

			// A: Add empty placeholder objects to slides that don't already have them
			this.slides.forEach(slide => {
				if (slide._slideLayout) genObj.addPlaceholdersToSlideLayouts(slide)
			})

			// B: Add all required folders and files
			zip.folder('_rels')
			zip.folder('docProps')
			zip.folder('ppt')?.folder('_rels')
			zip.folder('ppt/charts')?.folder('_rels')
			zip.folder('ppt/embeddings')
			zip.folder('ppt/media')
			zip.folder('ppt/slideLayouts')?.folder('_rels')
			zip.folder('ppt/slideMasters')?.folder('_rels')
			zip.folder('ppt/slides')?.folder('_rels')
			zip.folder('ppt/theme')
			zip.folder('ppt/notesMasters')?.folder('_rels')
			zip.folder('ppt/notesSlides')?.folder('_rels')
			const trackingParts = {
				revisionInfo: genRevision.wantsRevisionInfo(this.revisionInfo),
				changesInfo: genRevision.wantsChangesInfo(this.changesInfo),
			}
			zip.file('[Content_Types].xml', genXml.makeXmlContTypes(this.slides, this.slideLayouts, this.masterSlide, trackingParts)) // TODO: pass only `this` like below! 20200206
			zip.file('_rels/.rels', genXml.makeXmlRootRels())
			zip.file('docProps/app.xml', genXml.makeXmlApp(this.slides, this.company, this.documentProps)) // TODO: pass only `this` like below! 20200206
			zip.file('docProps/core.xml', genXml.makeXmlCore(this.title, this.subject, this.author, this.revision, this.created, this.modified, this.documentProps)) // TODO: pass only `this` like below! 20200206
			zip.file('ppt/_rels/presentation.xml.rels', genXml.makeXmlPresentationRels(this.slides, trackingParts))
			zip.file('ppt/theme/theme1.xml', genXml.makeXmlTheme(this))
			// notesMaster gets its own theme part (Office repair creates theme2 when notesMaster shares theme1; Juliussssssss 9bdfe09).
			zip.file('ppt/theme/theme2.xml', genXml.makeXmlTheme(this, 'notes'))
			zip.file('ppt/presentation.xml', genXml.makeXmlPresentation(this))
			zip.file('ppt/presProps.xml', genXml.makeXmlPresProps(this))
			zip.file('ppt/tableStyles.xml', genXml.makeXmlTableStyles(this._tableStyles))
			zip.file('ppt/viewProps.xml', genXml.makeXmlViewProps(this))

			// C: Create a Layout/Master/Rel/Slide file for each SlideLayout and Slide
			this.slideLayouts.forEach((layout, idx) => {
				zip.file(`ppt/slideLayouts/slideLayout${idx + 1}.xml`, genXml.makeXmlLayout(layout))
				zip.file(`ppt/slideLayouts/_rels/slideLayout${idx + 1}.xml.rels`, genXml.makeXmlSlideLayoutRel(idx + 1, this.slideLayouts))
			})
			// Modern threaded comments (MS-PPTX §2.16): resolve authors once, emit authors + per-slide comment parts.
			const hasComments = this.slides.some(s => (s.comments ?? []).length > 0)
			const commentAuthors = hasComments ? genComments.collectCommentAuthors(this.slides, this.commentAuthors) : []
			if (hasComments) {
				zip.folder('ppt/comments')
				zip.file('ppt/authors.xml', genComments.makeXmlCommentAuthors(commentAuthors))
			}
			if (trackingParts.revisionInfo) {
				zip.file(genRevision.REVISION_INFO_PART, genRevision.makeXmlRevisionInfo(this.revisionInfo as true | RevisionInfoProps))
			}
			if (trackingParts.changesInfo) {
				zip.file(genRevision.CHANGES_INFO_PART, genRevision.makeXmlChangesInfo())
			}
			this.slides.forEach((slide, idx) => {
				zip.file(`ppt/slides/slide${idx + 1}.xml`, genXml.makeXmlSlide(slide, this._sections))
				zip.file(`ppt/slides/_rels/slide${idx + 1}.xml.rels`, genXml.makeXmlSlideRel(this.slides, this.slideLayouts, idx + 1))
				// Create all slide notes related items. Notes of empty strings are created for slides which do not have notes specified, to keep track of _rels.
				zip.file(`ppt/notesSlides/notesSlide${idx + 1}.xml`, genXml.makeXmlNotesSlide(slide))
				zip.file(`ppt/notesSlides/_rels/notesSlide${idx + 1}.xml.rels`, genXml.makeXmlNotesSlideRel(idx + 1))
				if ((slide.comments ?? []).length > 0)
					zip.file(`ppt/comments/commentSlide${idx + 1}.xml`, genComments.makeXmlSlideComments(slide, commentAuthors))
				for (const rel of slide._rels) {
					if (rel.type !== SLIDE_OBJECT_TYPES.contentPart && rel.type !== SLIDE_OBJECT_TYPES.officeApp) continue
					zip.file(rel.Target.replace(/^\.\.\//, 'ppt/'), typeof rel.data === 'string' ? rel.data : '')
				}
			})
			zip.file('ppt/slideMasters/slideMaster1.xml', genXml.makeXmlMaster(this.masterSlide, this.slideLayouts))
			zip.file('ppt/slideMasters/_rels/slideMaster1.xml.rels', genXml.makeXmlMasterRel(this.masterSlide, this.slideLayouts))
			zip.file('ppt/notesMasters/notesMaster1.xml', genXml.makeXmlNotesMaster())
			zip.file('ppt/notesMasters/_rels/notesMaster1.xml.rels', genXml.makeXmlNotesMasterRel())

			// D: Create all Rels (images, media, chart data)
			const mediaPaths = new Set<string>()
			this.slideLayouts.forEach(layout => {
				this.createChartMediaRels(layout, zip, arrChartPromises, mediaPaths)
			})
			this.slides.forEach(slide => {
				this.createChartMediaRels(slide, zip, arrChartPromises, mediaPaths)
			})
			this.createChartMediaRels(this.masterSlide, zip, arrChartPromises, mediaPaths)

			// E: Wait for Promises (if any) then generate the PPTX file
			return await Promise.all(arrChartPromises).then(async () => {
				// E2: Embed custom fonts into the zip before packing (`addFont` → ppt/fonts/*.fntdata)
				await embedFontsIntoZip(zip, this._embedFonts)

				// Effective level: deprecated per-call boolean (when given) overrides the presentation-level setting.
				// `legacy` = boolean `true`, kept on JSZip's default DEFLATE level so existing callers see no perf change.
				let level: CompressionLevel | 'legacy' = this._compression
				if (typeof props.compression === 'boolean') {
					if (props.compression) warnDeprecatedOnce('write-compression', '`compression: true` on write()/writeFile()/stream() is deprecated - set it once on the presentation instead: `pptx.compression = "fast" | "best"`')
					level = props.compression ? 'legacy' : 'none'
				}
				const compression = level === 'none' ? 'STORE' : 'DEFLATE'
				const compressionOptions = level === 'fast' ? { level: 1 } : level === 'best' ? { level: 9 } : undefined

				if (props.outputType === 'STREAM') {
					// A: stream file
					return await zip.generateAsync({ type: 'nodebuffer', compression, compressionOptions })
				} else if (props.outputType) {
					// B: Node [fs]: Output type user option or default
					return await zip.generateAsync({ type: props.outputType, compression, compressionOptions })
				} else {
					// C: Browser: Output blob as app/ms-pptx
					return await zip.generateAsync({ type: 'blob', compression, compressionOptions })
				}
			})
		})
	}

	// EXPORT METHODS

	/**
	 * Export the current Presentation to stream
	 * @param {WriteBaseProps} props - output properties
	 * @returns {Promise<string | ArrayBuffer | Blob | Uint8Array>} file stream
	 */
	async stream(props?: WriteBaseProps): Promise<string | ArrayBuffer | Blob | Uint8Array> {
		return await this.exportPresentation({
			compression: props?.compression,
			outputType: 'STREAM',
		})
	}

	/**
	 * Export the current Presentation as JSZip content with the selected type
	 * @param {WriteProps} props output properties
	 * @returns {Promise<string | ArrayBuffer | Blob | Uint8Array>} file content in selected type
	 */
	async write(props?: WriteProps | WRITE_OUTPUT_TYPE): Promise<string | ArrayBuffer | Blob | Uint8Array> {
		// DEPRECATED: @deprecated v3.5.0 - outputType - [[remove in v4.0.0]]
		if (typeof props === 'string') warnDeprecatedOnce('write-string', 'write(outputType) as a string is deprecated - pass { outputType } instead')
		const propsOutpType: WRITE_OUTPUT_TYPE | undefined = typeof props === 'object' ? props?.outputType : props
		// Leave undefined when not passed so the presentation-level `compression` setting applies
		const propsCompress = typeof props === 'object' ? props?.compression : undefined

		return await this.exportPresentation({
			compression: propsCompress,
			outputType: propsOutpType,
		})
	}

	/**
	 * Export the current Presentation.
	 * Write the generated presentation to disk (Node) or trigger a download (browser).
	 * @param {WriteFileProps} props - output file properties
	 * @returns {Promise<string>} the presentation name
	 */
	async writeFile(props?: WriteFileProps | string): Promise<string> {
		// STEP 1: Figure out where we are running
		const isNode = isNodeRuntime()

		// STEP 2: Normalise the user arguments
		if (typeof props === 'string') {
			// DEPRECATED: @deprecated v3.5.0 - fileName - [[remove in v4.0.0]]
			warnDeprecatedOnce('writeFile-string', 'writeFile(string) is deprecated - pass { fileName } instead')
		}
		const writeProps: WriteFileProps = typeof props === 'string' ? { fileName: props } : (props ?? {})
		const { fileName: rawName = 'Presentation.pptx', compression } = writeProps
		const fileName = rawName.toLowerCase().endsWith('.pptx') ? rawName : `${rawName}.pptx`

		// STEP 3: Get the binary/Blob from exportPresentation()
		const outputType = isNode ? ('nodebuffer' as const) : undefined
		const data = await this.exportPresentation({ compression, outputType })

		// STEP 4: Write the file out
		if (isNode) {
			// Dynamically import to avoid bundling fs in the browser build
			const { promises: fs } = await importNodeBuiltin<typeof import('node:fs')>('fs')
			const { writeFile } = fs
			// Uint8Array covers Node Buffer without referencing the `Buffer` identifier
			// (Vite injects a broken buffer polyfill when it sees `Buffer.*` in the graph).
			if (data instanceof Uint8Array) await writeFile(fileName, data)
			return fileName
		}

		// Browser branch - push a download
		if (data instanceof Blob) await this.writeFileToBrowser(fileName, data)
		return fileName
	}

	// PRESENTATION METHODS

	/**
	 * Opt-in: register a custom font to embed in the exported PPTX.
	 * Nothing is embedded unless this is called — the default write path
	 * writes no Font parts, no fntdata content type, and no font relationships.
	 * - Call before `write` / `writeFile` / `stream`
	 * - Use the same `fontFace` name on text/shape options
	 * - Caller must have a third-party license that permits embedding `fontFile`
	 * @param {AddFontOptions} options - font face, file bytes, and format
	 * @example await pptx.addFont({ fontFace: 'MyFont', fontFile: ttfBuffer, fontType: 'ttf' })
	 */
	async addFont (options: AddFontOptions): Promise<void> {
		if (!options?.fontFace) throw new Error('addFont requires fontFace')
		if (!options.fontFile) throw new Error('addFont requires fontFile (ArrayBuffer)')
		if (!options.fontType || !['ttf', 'otf', 'woff', 'eot'].includes(options.fontType)) {
			throw new Error('addFont requires fontType: ttf | otf | woff | eot')
		}
		this._embedFonts.push({
			fontFace: options.fontFace,
			fontFile: options.fontFile,
			fontType: options.fontType,
		})
	}

	/**
	 * Add a new Section to Presentation
	 * @param {ISectionProps} section - section properties
	 * @example pptx.addSection({ title:'Charts' });
	 */
	addSection(section: SectionProps): void {
		if (!section) console.warn('addSection requires an argument')
		else if (!section.title) console.warn('addSection requires a title')

		const newSection: SectionProps = {
			_type: 'user',
			_slides: [],
			title: section.title,
			_id: section.id ? section.id.replace(/^\{|\}$/g, '') : section._id,
		}

		if (section.order) this.sections.splice(section.order, 0, newSection)
		else this._sections.push(newSection)
	}

	/**
	 * Add a new Slide to Presentation
	 * @param {AddSlideProps} options - slide options
	 * @returns {PresSlide} the new Slide
	 */
	addSlide(options?: AddSlideProps): PresSlide {
		// TODO: DEPRECATED: arg0 string "masterSlideName" dep as of 3.2.0
		const masterSlideName = typeof options === 'string' ? options : options?.masterName ? options.masterName : ''
		let slideLayout: SlideLayout = {
			_allocChartId: this.allocChartId,
			_name: this.LAYOUTS[DEF_PRES_LAYOUT].name,
			_presLayout: this.presLayout,
			_rels: [],
			_relsChart: [],
			_relsMedia: [],
			_slideNum: this.slides.length + 1,
			_slideObjects: [],
		}

		if (masterSlideName) {
			const tmpLayout = this.slideLayouts.filter(layout => layout._name === masterSlideName)[0]
			if (tmpLayout) slideLayout = tmpLayout
		}

		const newSlide = new Slide({
			addSlide: this.addNewSlide,
			allocChartId: this.allocChartId,
			getSlide: this.getSlide,
			presLayout: this.presLayout,
			setSlideNum: this.setSlideNumber,
			slideId: this.slides.length + 256,
			slideRId: this.slides.length + 2,
			slideNumber: this.slides.length + 1,
			slideLayout,
		})

		// A: Add slide to pres
		this._slides.push(newSlide)

		// B: Sections
		// B-1: Add slide to section (if any provided)
		// B-2: Handle slides without a section when sections are already is use ("loose" slides arent allowed, they all need a section)
		if (options?.sectionTitle) {
			const sect = this.sections.filter(section => section.title === options.sectionTitle)[0]
			if (!sect) console.warn(`addSlide: unable to find section with title: "${options.sectionTitle}"`)
			else (sect._slides ??= []).push(newSlide)
		} else if (this.sections && this.sections.length > 0 && (!options?.sectionTitle)) {
			const lastSect = this._sections[this.sections.length - 1]

			// CASE 1: The latest section is a default type - just add this one
			if (lastSect._type === 'default') (lastSect._slides ??= []).push(newSlide)
			// CASE 2: There latest section is NOT a default type - create the defualt, add this slide
			else {
				this._sections.push({
					title: `Default-${this.sections.filter(sect => sect._type === 'default').length + 1}`,
					_type: 'default',
					_slides: [newSlide],
				})
			}
		}

		if (options?.transition) newSlide.transition = options.transition

		return newSlide
	}

	/**
	 * Create a custom Slide Layout in any size
	 * @param {DefineLayoutProps} layout - layout properties (`width`/`height` or `w`/`h`)
	 * @example pptx.defineLayout({ name:'A3', width:16.5, height:11.7 });
	 */
	defineLayout(layout: DefineLayoutProps): void {
		// @see https://support.office.com/en-us/article/Change-the-size-of-your-slides-040a811c-be43-40b9-8d04-0de5ed79987e
		// NOTE: `w`/`h` are accepted as aliases so layouts read like every other sized object (issue #29)
		const width = layout?.width ?? layout?.w
		const height = layout?.height ?? layout?.h

		if (!layout) console.warn('defineLayout requires `{name, width, height}`')
		else if (!layout.name) console.warn('defineLayout requires `name`')
		else if (!width) console.warn('defineLayout requires `width` (or `w`)')
		else if (!height) console.warn('defineLayout requires `height` (or `h`)')
		else if (typeof height !== 'number') console.warn('defineLayout `height` should be a number (inches)')
		else if (typeof width !== 'number') console.warn('defineLayout `width` should be a number (inches)')

		this.LAYOUTS[layout.name] = {
			name: layout.name,
			_sizeW: Math.round(Number(width) * EMU),
			_sizeH: Math.round(Number(height) * EMU),
			width: Math.round(Number(width) * EMU),
			height: Math.round(Number(height) * EMU),
		}
	}

	/**
	 * Create a new slide master [layout] for the Presentation
	 * @param {SlideMasterProps} props - layout properties
	 */
	defineSlideMaster(props: SlideMasterProps): void {
		// (ISSUE#406;PULL#1176) deep clone the props object to avoid mutating the original object
		const propsClone = JSON.parse(JSON.stringify(props))
		if (!propsClone.title) throw new Error('defineSlideMaster() object argument requires a `title` value. (https://gitbrent.github.io/PptxGenJS/docs/masters.html)')

		const newLayout: SlideLayout = {
			_allocChartId: this.allocChartId,
			_margin: propsClone.margin || DEF_SLIDE_MARGIN_IN,
			_name: propsClone.title,
			_presLayout: this.presLayout,
			_rels: [],
			_relsChart: [],
			_relsMedia: [],
			_slide: undefined,
			_slideNum: 1000 + this.slideLayouts.length + 1,
			_slideNumberProps: propsClone.slideNumber || null,
			_slideObjects: [],
			background: propsClone.background || null,
			bkgd: propsClone.bkgd || null,
			layoutType: propsClone.layoutType,
			matchingName: propsClone.matchingName,
			preserve: propsClone.preserve,
			showMasterShapes: propsClone.showMasterShapes,
			showMasterPlaceholderAnimation: propsClone.showMasterPlaceholderAnimation,
			userDrawn: propsClone.userDrawn,
			colorMapOverride: propsClone.colorMapOverride,
			transition: propsClone.transition,
		}

		// STEP 1: Create the Slide Master/Layout
		genObj.createSlideMaster(propsClone, newLayout)

		// STEP 2: Add it to layout defs
		this.slideLayouts.push(newLayout)

		// STEP 3: Add background (image data/path must be captured before `exportPresentation()` is called)
		if (propsClone.background || propsClone.bkgd) genObj.addBackgroundDefinition(propsClone.background, newLayout)

		// STEP 4: Add slideNumber to master slide (if any)
		if (newLayout._slideNumberProps && !this.masterSlide._slideNumberProps) this.masterSlide._slideNumberProps = newLayout._slideNumberProps
	}

	// HTML-TO-SLIDES METHODS

	/**
	 * Reproduces an HTML table as a PowerPoint table - including column widths, style, etc. - creates 1 or more slides as needed
	 * @param {string} eleId - table HTML element ID
	 * @param {TableToSlidesProps} options - generation options
	 */
	tableToSlides(eleId: string, options: TableToSlidesProps = {}): void {
		// @note set the `PPTXGENJS_DEBUG` env var for verbose output of the layout process
		genTable.genTableToSlides(
			this,
			eleId,
			options,
			options?.masterSlideName ? this.slideLayouts.filter(layout => layout._name === options.masterSlideName)[0] : undefined
		)
	}
}
