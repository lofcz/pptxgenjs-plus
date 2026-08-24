/**
 * PptxGenJS: Slide Class
 */

import { CHART_NAME, CHARTEX_NAME, SHAPE_NAME } from './core-enums'
import {
	AddSlideProps,
	AnimationConfig,
	BackgroundProps,
	ConnectorProps,
	Group,
	GroupProps,
	CommentProps,
	ContentPartProps,
	DesignerTag,
	HexColor,
	InkProps,
	LaserTracePoint,
	OfficeAppProps,
	SlideShowEvent,
	IChartMulti,
	IChartOpts,
	IOptsChartData,
	ISlideObject,
	ISlideRel,
	ISlideRelChart,
	ISlideRelMedia,
	ImageProps,
	MediaProps,
	PresLayout,
	PresSlide,
	SectionZoomProps,
	ShapeProps,
	SlideLayout,
	SlideNumberProps,
	SlideTransitionProps,
	SummaryZoomProps,
	TableProps,
	TableRow,
	TextProps,
	TextPropsOptions,
	ZoomProps,
} from './core-interfaces'
import * as genObj from './gen-objects'

/**
 * Copy caller-supplied options so the library never mutates objects it does not own (issue #20)
 * @note shallow by design - the generators rewrite top-level props only (nested props are read, not written)
 */
function cloneOpts<T extends object> (options?: T): T {
	return { ...(options ?? {}) } as T
}

/** Run options for `addText(string)` — shape effects stay on the shape `effectLst` (issue #84). */
function cloneTextRunOpts (options?: TextPropsOptions): TextPropsOptions | undefined {
	if (!options) return undefined
	const runOpts = { ...options }
	delete runOpts.glow
	delete runOpts.softEdge
	delete runOpts.reflection
	delete runOpts.shadow
	delete runOpts.blur
	delete runOpts.fillOverlay
	delete runOpts.effectDag
	return runOpts
}

export default class Slide {
	private readonly _setSlideNum: (value: SlideNumberProps) => void

	public addSlide: (options?: AddSlideProps) => PresSlide
	public getSlide: (slideNum: number) => PresSlide
	public _allocChartId: () => number
	public _name: string
	public _presLayout: PresLayout
	public _rels: ISlideRel[]
	public _relsChart: ISlideRelChart[]
	public _relsMedia: ISlideRelMedia[]
	public _rId: number
	public _slideId: number
	public _slideLayout: SlideLayout
	public _slideNum: number
	public _slideNumberProps?: SlideNumberProps
	public _slideObjects: ISlideObject[]
	public _newAutoPagedSlides?: PresSlide[]
	public transition?: SlideTransitionProps
	public comments?: CommentProps[]
	public laserTraces?: LaserTracePoint[][]
	public showEvents?: SlideShowEvent[]
	public creationId?: number | true
	/** Designer Service tags on this slide's `p:sldId` (MS-PPTX §2.2.20). Opt-in. */
	public designTags?: DesignerTag[]

	constructor(params: {
		addSlide: (options?: AddSlideProps) => PresSlide
		allocChartId: () => number
		getSlide: (slideNum: number) => PresSlide
		presLayout: PresLayout
		setSlideNum: (value: SlideNumberProps) => void
		slideId: number
		slideRId: number
		slideNumber: number
		slideLayout: SlideLayout
	}) {
		this.addSlide = params.addSlide
		this.getSlide = params.getSlide
		this._allocChartId = params.allocChartId
		this._name = `Slide ${params.slideNumber}`
		this._presLayout = params.presLayout
		this._rId = params.slideRId
		this._rels = []
		this._relsChart = []
		this._relsMedia = []
		this._setSlideNum = params.setSlideNum
		this._slideId = params.slideId
		this._slideLayout = params.slideLayout
		this._slideNum = params.slideNumber
		this._slideObjects = []
		/** NOTE: Slide Numbers: In order for Slide Numbers to function they need to be in all 3 files: master/layout/slide
		 * `defineSlideMaster` and `addNewSlide.slideNumber` will add {slideNumber} to `this.masterSlide` and `this.slideLayouts`
		 * so, lastly, add to the Slide now.
		 */
		this._slideNumberProps = this._slideLayout?._slideNumberProps ? this._slideLayout._slideNumberProps : undefined
	}

	/**
	 * Background color
	 * @type {string|BackgroundProps}
	 * @deprecated in v3.3.0 - use `background` instead
	 */
	private _bkgd?: string | BackgroundProps
	public set bkgd(value: string | BackgroundProps) {
		this._bkgd = value
		if (!this._background || !this._background.color) {
			if (!this._background) this._background = {}
			if (typeof value === 'string') this._background.color = value
		}
	}

	public get bkgd(): string | BackgroundProps | undefined {
		return this._bkgd
	}

	/**
	 * Background color or image
	 * @type {BackgroundProps}
	 * @example solid color `background: { color:'FF0000' }`
	 * @example color+trans `background: { color:'FF0000', transparency:0.5 }`
	 * @example base64 `background: { data:'image/png;base64,ABC[...]123' }`
	 * @example url `background: { path:'https://some.url/image.jpg'}`
	 * @since v3.3.0
	 */
	private _background?: BackgroundProps
	public set background(props: BackgroundProps) {
		this._background = props
		// Add background (image data/path must be captured before `exportPresentation()` is called)
		if (props) genObj.addBackgroundDefinition(props, this)
	}

	public get background(): BackgroundProps | undefined {
		return this._background
	}

	/**
	 * Default font color
	 * @type {HexColor}
	 */
	private _color?: HexColor
	public set color(value: HexColor) {
		this._color = value
	}

	public get color(): HexColor | undefined {
		return this._color
	}

	/**
	 * @type {boolean}
	 */
	private _hidden = false
	public set hidden(value: boolean) {
		this._hidden = value
	}

	public get hidden(): boolean {
		return this._hidden
	}

	/**
	 * @type {SlideNumberProps}
	 */
	public set slideNumber(value: SlideNumberProps) {
		// NOTE: Slide Numbers: In order for Slide Numbers to function they need to be in all 3 files: master/layout/slide
		this._slideNumberProps = value
		this._setSlideNum(value)
	}

	public get slideNumber(): SlideNumberProps | undefined {
		return this._slideNumberProps
	}

	public get newAutoPagedSlides(): PresSlide[] | undefined {
		return this._newAutoPagedSlides
	}

	/**
	 * Add chart to Slide
	 * @param {CHART_NAME|IChartMulti[]} type - chart type
	 * @param {object[]} data - data object
	 * @param {IChartOpts} options - chart options
	 * @return {Slide} this Slide
	 */
	addChart(type: CHART_NAME | CHARTEX_NAME | IChartMulti[], data: IOptsChartData[], options?: IChartOpts): Slide {
		// FUTURE: TODO-VERSION-4: Remove first arg - only take data and opts, with "type" required on opts
		genObj.addChartDefinition(this, type, Array.isArray(data) ? data.map(item => ({ ...item })) : data, cloneOpts(options))
		return this
	}

	/**
	 * Add image to Slide
	 * @param {ImageProps} options - image options
	 * @return {Slide} this Slide
	 */
	addImage(options: ImageProps): Slide {
		genObj.addImageDefinition(this, cloneOpts(options))
		return this
	}

	/**
	 * Add media (audio/video) to Slide
	 * @param {MediaProps} options - media options
	 * @return {Slide} this Slide
	 */
	addMedia(options: MediaProps): Slide {
		genObj.addMediaDefinition(this, cloneOpts(options))
		return this
	}

	/**
	 * Add speaker notes to Slide
	 * @docs https://gitbrent.github.io/PptxGenJS/docs/speaker-notes.html
	 * @param {string} notes - notes to add to slide
	 * @return {Slide} this Slide
	 */
	addNotes(notes: string): Slide {
		genObj.addNotesDefinition(this, notes)
		return this
	}

	/**
	 * Add a group (`p:grpSp`) to this slide. Child x/y are relative to the group origin.
	 * Existing addShape/addImage/addText APIs are unchanged.
	 * @example slide.addGroup({ x: 1, y: 1, w: 4, h: 3 }, g => { g.addShape(pptx.ShapeType.rect, { x: 0, y: 0, w: 1, h: 1 }) })
	 */
	addGroup(options: GroupProps, build?: (group: Group) => void): Slide {
		const children: ISlideObject[] = options._objects ? [...options._objects] : []
		if (build) build(genObj.createGroupBuilder(this, children))
		genObj.addGroupDefinition(this, cloneOpts(options), children)
		return this
	}

	/**
	 * Add a connector (`p:cxnSp`). Additive convenience over `line.isConnector`.
	 * Glue with `start`/`end` `objectName`s, or pass numeric `line.sourceId`/`targetId`.
	 * @example slide.addConnector({ type: 'straightConnector1', start: { shape: 'BoxA', site: 3 }, end: { shape: 'BoxB', site: 1 } })
	 */
	addConnector(options?: ConnectorProps): Slide {
		genObj.addConnectorDefinition(this, cloneOpts(options ?? {}))
		return this
	}

	/**
	 * Add a slide transition (`<p:transition>`, ECMA-376 §19.3.1.50 / MS-PPTX §2.2.1)
	 * @param {SlideTransitionProps} options - transition options
	 * @example slide.addTransition({ type: 'morph', duration: 800 })
	 * @return {Slide} this Slide
	 */
	addTransition(options: SlideTransitionProps): Slide {
		this.transition = options
		return this
	}

	/**
	 * Attach an object animation to the most recently added shape/text/image.
	 * Equivalent to passing `animation` on that object's options. Emitted as `<p:timing>`.
	 * @param {string | AnimationConfig} animation - preset name or config
	 * @example slide.addText('Hello', { x: 0.5, y: 0.5, w: 3, h: 0.5 }).addAnimation('fadein')
	 * @example slide.addAnimation({ type: 'flyin', direction: 'left', duration: 500 })
	 * @return {Slide} this Slide
	 */
	addAnimation(animation: string | AnimationConfig): Slide {
		const last = this._slideObjects[this._slideObjects.length - 1]
		if (!last) {
			throw new Error('addAnimation() requires a slide object; add text, a shape, or an image first')
		}
		last.options = last.options ?? {}
		last.options.animation = typeof animation === 'string' ? animation : { ...animation }
		return this
	}

	/**
	 * Add a threaded comment to this slide (MS-PPTX §2.16)
	 * @param {CommentProps} comment - comment options
	 * @example slide.addComment({ text: 'Review this', author: 'Ada', x: 1, y: 1 })
	 * @return {Slide} this Slide
	 */
	addComment(comment: CommentProps): Slide {
		this.comments = this.comments ?? []
		this.comments.push(comment)
		return this
	}

	/**
	 * Add a Slide Zoom navigation object (MS-PPTX §2.10)
	 * @param {ZoomProps} options - zoom options
	 * @example slide.addZoom({ slideNum: 3, x: 1, y: 4, w: 2, h: 1.13 })
	 * @return {Slide} this Slide
	 */
	addZoom(options: ZoomProps): Slide {
		genObj.addZoomDefinition(this, options)
		return this
	}

	/**
	 * Add a Section Zoom navigation object (MS-PPTX §2.9)
	 * @param {SectionZoomProps} options - section zoom options
	 * @example slide.addSectionZoom({ sectionTitle: 'Intro', x: 1, y: 4, w: 2, h: 1.13 })
	 * @return {Slide} this Slide
	 */
	addSectionZoom(options: SectionZoomProps): Slide {
		genObj.addSectionZoomDefinition(this, options)
		return this
	}

	/**
	 * Add a Summary Zoom navigation object (MS-PPTX §2.11)
	 * @param {SummaryZoomProps} options - summary zoom options
	 * @example slide.addSummaryZoom({ sectionTitle: 'Intro', x: 1, y: 4, w: 2, h: 1.13 })
	 * @return {Slide} this Slide
	 */
	addSummaryZoom(options: SummaryZoomProps): Slide {
		genObj.addSummaryZoomDefinition(this, options)
		return this
	}

	/**
	 * Embed a content part (MS-PPTX §2.2.3). Opt-in; `sp` fallback.
	 * @param {ContentPartProps} options - content-part options
	 * @example slide.addContentPart({ data: '<payload/>', x: 1, y: 1, w: 2, h: 1 })
	 * @return {Slide} this Slide
	 */
	addContentPart(options: ContentPartProps): Slide {
		genObj.addContentPartDefinition(this, cloneOpts(options))
		return this
	}

	/**
	 * Embed ink as a content part (MS-PPTX §2.2.3.1). Opt-in; `pic` fallback.
	 * @param {InkProps} options - ink options
	 * @example slide.addInk({ data: inkMl, cover: 'data:image/png;base64,...', x: 1, y: 1, w: 2, h: 1 })
	 * @return {Slide} this Slide
	 */
	addInk(options: InkProps): Slide {
		genObj.addInkDefinition(this, cloneOpts(options))
		return this
	}

	/**
	 * Reference an Office App (MS-PPTX §2.2.13 / MS-OWEXML webextensionref). Opt-in; `pic` fallback.
	 * @param {OfficeAppProps} options - Office App options
	 * @example slide.addOfficeApp({ reference: { id: 'WA00000', version: '1.0.0.0', storeType: 'OMEX' }, x: 1, y: 1, w: 3, h: 2 })
	 * @return {Slide} this Slide
	 */
	addOfficeApp(options: OfficeAppProps): Slide {
		genObj.addOfficeAppDefinition(this, cloneOpts(options))
		return this
	}

	/**
	 * Add shape to Slide
	 * @param {SHAPE_NAME} shapeName - shape name
	 * @param {ShapeProps} options - shape options
	 * @return {Slide} this Slide
	 */
	addShape(shapeName: SHAPE_NAME, options?: ShapeProps): Slide {
		// NOTE: As of v3.1.0, <script> users are passing the old shape object from the shapes file (orig to the project)
		// But React/TypeScript users are passing the shapeName from an enum, which is a simple string, so lets cast
		// <script./> => `pptx.shapes.RECTANGLE` [string] "rect" ... shapeName['name'] = 'rect'
		// TypeScript => `pptxgen.shapes.RECTANGLE` [string] "rect" ... shapeName = 'rect'
		// let shapeNameDecode = typeof shapeName === 'object' && shapeName['name'] ? shapeName['name'] : shapeName
		genObj.addShapeDefinition(this, shapeName, cloneOpts(options))
		return this
	}

	/**
	 * Add table to Slide
	 * @param {TableRow[]} tableRows - table rows
	 * @param {TableProps} options - table options
	 * @return {Slide} this Slide
	 */
	addTable(tableRows: TableRow[], options?: TableProps): Slide {
		// FUTURE: we pass `this` - we dont need to pass layouts - they can be read from this!
		this._newAutoPagedSlides = genObj.addTableDefinition(this, tableRows, cloneOpts(options), this._slideLayout, this._presLayout, this.addSlide, this.getSlide)
		return this
	}

	/**
	 * Add text to Slide
	 * @param {string|TextProps[]} text - text string or complex object
	 * @param {TextPropsOptions} options - text options
	 * @return {Slide} this Slide
	 */
	addText(text: string | TextProps[], options?: TextPropsOptions): Slide {
		// String/number text is one run. Keep text styling on that run, but leave
		// glow/softEdge/reflection/shadow/blur on the shape `effectLst` (issue #84).
		const textParam = typeof text === 'string' || typeof text === 'number' ? [{ text, options: cloneTextRunOpts(options) }] : text
		genObj.addTextDefinition(this, textParam, cloneOpts(options), false)
		return this
	}

	/**
	 * Add WordArt to Slide — text with a DrawingML warp (`a:prstTxWarp`) and/or a gradient run fill.
	 * Defaults to centered `textNoShape` (no transform); options override.
	 * @param {string|TextProps[]} text - text string or complex object
	 * @param {TextPropsOptions} options - text/WordArt options (`presetShape`, `gradient`, etc.)
	 * @return {Slide} this Slide
	 */
	addWordArt(text: string | TextProps[], options?: TextPropsOptions): Slide {
		const opts: TextPropsOptions = { align: 'center', presetShape: 'textNoShape', ...options }
		return this.addText(text, opts)
	}
}
