/**
 * PptxGenJS: Slide Object Generators
 */

import {
	BARCHART_COLORS,
	CHART_NAME,
	CHART_TYPE,
	DEF_CELL_BORDER,
	DEF_CELL_MARGIN_IN,
	DEF_CHART_BORDER,
	DEF_FONT_COLOR,
	DEF_FONT_SIZE,
	DEF_SHAPE_LINE_COLOR,
	DEF_SLIDE_MARGIN_IN,
	EMU,
	IMG_PLAYBTN,
	MASTER_OBJECTS,
	PIECHART_COLORS,
	SCHEME_COLOR_NAMES,
	SHAPE_NAME,
	SHAPE_TYPE,
	SLIDE_OBJECT_TYPES,
	TEXT_HALIGN,
	TEXT_VALIGN,
} from './core-enums'
import {
	AddSlideProps,
	BackgroundProps,
	BorderProps,
	Group,
	GroupProps,
	IChartMulti,
	IChartOptsLib,
	IOptsChartData,
	ISlideObject,
	ImageProps,
	MediaProps,
	ObjectOptions,
	ContentPartProps,
	InkProps,
	OfficeAppProps,
	SectionZoomProps,
	SummaryZoomProps,
	ZoomProps,
	OptsChartGridLine,
	PresLayout,
	PresSlide,
	ShapeFillProps,
	ShapeLineProps,
	ShapeProps,
	SlideLayout,
	SlideMasterProps,
	TableCell,
	TableProps,
	TableRow,
	TextProps,
	TextPropsOptions,
} from './core-interfaces'
import { getSlidesForTableRows } from './gen-tables'
import { applySvgFillToDataUrl, encodeXmlEntities, getNewRelId, getSmartParseNumber, inch2Emu, marginToEmu, resolveDataLabelPosition, valToPts, correctShadowOptions, warnDeprecatedOnce } from './gen-utils'
import { CT_INKML, CT_WEBEXTENSION, makeXmlWebExtension } from './xml/content-parts'

/** Valid OOXML preset-geometry strings (the values of `SHAPE_TYPE`) - anything else corrupts the file. */
const VALID_SHAPE_PRESETS = new Set<string>(Object.values(SHAPE_TYPE))

/** Type guard: true when `arr` is a flat (1-level) array rather than a nested one. */
function isFlatStringArray(arr: unknown): arr is string[] {
	return Array.isArray(arr) && (arr.length === 0 || !Array.isArray(arr[0]))
}

/** Unwrap master `text` / placeholder content so TextProps[] is not double-wrapped as a single run. */
function masterTextRuns (text: string | TextProps[] | undefined): TextProps[] {
	return Array.isArray(text) ? text : [{ text }]
}

/**
 * Transforms a slide definition to a slide object that is then passed to the XML transformation process.
 * @param {SlideMasterProps} props - slide definition
 * @param {PresSlide|SlideLayout} target - empty slide object that should be updated by the passed definition
 */
export function createSlideMaster(props: SlideMasterProps, target: SlideLayout): void {
	// STEP 1: Add background if either the slide or layout has background props
	// if (props.background || target.background) addBackgroundDefinition(props.background, target)
	if (props.bkgd) target.bkgd = props.bkgd // DEPRECATED: (remove in v4.0.0)

	// STEP 2: Add all Slide Master objects in the order they were given
	if (props.objects && Array.isArray(props.objects) && props.objects.length > 0) {
		props.objects.forEach((object, idx) => {
			const key = Object.keys(object)[0]
			if (MASTER_OBJECTS[key] && key === 'chart') addChartDefinition(target, object[key].type, object[key].data, object[key].opts)
			else if (MASTER_OBJECTS[key] && key === 'image') addImageDefinition(target, object[key])
			else if (MASTER_OBJECTS[key] && key === 'line') addShapeDefinition(target, SHAPE_TYPE.LINE, object[key])
			else if (MASTER_OBJECTS[key] && key === 'rect') addShapeDefinition(target, SHAPE_TYPE.RECTANGLE, object[key])
			else if (MASTER_OBJECTS[key] && key === 'roundRect') addShapeDefinition(target, SHAPE_TYPE.ROUNDED_RECTANGLE, object[key])
			else if (MASTER_OBJECTS[key] && key === 'shape') addShapeDefinition(target, object[key].type, object[key].options ?? {})
			else if (MASTER_OBJECTS[key] && key === 'media') addMediaDefinition(target as unknown as PresSlide, object[key])
			else if (MASTER_OBJECTS[key] && key === 'table') {
				// auto-paging makes no sense on a master - a master has no "next slide" to flow onto
				const tableProps: TableProps = { ...object[key].options, autoPage: false }
				const noPaging = (): never => {
					throw new Error('defineSlideMaster: tables on a master cannot auto-page')
				}
				addTableDefinition(target as unknown as PresSlide, object[key].rows, tableProps, target, target._presLayout, noPaging, noPaging)
			}
			else if (MASTER_OBJECTS[key] && key === 'text') addTextDefinition(target, masterTextRuns(object[key].text), object[key].options, false)
			else if (MASTER_OBJECTS[key] && key === 'placeholder') {
				// TODO: 20180820: Check for existing `name`?
				object[key].options.placeholder = object[key].options.name
				delete object[key].options.name // remap name for earier handling internally
				object[key].options._placeholderType = object[key].options.type
				delete object[key].options.type // remap name for earier handling internally
				object[key].options._placeholderIdx = 100 + idx
				addTextDefinition(target, masterTextRuns(object[key].text), object[key].options, true)
				// TODO: ISSUE#599 - only text is suported now (add more below)
				// else if (object[key].image) addImageDefinition(target, object[key].image)
				/* 20200120: So... image placeholders go into the "slideLayoutN.xml" file and addImage doesnt do this yet...
					<p:sp>
				  <p:nvSpPr>
					<p:cNvPr id="7" name="Picture Placeholder 6">
					  <a:extLst>
						<a:ext uri="{FF2B5EF4-FFF2-40B4-BE49-F238E27FC236}">
						  <a16:creationId xmlns:a16="http://schemas.microsoft.com/office/drawing/2014/main" id="{CE1AE45D-8641-0F4F-BDB5-080E69CCB034}"/>
						</a:ext>
					  </a:extLst>
					</p:cNvPr>
					<p:cNvSpPr>
				*/
			}
		})
	}

	// STEP 3: Add Slide Numbers (NOTE: Do this last so numbers are not covered by objects!)
	if (props.slideNumber && typeof props.slideNumber === 'object') target._slideNumberProps = props.slideNumber
}

/**
 * Generate the chart based on input data.
 * OOXML Chart Spec: ISO/IEC 29500-1:2016(E)
 *
 * @param {CHART_NAME | IChartMulti[]} `type` should belong to: 'column', 'pie'
 * @param {[]} `data` a JSON object with follow the following format
 * @param {IChartOptsLib} `opt` chart options
 * @param {PresSlide} `target` slide object that the chart will be added to
 * @return {object} chart object
 * {
 *    title: 'eSurvey chart',
 *    data: [
 *        {
 *            name: 'Income',
 *            labels: ['2005', '2006', '2007', '2008', '2009'],
 *            values: [23.5, 26.2, 30.1, 29.5, 24.6]
 *        },
 *        {
 *            name: 'Expense',
 *            labels: ['2005', '2006', '2007', '2008', '2009'],
 *            values: [18.1, 22.8, 23.9, 25.1, 25]
 *        }
 *    ]
 * }
 */
export function addChartDefinition(target: PresSlide | SlideLayout, type: CHART_NAME | IChartMulti[], data: IOptsChartData[], opt: IChartOptsLib): object {
	function correctGridLineOptions(glOpts: OptsChartGridLine): void {
		if (!glOpts || glOpts.style === 'none') return
		if (glOpts.size !== undefined && (isNaN(Number(glOpts.size)) || glOpts.size <= 0)) {
			console.warn('Warning: chart.gridLine.size must be greater than 0.')
			delete glOpts.size // delete prop to used defaults
		}
		if (glOpts.style && !['solid', 'dash', 'dot'].includes(glOpts.style)) {
			console.warn('Warning: chart.gridLine.style options: `solid`, `dash`, `dot`.')
			delete glOpts.style
		}
		if (glOpts.cap && !['flat', 'square', 'round'].includes(glOpts.cap)) {
			console.warn('Warning: chart.gridLine.cap options: `flat`, `square`, `round`.')
			delete glOpts.cap
		}
	}

	const chartId = target._allocChartId()
	const resultObject: ISlideObject = {
		_type: SLIDE_OBJECT_TYPES.chart,
		text: undefined,
		options: undefined,
		chartRid: undefined,
	}
	// DESIGN: `type` can an object (ex: `pptx.charts.DOUGHNUT`) or an array of chart objects
	// EX: addChartDefinition([ { type:pptx.charts.BAR, data:{name:'', labels:[], values[]} }, {<etc>} ])
	// Multi-Type Charts
	let tmpOpt: IChartOptsLib | IOptsChartData[]
	let tmpData: IOptsChartData[] = []
	if (Array.isArray(type)) {
		// For multi-type charts there needs to be data for each type,
		// as well as a single data source for non-series operations.
		// The data is indexed below to keep the data in order when segmented
		// into types.
		type.forEach(obj => {
			tmpData = tmpData.concat(obj.data)
		})
		// NOTE: the legacy multi-type call shape passed the options object as the 2nd arg (data lives inside the specs),
		// so only fall back to `data` when no real options were supplied (issue #25)
		tmpOpt = opt && Object.keys(opt).length > 0 ? opt : data
	} else {
		tmpData = data
		tmpOpt = opt
	}
	tmpData.forEach((item, i) => {
		item._dataIndex = i

		// Converts the 'labels' array from string[] to string[][] (or the respective primitive type), if needed.
		// NOTE: `labels` is typed `string[][]` but callers may pass a flat `string[]`; the type guard narrows that case.
		if (item.labels !== undefined && isFlatStringArray(item.labels)) {
			item.labels = [item.labels]
		}
	})
	const options: IChartOptsLib = tmpOpt && typeof tmpOpt === 'object' && !Array.isArray(tmpOpt) ? tmpOpt : {}

	// STEP 1: TODO: check for reqd fields, correct type, etc
	// `type` exists in CHART_TYPE
	// Array.isArray(data)
	/*
		if ( Array.isArray(rel.data) && rel.data.length > 0 && typeof rel.data[0] === 'object'
			&& rel.data[0].labels && Array.isArray(rel.data[0].labels)
			&& rel.data[0].values && Array.isArray(rel.data[0].values) ) {
			obj = rel.data[0];
		}
		else {
			console.warn("USAGE: addChart( 'pie', [ {name:'Sales', labels:['Jan','Feb'], values:[10,20]} ], {x:1, y:1} )");
			return;
		}
		*/

	// STEP 2: Set default options/decode user options
	// A: Core
	options._type = type
	const explicitChartX = typeof options.x !== 'undefined' && options.x != null && !isNaN(Number(options.x))
	const explicitChartY = typeof options.y !== 'undefined' && options.y != null && !isNaN(Number(options.y))
	const explicitChartW = options.w !== undefined && options.w !== null
	const explicitChartH = options.h !== undefined && options.h !== null
	options.x = explicitChartX ? options.x : 1
	options.y = explicitChartY ? options.y : 1
	options.w = options.w || '50%'
	options.h = options.h || '50%'
	options.objectName = options.objectName
		? encodeXmlEntities(options.objectName)
		: `Chart ${target._slideObjects.filter(obj => obj._type === SLIDE_OBJECT_TYPES.chart).length}`

	// B: Options: misc
	if (!['bar', 'col'].includes(options.barDir || '')) options.barDir = 'col'

	// barGrouping: "21.2.3.17 ST_Grouping (Grouping)"
	// barGrouping must be handled before data label validation as it can affect valid label positioning
	if (options._type === CHART_TYPE.AREA) {
		if (!['stacked', 'standard', 'percentStacked'].includes(options.barGrouping || '')) options.barGrouping = 'standard'
	}
	if (options._type === CHART_TYPE.BAR) {
		if (!['clustered', 'stacked', 'percentStacked'].includes(options.barGrouping || '')) options.barGrouping = 'clustered'
	}
	if (options._type === CHART_TYPE.BAR3D) {
		if (!['clustered', 'stacked', 'standard', 'percentStacked'].includes(options.barGrouping || '')) options.barGrouping = 'standard'
	}
	if (options.barGrouping?.includes('tacked')) {
		if (!options.barGapWidthPct) options.barGapWidthPct = 50
	}
	// Translate friendly data label positions to OOXML codes and drop values the chart type rejects
	// NOTE: an unsupported `c:dLblPos` makes PowerPoint declare the file corrupt, so it is dropped with a warning
	// NOTE: a multi-type chart has no single type to validate against, so only the name translation applies there
	if (options.dataLabelPosition) {
		const chartType = Array.isArray(options._type) ? undefined : options._type
		options.dataLabelPosition = resolveDataLabelPosition(options.dataLabelPosition, chartType, options.barGrouping) as typeof options.dataLabelPosition
		if (!options.dataLabelPosition) delete options.dataLabelPosition
	}
	options.dataLabelBkgrdColors = options.dataLabelBkgrdColors || !options.dataLabelBkgrdColors ? options.dataLabelBkgrdColors : false
	if (!['b', 'l', 'r', 't', 'tr'].includes(options.legendPos || '')) options.legendPos = 'r'

	// 3D bar: ST_Shape
	if (!['cone', 'coneToMax', 'box', 'cylinder', 'pyramid', 'pyramidToMax'].includes(options.bar3DShape || '')) options.bar3DShape = 'box'
	// lineDataSymbol: http://www.datypic.com/sc/ooxml/a-val-32.html
	// Spec has [plus,star,x] however neither PPT2013 nor PPT-Online support them
	if (!['circle', 'dash', 'diamond', 'dot', 'none', 'square', 'triangle'].includes(options.lineDataSymbol || '')) options.lineDataSymbol = 'circle'
	if (!['gap', 'span'].includes(options.displayBlanksAs || '')) options.displayBlanksAs = 'span'
	if (!['standard', 'marker', 'filled'].includes(options.radarStyle || '')) options.radarStyle = 'standard'
	options.lineDataSymbolSize = options.lineDataSymbolSize && !isNaN(options.lineDataSymbolSize) ? options.lineDataSymbolSize : 6
	options.lineDataSymbolLineSize = options.lineDataSymbolLineSize && !isNaN(options.lineDataSymbolLineSize) ? valToPts(options.lineDataSymbolLineSize) : valToPts(0.75)
	// `layout` allows the override of PPT defaults to maximize space
	if (options.layout) {
		const layout = options.layout
		;['x', 'y', 'w', 'h'].forEach(key => {
			const val = layout[key]
			if (isNaN(Number(val)) || val < 0 || val > 1) {
				console.warn('Warning: chart.layout.' + key + ' can only be 0-1')
				delete layout[key] // remove invalid value so that default will be used
			}
		})
	}

	// Set gridline defaults
	options.catGridLine = options.catGridLine || (options._type === CHART_TYPE.SCATTER ? { color: 'D9D9D9', size: 1 } : { style: 'none' })
	options.valGridLine = options.valGridLine || (options._type === CHART_TYPE.SCATTER ? { color: 'D9D9D9', size: 1 } : {})
	options.serGridLine = options.serGridLine || (options._type === CHART_TYPE.SCATTER ? { color: 'D9D9D9', size: 1 } : { style: 'none' })
	correctGridLineOptions(options.catGridLine)
	correctGridLineOptions(options.valGridLine)
	correctGridLineOptions(options.serGridLine)
	if (options.shadow) options.shadow = correctShadowOptions(options.shadow)

	// C: Options: plotArea
	options.showDataTable = options.showDataTable || !options.showDataTable ? options.showDataTable : false
	options.showDataTableHorzBorder = options.showDataTableHorzBorder || !options.showDataTableHorzBorder ? options.showDataTableHorzBorder : true
	options.showDataTableVertBorder = options.showDataTableVertBorder || !options.showDataTableVertBorder ? options.showDataTableVertBorder : true
	options.showDataTableOutline = options.showDataTableOutline || !options.showDataTableOutline ? options.showDataTableOutline : true
	options.showDataTableKeys = options.showDataTableKeys || !options.showDataTableKeys ? options.showDataTableKeys : true
	options.showLabel = options.showLabel || !options.showLabel ? options.showLabel : false
	options.showLegend = options.showLegend || !options.showLegend ? options.showLegend : false
	options.showPercent = options.showPercent || !options.showPercent ? options.showPercent : true
	options.showTitle = options.showTitle || !options.showTitle ? options.showTitle : false
	options.showValue = options.showValue || !options.showValue ? options.showValue : false
	options.showLeaderLines = options.showLeaderLines || !options.showLeaderLines ? options.showLeaderLines : false
	options.catAxisLineShow = typeof options.catAxisLineShow !== 'undefined' ? options.catAxisLineShow : true
	options.valAxisLineShow = typeof options.valAxisLineShow !== 'undefined' ? options.valAxisLineShow : true
	options.serAxisLineShow = typeof options.serAxisLineShow !== 'undefined' ? options.serAxisLineShow : true

	const v3DRotX = options.v3DRotX
	options.v3DRotX = v3DRotX !== undefined && !isNaN(v3DRotX) && v3DRotX >= -90 && v3DRotX <= 90 ? v3DRotX : 30
	const v3DRotY = options.v3DRotY
	options.v3DRotY = v3DRotY !== undefined && !isNaN(v3DRotY) && v3DRotY >= 0 && v3DRotY <= 360 ? v3DRotY : 30
	options.v3DRAngAx = options.v3DRAngAx || !options.v3DRAngAx ? options.v3DRAngAx : true
	const v3DPerspective = options.v3DPerspective
	options.v3DPerspective = v3DPerspective !== undefined && !isNaN(v3DPerspective) && v3DPerspective >= 0 && v3DPerspective <= 240 ? v3DPerspective : 30

	// D: Options: chart
	const barGapWidthPct = options.barGapWidthPct
	options.barGapWidthPct = barGapWidthPct !== undefined && !isNaN(barGapWidthPct) && barGapWidthPct >= 0 && barGapWidthPct <= 1000 ? barGapWidthPct : 150
	const barGapDepthPct = options.barGapDepthPct
	options.barGapDepthPct = barGapDepthPct !== undefined && !isNaN(barGapDepthPct) && barGapDepthPct >= 0 && barGapDepthPct <= 1000 ? barGapDepthPct : 150

	options.chartColors = Array.isArray(options.chartColors)
		? options.chartColors
		: options._type === CHART_TYPE.PIE || options._type === CHART_TYPE.DOUGHNUT
			? PIECHART_COLORS
			: BARCHART_COLORS
	options.chartColorsOpacity = options.chartColorsOpacity && !isNaN(options.chartColorsOpacity) ? options.chartColorsOpacity : undefined
	// DEPRECATED: v3.11.0 - use `plotArea.border` vvv
	options.border = options.border && typeof options.border === 'object' ? options.border : undefined
	if (options.border) options.border.pt = options.border.width ?? (!options.border.pt || isNaN(options.border.pt) ? DEF_CHART_BORDER.pt : options.border.pt)
	if (options.border && (!options.border.color || typeof options.border.color !== 'string')) options.border.color = DEF_CHART_BORDER.color
	// DEPRECATED: (remove above in v4.0) ^^^
	options.plotArea = options.plotArea || {}
	options.plotArea.border = options.plotArea.border && typeof options.plotArea.border === 'object' ? options.plotArea.border : undefined
	if (options.plotArea.border) {
		options.plotArea.border.pt = options.plotArea.border.width ?? (!options.plotArea.border.pt || isNaN(options.plotArea.border.pt) ? DEF_CHART_BORDER.pt : options.plotArea.border.pt)
	}
	if (options.plotArea.border && (!options.plotArea.border.color || typeof options.plotArea.border.color !== 'string')) { options.plotArea.border.color = DEF_CHART_BORDER.color }
	if (options.border) {
		warnDeprecatedOnce('chart-border', 'chart `border` is deprecated - use `plotArea.border`')
		options.plotArea.border = options.border // @deprecated [[remove in v4.0]]
	}
	options.plotArea.fill = options.plotArea.fill || { color: undefined, transparency: undefined }
	if (options.fill && options.plotArea.fill) {
		warnDeprecatedOnce('chart-fill', 'chart `fill` is deprecated - use `plotArea.fill`')
		options.plotArea.fill.color = options.fill // @deprecated [[remove in v4.0]]
	}
	//
	options.chartArea = options.chartArea || {}
	options.chartArea.border = options.chartArea.border && typeof options.chartArea.border === 'object' ? options.chartArea.border : undefined
	if (options.chartArea.border) {
		options.chartArea.border = {
			color: options.chartArea.border.color || DEF_CHART_BORDER.color,
			pt: options.chartArea.border.width ?? options.chartArea.border.pt ?? DEF_CHART_BORDER.pt,
		}
	}
	options.chartArea.roundedCorners = typeof options.chartArea.roundedCorners === 'boolean' ? options.chartArea.roundedCorners : true
	//
	options.dataBorder = options.dataBorder && typeof options.dataBorder === 'object' ? options.dataBorder : undefined
	if (options.dataBorder) options.dataBorder.pt = options.dataBorder.width ?? (!options.dataBorder.pt || isNaN(options.dataBorder.pt) ? 0.75 : options.dataBorder.pt)
	if (options.dataBorder && options.dataBorder.color) {
		const borderColor = options.dataBorder.color
		const isHexColor = typeof borderColor === 'string' && borderColor.length === 6 && /^[0-9A-Fa-f]{6}$/.test(borderColor)
		const isSchemeColor = Object.values(SCHEME_COLOR_NAMES).some(schemeColor => schemeColor === borderColor)
		if (!isHexColor && !isSchemeColor) {
			options.dataBorder.color = 'F9F9F9' // Fallback if neither hex nor scheme color
		}
	}
	//
	if (!options.dataLabelFormatCode && options._type === CHART_TYPE.SCATTER) options.dataLabelFormatCode = 'General'
	if (!options.dataLabelFormatCode && (options._type === CHART_TYPE.PIE || options._type === CHART_TYPE.DOUGHNUT)) { options.dataLabelFormatCode = options.showPercent ? '0%' : 'General' }
	options.dataLabelFormatCode = options.dataLabelFormatCode && typeof options.dataLabelFormatCode === 'string' ? options.dataLabelFormatCode : '#,##0'
	//
	// Set default format for Scatter chart labels to custom string if not defined
	if (!options.dataLabelFormatScatter && options._type === CHART_TYPE.SCATTER) options.dataLabelFormatScatter = 'custom'
	//
	options.lineSize = typeof options.lineSize === 'number' ? options.lineSize : 2
	options.valAxisMajorUnit = typeof options.valAxisMajorUnit === 'number' ? options.valAxisMajorUnit : undefined

	if (options._type === CHART_TYPE.AREA || options._type === CHART_TYPE.BAR || options._type === CHART_TYPE.BAR3D || options._type === CHART_TYPE.LINE) {
		options.catAxisMultiLevelLabels = !!options.catAxisMultiLevelLabels
	} else {
		delete options.catAxisMultiLevelLabels
	}

	// STEP 4: Set props (_type already set to chart on the literal above).
	// The chart object stores the position/placeholder that gen-xml reads; chart `fill`
	// is a color string (vs ObjectOptions' ShapeFillProps) and is never read here, so it
	// is reconciled to undefined to keep the stored object a valid ObjectOptions.
	resultObject.options = {
		...options,
		fill: undefined,
		_explicitX: explicitChartX,
		_explicitY: explicitChartY,
		_explicitW: explicitChartW,
		_explicitH: explicitChartH,
	}
	resultObject.chartRid = getNewRelId(target)

	// STEP 5: Add this chart to this Slide Rels (rId/rels count spans all slides! Count all images to get next rId)
	target._relsChart.push({
		rId: getNewRelId(target),
		data: tmpData,
		opts: options,
		type: options._type,
		globalId: chartId,
		fileName: `chart${chartId}.xml`,
		Target: `/ppt/charts/chart${chartId}.xml`,
	})

	target._slideObjects.push(resultObject)
	return resultObject
}

/**
 * Adds an image object to a slide definition.
 * This method can be called with only two args (opt, target) - this is supposed to be the only way in future.
 * @param {ImageProps} `opt` - object containing `path`/`data`, `x`, `y`, etc.
 * @param {PresSlide} `target` - slide that the image should be added to (if not specified as the 2nd arg)
 * @note: Remote images (eg: "http://whatev.com/blah"/from web and/or remote server arent supported yet - we'd need to create an <img>, load it, then send to canvas
 * @see: https://stackoverflow.com/questions/164181/how-to-fetch-a-remote-image-to-display-in-a-canvas)
 */
/**
 * Media filename stem. Master/layout `_slideNum` values overlap the slide-number range (masters start at
 * 1000 and slides also reach 1000+), so a layout image `image-1002-1` would collide with slide #1002's
 * `image-1002-1` and one overwrites the other in the zip (issue #1416). Prefix layout/master media with
 * `layout-` so the two namespaces can never collide; the slide namespace keeps its plain `image-<n>-<i>`.
 */
function imageBorderToLine (border: ImageProps['border']): ShapeLineProps | undefined {
	if (!border) return undefined
	const side = Array.isArray(border) ? border[0] : border
	if (!side || side.type === 'none') return undefined
	return { color: side.color, width: side.width ?? side.pt }
}

function imageTargetStem(target: PresSlide | SlideLayout): string {
	// PresSlide has `_rId`; SlideLayout/master does not. Masters number from 1000 and collide with slide #1000+.
	if (!('_rId' in target)) return `layout-image-${target._relsMedia.length + 1}`
	return `image-${target._slideNum}-${target._relsMedia.length + 1}`
}

function isSvgFillSource (data: string, path: string): boolean {
	const src = (data || path).toLowerCase()
	return src.includes('image/svg+xml') || src.endsWith('.svg') || src.includes('.svg?')
}

/**
 * Register a table-cell image fill as slide media so `a:blipFill` can embed it
 * (Dominik-von-Burg PR 1461). Must run at addTable time so encodeSlideMediaRels sees the rels.
 */
function registerTableCellImageFill (target: PresSlide, fill: ShapeFillProps | undefined): number | undefined {
	const data = typeof fill?.data === 'string' ? fill.data : ''
	const path = typeof fill?.path === 'string' ? fill.path : ''
	if (!data && !path) return undefined

	const isSvg = isSvgFillSource(data, path)
	if (isSvg) {
		const pngRid = getNewRelId(target)
		target._relsMedia.push({
			path: path || data + 'png',
			type: 'image/png',
			extn: 'png',
			data: data || '',
			rId: pngRid,
			isDuplicate: false,
			isSvgPng: true,
			Target: `../media/${imageTargetStem(target)}.png`,
		})
		const svgRid = getNewRelId(target)
		target._relsMedia.push({
			path: path || data,
			type: 'image/svg+xml',
			extn: 'svg',
			data: data || '',
			rId: svgRid,
			isDuplicate: false,
			Target: `../media/${imageTargetStem(target)}.svg`,
		})
		return svgRid
	}

	let extn = (path.split('?')[0].split('#')[0].split('.').pop() || 'png').toLowerCase()
	const mime = data ? /image\/([\w+]+);/.exec(data) : null
	if (mime?.[1]) extn = mime[1] === 'svg+xml' ? 'svg' : mime[1]
	if (extn === 'jpg') extn = 'jpeg'
	const imgType = 'image/' + extn
	const dupeItem = target._relsMedia.filter(item =>
		item.type === imgType && !item.isDuplicate && !item.isSvgPng && (
			(path && item.path === path) ||
			(data && item.data && item.data === data)
		)
	)[0]
	const rId = getNewRelId(target)
	target._relsMedia.push({
		path: path || 'preencoded.' + extn,
		type: imgType,
		extn,
		data: data || '',
		rId,
		isDuplicate: !!(dupeItem?.Target),
		Target: dupeItem?.Target ? dupeItem.Target : `../media/${imageTargetStem(target)}.${extn}`,
	})
	return rId
}

export function addImageDefinition(target: PresSlide | SlideLayout, opt: ImageProps): void {
	const newObject: ISlideObject = {
		_type: SLIDE_OBJECT_TYPES.image,
		text: undefined,
		options: undefined,
		image: undefined,
		imageRid: undefined,
		hyperlink: undefined,
	}
	// FIRST: Set vars for this image (object param replaces positional args in 1.1.0)
	const intPosX = opt.x || 0
	const intPosY = opt.y || 0
	const intWidth = opt.w || 0
	const intHeight = opt.h || 0
	const sizing = opt.sizing || null
	const objHyperlink = opt.hyperlink || ''
	let strImageData = opt.data || ''
	const strImagePath = opt.path || ''
	let imageRelId = getNewRelId(target)
	const objectName = opt.objectName ? encodeXmlEntities(opt.objectName) : `Image ${target._slideObjects.filter(obj => obj._type === SLIDE_OBJECT_TYPES.image).length}`

	// REALITY-CHECK:
	if (!strImagePath && !strImageData) {
		console.error('ERROR: addImage() requires either \'data\' or \'path\' parameter!')
		return
	} else if (strImagePath && typeof strImagePath !== 'string') {
		console.error(`ERROR: addImage() 'path' should be a string, ex: {path:'/img/sample.png'} - you sent ${String(strImagePath)}`)
		return
	} else if (strImageData && typeof strImageData !== 'string') {
		console.error(`ERROR: addImage() 'data' should be a string, ex: {data:'image/png;base64,NMP[...]'} - you sent ${String(strImageData)}`)
		return
	} else if (strImageData && typeof strImageData === 'string' && !strImageData.toLowerCase().includes('base64,')) {
		console.error('ERROR: Image `data` value lacks a base64 header! Ex: \'image/png;base64,NMP[...]\')')
		return
	}

	// STEP 1: Set extension
	// NOTE: Split to address URLs with params (eg: `path/brent.jpg?someParam=true`)
	let strImgExtn = (
		strImagePath
			.substring(strImagePath.lastIndexOf('/') + 1)
			.split('?')[0]
			.split('.')
			.pop()
			?.split('#')[0] || 'png'
	).toLowerCase()

	// However, pre-encoded images can be whatever mime-type they want (and good for them!)
	const strImageMimeMatch = strImageData ? /image\/(\w+);/.exec(strImageData) : null
	if (strImageMimeMatch && strImageMimeMatch.length > 0) {
		strImgExtn = strImageMimeMatch[1]
	} else if (strImageData?.toLowerCase().includes('image/svg+xml')) {
		strImgExtn = 'svg'
	}

	// STEP 2: Set type/path
	newObject._type = SLIDE_OBJECT_TYPES.image
	newObject.image = strImagePath || 'preencoded.png'

	// STEP 3: Set image properties & options
	// NOTE: when neither `w` nor `h` is given, the 1x1 inch default below is replaced with the image's
	// natural size during export (see `applyNaturalImageSizes`) - the bytes aren't loaded yet at this point
	newObject.options = {
		x: intPosX || 0,
		y: intPosY || 0,
		w: intWidth || 1,
		h: intHeight || 1,
		altText: opt.altText || '',
		rounding: typeof opt.rounding === 'boolean' ? opt.rounding : false,
		rectRadius: opt.rectRadius,
		sizing: sizing ?? undefined,
		placeholder: opt.placeholder,
		rotate: opt.rotate || 0,
		flipV: opt.flipV || false,
		flipH: opt.flipH || false,
		transparency: opt.transparency || 0,
		objectName,
		shadow: opt.shadow ? correctShadowOptions(opt.shadow) : undefined,
		glow: opt.glow,
		softEdge: opt.softEdge,
		reflection: opt.reflection,
		blur: opt.blur,
		line: opt.line || imageBorderToLine(opt.border),
		_sizeFromImage: !intWidth && !intHeight,
		// addImage always writes x/y/w/h defaults, so remember which axes the caller actually set
		// (placeholder geometry must not overwrite explicit placement; fujita-h d7e3e93 / #996).
		_explicitX: opt.x !== undefined && opt.x !== null,
		_explicitY: opt.y !== undefined && opt.y !== null,
		_explicitW: opt.w !== undefined && opt.w !== null,
		_explicitH: opt.h !== undefined && opt.h !== null,
		// Images build options explicitly (unlike shape/text which pass through opts), so copy animation here
		animation: opt.animation,
		appearOnClick: opt.appearOnClick,
	}

	// STEP 4: Add this image to this Slide Rels (rId/rels count spans all slides! Count all images to get next rId)
	if (strImgExtn === 'svg') {
		// SVG files consume *TWO* rId's: (a png version and the svg image)
		// <Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image1.png"/>
		// <Relationship Id="rId4" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="../media/image2.svg"/>
		// Martin-N: optional SVG recolor via `fill.color` (hex). Bake into data when already loaded; otherwise stash on the rel for encodeSlideMediaRels.
		const svgFillColor = typeof opt.fill?.color === 'string' ? opt.fill.color : undefined
		if (svgFillColor && strImageData) {
			strImageData = applySvgFillToDataUrl(strImageData, svgFillColor)
		}
		target._relsMedia.push({
			path: strImagePath || strImageData + 'png',
			type: 'image/png',
			extn: 'png',
			data: strImageData || '',
			rId: imageRelId,
			Target: `../media/${imageTargetStem(target)}.png`,
			isSvgPng: true,
			fill: opt.fill,
			svgSize: { w: getSmartParseNumber(newObject.options?.w, 'X', target._presLayout), h: getSmartParseNumber(newObject.options?.h, 'Y', target._presLayout) },
		})
		newObject.imageRid = imageRelId
		target._relsMedia.push({
			path: strImagePath || strImageData,
			type: 'image/svg+xml',
			extn: strImgExtn,
			data: strImageData || '',
			rId: imageRelId + 1,
			Target: `../media/${imageTargetStem(target)}.${strImgExtn}`,
			fill: opt.fill,
		})
		newObject.imageRid = imageRelId + 1
		imageRelId++ // NOTE: the SVG branch consumed two rIds - keep the counter on the last one used (issue #19)
	} else {
		// PERF: Duplicate media should reuse existing `Target` value and not create an additional copy.
		// OPC (ECMA-376 Part 2) lets many relationships point at one part, so identical image bytes are stored once.
		// Match on the source path when given, else on identical base64 `data` (issue #1339) - same content, one media file.
		const imgType = 'image/' + strImgExtn
		const dupeItem = target._relsMedia.filter(item =>
			item.type === imgType && !item.isDuplicate && !item.isSvgPng && (
				(strImagePath && item.path === strImagePath) ||
				(strImageData && item.data && item.data === strImageData)
			)
		)[0]

		target._relsMedia.push({
			path: strImagePath || 'preencoded.' + strImgExtn,
			type: imgType,
			extn: strImgExtn,
			data: strImageData || '',
			rId: imageRelId,
			isDuplicate: !!(dupeItem?.Target),
			Target: dupeItem?.Target ? dupeItem.Target : `../media/${imageTargetStem(target)}.${strImgExtn}`,
		})
		newObject.imageRid = imageRelId
	}

	// STEP 5: Hyperlink support
	if (typeof objHyperlink === 'object') {
		if (!objHyperlink.url && !objHyperlink.slide) throw new Error('ERROR: `hyperlink` option requires either: `url` or `slide`')
		else {
			imageRelId++

			target._rels.push({
				type: SLIDE_OBJECT_TYPES.hyperlink,
				data: objHyperlink.slide ? 'slide' : 'dummy',
				rId: imageRelId,
				Target: encodeXmlEntities(objHyperlink.url || objHyperlink.slide?.toString() || ''),
			})

			objHyperlink._rId = imageRelId
			newObject.hyperlink = objHyperlink
		}
	}

	// STEP 6: Add object to slide
	target._slideObjects.push(newObject)
}

/**
 * Adds a media object to a slide definition.
 * @param {PresSlide} `target` - slide object that the media will be added to
 * @param {MediaProps} `opt` - media options
 */
export function addMediaDefinition(target: PresSlide, opt: MediaProps): void {
	const intPosX = opt.x || 0
	const intPosY = opt.y || 0
	const intSizeX = opt.w || 2
	const intSizeY = opt.h || 2
	const strData = opt.data || ''
	const strLink = opt.link || ''
	const strPath = opt.path || ''
	const strType = opt.type || 'audio'
	let strExtn = ''
	const strCover = opt.cover || IMG_PLAYBTN
	const objectName = opt.objectName ? encodeXmlEntities(opt.objectName) : `Media ${target._slideObjects.filter(obj => obj._type === SLIDE_OBJECT_TYPES.media).length}`
	const slideData: ISlideObject = { _type: SLIDE_OBJECT_TYPES.media }

	// `link` with no `path`/`data` used to throw, so treating it as "reference, do not embed" cannot
	// change any deck that works today
	const isLinked = !!strLink && !strPath && !strData && (strType === 'audio' || strType === 'video')
	if (strLink && (strPath || strData) && strType !== 'online') {
		console.warn('[pptxgenjs] addMedia: `link` is ignored when `path` or `data` is given - the media is embedded')
	}

	// STEP 1: REALITY-CHECK - what each media kind needs, rather than one chain of conditions
	if (strType === 'audioCd') {
		// `a:st`/`a:end` and their `@track` are all required by CT_AudioCD
		if (typeof opt.audioCd?.start?.track !== 'number' || typeof opt.audioCd?.end?.track !== 'number') {
			throw new Error('addMedia() error: `type:"audioCd"` requires `audioCd.start.track` and `audioCd.end.track`')
		}
	} else if (!strPath && !strData && strType !== 'online' && !isLinked) {
		throw new Error('addMedia() error: either `data` or `path` are required!')
	} else if (strData && !strData.toLowerCase().includes('base64,')) {
		throw new Error('addMedia() error: `data` value lacks a base64 header! Ex: \'video/mpeg;base64,NMP[...]\')')
	}
	if (strCover && !strCover.toLowerCase().includes('base64,')) {
		throw new Error('addMedia() error: `cover` value lacks a base64 header! Ex: \'data:image/png;base64,iV[...]\')')
	}
	// Online Video: requires `link`
	if (strType === 'online' && !strLink) {
		throw new Error('addMedia() error: online videos require `link` value')
	}
	// Timing-tree playback (ECMA-376 §19.5 CT_TLMediaNode): no silent fallbacks.
	// `fullScrn` exists only on CT_TLMediaNodeVideo; linked `online` media has no embedded media node.
	if (opt.fullScreen && strType !== 'video') {
		throw new Error('addMedia() error: `fullScreen` is only valid on type "video"')
	}
	if (strType === 'online' && (opt.autoplay || opt.loop || opt.fullScreen || opt.mute)) {
		throw new Error('addMedia() error: autoplay/loop/fullScreen/mute are not valid on type "online"')
	}

	// FIXME: 20190707
	// strType = strData ? strData.split(';')[0].split('/')[0] : strType
	strExtn = opt.extn || (strType === 'wav' ? 'wav' : strData ? strData.split(';')[0].split('/')[1] : (strPath || strLink).split('.').pop()) || 'mp3'

	// STEP 2: Set type, media
	slideData.mtype = strType
	slideData.media = strPath || 'preencoded.mov'
	slideData.options = {}

	// STEP 3: Set media properties & options
	slideData.options.x = intPosX
	slideData.options.y = intPosY
	slideData.options.w = intSizeX
	slideData.options.h = intSizeY
	slideData.options._explicitX = opt.x !== undefined && opt.x !== null
	slideData.options._explicitY = opt.y !== undefined && opt.y !== null
	slideData.options._explicitW = opt.w !== undefined && opt.w !== null
	slideData.options._explicitH = opt.h !== undefined && opt.h !== null
	slideData.options.objectName = objectName
	slideData.options.animation = opt.animation
	slideData.options.appearOnClick = opt.appearOnClick
	// MS-PPTX §2.3.3.14 media extras + §2.2.14 narration flag
	slideData.options.trim = opt.trim
	slideData.options.fade = opt.fade
	slideData.options.bookmarks = opt.bookmarks
	slideData.options.isNarration = opt.isNarration
	// Playback behaviour (drives the slide timing tree — ECMA-376 §19.5 CT_TLMediaNode)
	slideData.options.autoplay = opt.autoplay
	slideData.options.loop = opt.loop
	slideData.options.fullScreen = opt.fullScreen
	slideData.options.mute = opt.mute
	slideData.options.contentType = opt.contentType
	slideData.options.audioCd = opt.audioCd
	slideData.options.isPhoto = opt.isPhoto
	slideData.options.userDrawn = opt.userDrawn
	slideData.options.isLinked = isLinked

	// STEP 4: Add this media to this Slide Rels (rId/rels count spans all slides! Count all media to get next rId)
	/**
	 * NOTE:
	 * - rId starts at 2 (hence the intRels+1 below) as slideLayout.xml is rId=1!
	 *
	 * NOTE:
	 * - Audio/Video files consume *TWO* rId's:
	 * <Relationship Id="rId2" Target="../media/media1.mov" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"/>
	 * <Relationship Id="rId3" Target="../media/media1.mov" Type="http://schemas.microsoft.com/office/2007/relationships/media"/>
	 */
	if (strType === 'audioCd') {
		// CD audio references the listener's drive: no media part, no media relationship, cover only
		const coverRid = getNewRelId(target)
		target._relsMedia.push({
			path: 'preencoded.png',
			type: 'image/png',
			extn: 'png',
			data: strCover,
			rId: coverRid,
			Target: `../media/${imageTargetStem(target)}.png`,
		})
		slideData._coverRid = coverRid
	} else if (strType === 'wav') {
		// `a:wavAudioFile` is the legacy embedded-WAV element: one audio relationship, no `p14:media`
		const relId1 = getNewRelId(target)
		target._relsMedia.push({
			path: strPath || 'preencoded.wav',
			type: 'audio/wav',
			extn: 'wav',
			data: strData || '',
			rId: relId1,
			Target: `../media/media-${target._slideNum}-${target._relsMedia.length + 1}.wav`,
		})
		slideData.mediaRid = relId1
		const coverRid = getNewRelId(target)
		target._relsMedia.push({
			path: 'preencoded.png',
			type: 'image/png',
			extn: 'png',
			data: strCover,
			rId: coverRid,
			Target: `../media/${imageTargetStem(target)}.png`,
		})
		slideData._coverRid = coverRid
	} else if (strType === 'online') {
		const relId1 = getNewRelId(target)
		// A: Add video
		target._relsMedia.push({
			path: strPath || 'preencoded' + strExtn,
			data: 'dummy',
			type: 'online',
			extn: strExtn,
			rId: relId1,
			Target: strLink,
		})
		slideData.mediaRid = relId1

		// B: Add cover (preview/overlay) image
		target._relsMedia.push({
			path: 'preencoded.png',
			data: strCover,
			type: 'image/png',
			extn: 'png',
			rId: getNewRelId(target),
			Target: `../media/${imageTargetStem(target)}.png`,
		})
	} else {
		// PERF: Duplicate media should reuse existing `Target` value and not create an additional copy
		const dupeItem = target._relsMedia.filter(item => item.path && item.path === strPath && item.type === strType + '/' + strExtn && !item.isDuplicate)[0]

		// A: "relationships/video"
		// Linked media keeps this exact three-relationship shape - the rIds the emitter derives from
		// `mediaRid` stay valid - and only points the targets outside the package
		const relId1 = getNewRelId(target)
		target._relsMedia.push({
			path: strPath || 'preencoded' + strExtn,
			type: strType + '/' + strExtn,
			extn: strExtn,
			data: strData || '',
			rId: relId1,
			isDuplicate: !!(dupeItem?.Target),
			isLinked,
			Target: isLinked ? strLink : (dupeItem?.Target ? dupeItem.Target : `../media/media-${target._slideNum}-${target._relsMedia.length + 1}.${strExtn}`),
		})
		slideData.mediaRid = relId1

		// B: "relationships/media"
		target._relsMedia.push({
			path: strPath || 'preencoded' + strExtn,
			type: strType + '/' + strExtn,
			extn: strExtn,
			data: strData || '',
			rId: getNewRelId(target),
			isDuplicate: !!(dupeItem?.Target),
			isLinked,
			Target: isLinked ? strLink : (dupeItem?.Target ? dupeItem.Target : `../media/media-${target._slideNum}-${target._relsMedia.length + 0}.${strExtn}`),
		})

		// C: Add cover (preview/overlay) image
		target._relsMedia.push({
			path: 'preencoded.png',
			type: 'image/png',
			extn: 'png',
			data: strCover,
			rId: getNewRelId(target),
			Target: `../media/${imageTargetStem(target)}.png`,
		})
	}

	// LAST
	target._slideObjects.push(slideData)
}

/**
 * Adds a Slide Zoom navigation object to a slide (MS-PPTX §2.10).
 * Emits an `mc:AlternateContent` block: a `p16:sldZm` Choice plus a `pic` Fallback so older readers
 * still render the thumbnail. The cover/thumbnail image is a media rel on the slide.
 * @param {PresSlide} `target` slide object
 * @param {ZoomProps} `opt` zoom options
 */
export function addZoomDefinition(target: PresSlide, opt: ZoomProps): void {
	pushZoomObject(target, 'slide', opt, obj => {
		obj.zoomSlideNum = opt.slideNum
	})
}

/**
 * Adds a Section Zoom object (MS-PPTX §2.9) — anchors to a section by title.
 */
export function addSectionZoomDefinition(target: PresSlide, opt: SectionZoomProps): void {
	pushZoomObject(target, 'section', opt, obj => {
		obj.zoomSectionTitle = opt.sectionTitle
	})
}

/**
 * Adds a Summary Zoom object (MS-PPTX §2.11) — anchors to a section by title, carries layout factors.
 */
export function addSummaryZoomDefinition(target: PresSlide, opt: SummaryZoomProps): void {
	pushZoomObject(target, 'summary', opt, obj => {
		const titles = [
			...(opt.sectionTitle ? [opt.sectionTitle] : []),
			...(opt.sectionTitles ?? []),
		].filter((title, idx, all) => title && all.indexOf(title) === idx)
		if (titles.length === 0) throw new Error('addSummaryZoom() error: provide `sectionTitle` or `sectionTitles`')
		obj.zoomSectionTitle = titles[0]
		obj.zoomSectionTitles = titles.slice(1)
		if (obj.options) {
			obj.options.zoomTitle = opt.title
			obj.options.zoomDescr = opt.descr
			obj.options.offsetFactorX = opt.offsetFactorX
			obj.options.offsetFactorY = opt.offsetFactorY
			obj.options.scaleFactorX = opt.scaleFactorX
			obj.options.scaleFactorY = opt.scaleFactorY
		}
	})
}

/**
 * Shared zoom-object builder (slide/section/summary). Registers the cover/thumbnail image rel and
 * pushes the zoom `ISlideObject`; `configure` applies the zoom-kind-specific fields.
 */
function pushZoomObject(target: PresSlide, kind: 'slide' | 'section' | 'summary', opt: ZoomProps | SectionZoomProps | SummaryZoomProps, configure: (obj: ISlideObject) => void): void {
	const objectName = opt.objectName ? encodeXmlEntities(opt.objectName) : `Zoom ${target._slideObjects.filter(obj => obj._type === SLIDE_OBJECT_TYPES.zoom).length + 1}`
	const strCover = opt.cover || IMG_PLAYBTN
	if (!strCover.toLowerCase().includes('base64,')) {
		throw new Error('addZoom() error: `cover` must be a base64 data URI (Ex: \'data:image/png;base64,iV[...]\')')
	}

	const slideData: ISlideObject = { _type: SLIDE_OBJECT_TYPES.zoom, zoomKind: kind }
	slideData.options = {
		x: opt.x || 0,
		y: opt.y || 0,
		w: opt.w || 2,
		h: opt.h || 1.13,
		altText: opt.altText || '',
		objectName,
		returnToParent: opt.returnToParent,
		showBg: opt.showBg,
		transitionDur: opt.transitionDur,
	} as ObjectOptions
	configure(slideData)

	// Cover/thumbnail image rel
	slideData.zoomRid = getNewRelId(target)
	target._relsMedia.push({
		path: 'preencoded.png',
		data: strCover,
		type: 'image/png',
		extn: 'png',
		rId: slideData.zoomRid,
		Target: `../media/${imageTargetStem(target)}.png`,
	})

	target._slideObjects.push(slideData)
}

function pushCoverImage (target: PresSlide, cover: string | undefined, method: string): number {
	const strCover = cover || IMG_PLAYBTN
	if (!strCover.toLowerCase().includes('base64,')) {
		throw new Error(`${method}() error: \`cover\` must be a base64 data URI (Ex: 'data:image/png;base64,iV[...]')`)
	}
	const rId = getNewRelId(target)
	target._relsMedia.push({
		path: 'preencoded.png',
		data: strCover,
		type: 'image/png',
		extn: 'png',
		rId,
		Target: `../media/${imageTargetStem(target)}.png`,
	})
	return rId
}

function nextExtPartName (target: PresSlide, prefix: string, extn: string): string {
	const n = target._rels.filter(rel => rel.type === SLIDE_OBJECT_TYPES.contentPart || rel.type === SLIDE_OBJECT_TYPES.officeApp).length + 1
	return `${prefix}-${target._slideNum ?? 0}-${n}.${extn}`
}

/**
 * Adds a content part (MS-PPTX §2.2.3). Choice is `p:contentPart`; Fallback is `sp`.
 */
export function addContentPartDefinition (target: PresSlide, opt: ContentPartProps, kind: 'content' | 'ink' = 'content'): void {
	if (typeof opt.data !== 'string' || !opt.data) {
		throw new Error(`${kind === 'ink' ? 'addInk' : 'addContentPart'}() error: \`data\` is required`)
	}

	const isInk = kind === 'ink'
	const count = target._slideObjects.filter(obj => obj._type === SLIDE_OBJECT_TYPES.contentPart && obj.contentPartKind === kind).length + 1
	const objectName = opt.objectName ? encodeXmlEntities(opt.objectName) : (isInk ? `Ink ${count}` : `Content Part ${count}`)
	const fileName = opt.fileName || (isInk ? nextExtPartName(target, 'ink', 'xml') : nextExtPartName(target, 'contentPart', 'xml'))
	const targetPath = isInk ? `../ink/${fileName}` : `../contentParts/${fileName}`

	const slideData: ISlideObject = {
		_type: SLIDE_OBJECT_TYPES.contentPart,
		contentPartKind: kind,
		contentPartBwMode: opt.bwMode,
		options: {
			x: opt.x || 0,
			y: opt.y || 0,
			w: opt.w || 2,
			h: opt.h || 1,
			altText: opt.altText || '',
			objectName,
		},
	}

	slideData.contentPartRid = getNewRelId(target)
	target._rels.push({
		type: SLIDE_OBJECT_TYPES.contentPart,
		data: opt.data,
		rId: slideData.contentPartRid,
		Target: targetPath,
		fileName,
		contentType: opt.contentType || (isInk ? CT_INKML : 'application/xml'),
	})

	if (isInk) slideData.coverRid = pushCoverImage(target, (opt as InkProps).cover, 'addInk')

	target._slideObjects.push(slideData)
}

/**
 * Adds ink as a content part (MS-PPTX §2.2.3.1). Fallback is `pic`.
 */
export function addInkDefinition (target: PresSlide, opt: InkProps): void {
	addContentPartDefinition(target, opt, 'ink')
}

/**
 * Adds an Office App reference (MS-PPTX §2.2.13 / MS-OWEXML webextensionref). Fallback is `pic`.
 */
export function addOfficeAppDefinition (target: PresSlide, opt: OfficeAppProps): void {
	if (!opt.reference?.id) throw new Error('addOfficeApp() error: `reference.id` is required')

	const count = target._slideObjects.filter(obj => obj._type === SLIDE_OBJECT_TYPES.officeApp).length + 1
	const objectName = opt.objectName ? encodeXmlEntities(opt.objectName) : `Office App ${count}`
	const fileName = nextExtPartName(target, 'webextension', 'xml')

	const slideData: ISlideObject = {
		_type: SLIDE_OBJECT_TYPES.officeApp,
		options: {
			x: opt.x || 0,
			y: opt.y || 0,
			w: opt.w || 3,
			h: opt.h || 2,
			altText: opt.altText || '',
			objectName,
		},
	}

	slideData.contentPartRid = getNewRelId(target)
	target._rels.push({
		type: SLIDE_OBJECT_TYPES.officeApp,
		data: makeXmlWebExtension(opt),
		rId: slideData.contentPartRid,
		Target: `../webextensions/${fileName}`,
		fileName,
		contentType: CT_WEBEXTENSION,
	})
	slideData.coverRid = pushCoverImage(target, opt.cover, 'addOfficeApp')

	target._slideObjects.push(slideData)
}

/**
 * Adds Notes to a slide.
 * @param {PresSlide} `target` slide object
 * @param {string} `notes`
 * @since 2.3.0
 */
export function addNotesDefinition(target: PresSlide, notes: string): void {
	target._slideObjects.push({
		_type: SLIDE_OBJECT_TYPES.notes,
		text: [{ text: notes }],
	})
}

/**
 * Adds a shape object to a slide definition.
 * @param {PresSlide} target slide object that the shape should be added to
 * @param {SHAPE_NAME} shapeName shape name
 * @param {ShapeProps} opts shape options
 */
/**
 * Map the deprecated v3.x line props (`line` as color string, `lineSize`/`lineDash`/`lineHead`/`lineTail`)
 * onto `line: ShapeLineProps`, warning once per prop. Shared by the shape and text paths; delete in the next major.
 * @param {object} opts - options object carrying the deprecated props
 * @param {ShapeLineProps} newLineOpts - the normalized line object to fill when `line` was a color string
 */
function normalizeDeprecatedLineProps(
	opts: { line?: ShapeLineProps | string, lineSize?: number, lineDash?: ShapeLineProps['dashType'], lineHead?: ShapeLineProps['beginArrowType'], lineTail?: ShapeLineProps['endArrowType'] },
	newLineOpts: ShapeLineProps
): void {
	if (typeof opts.line === 'string') {
		warnDeprecatedOnce('line-string', '`line: "<color>"` (string) is deprecated - use `line: { color: "..." }`')
		newLineOpts.color = opts.line // @deprecated [remove in next major]
		opts.line = newLineOpts
	}
	if (typeof opts.line !== 'object') return
	if (typeof opts.lineSize === 'number') { warnDeprecatedOnce('lineSize', '`lineSize` is deprecated - use `line.width`'); opts.line.width = opts.lineSize }
	if (typeof opts.lineDash === 'string') { warnDeprecatedOnce('lineDash', '`lineDash` is deprecated - use `line.dashType`'); opts.line.dashType = opts.lineDash }
	if (typeof opts.lineHead === 'string') { warnDeprecatedOnce('lineHead', '`lineHead` is deprecated - use `line.beginArrowType`'); opts.line.beginArrowType = opts.lineHead }
	if (typeof opts.lineTail === 'string') { warnDeprecatedOnce('lineTail', '`lineTail` is deprecated - use `line.endArrowType`'); opts.line.endArrowType = opts.lineTail }
}

export function addShapeDefinition(target: PresSlide | SlideLayout, shapeName: SHAPE_NAME, opts: ShapeProps): void {
	const options = typeof opts === 'object' ? opts : {}
	options.line = options.line || { type: 'none' }
	const newObject: ISlideObject = {
		_type: SLIDE_OBJECT_TYPES.text,
		shape: shapeName || SHAPE_TYPE.RECTANGLE,
		options,
		text: undefined,
	}

	// Reality check
	if (!shapeName) throw new Error('Missing/Invalid shape parameter! Example: `addShape(pptxgen.shapes.LINE, {x:1, y:1, w:1, h:1});`')

	// Validate preset geometry: an invalid preset string (e.g. 'oval' instead of the OOXML preset 'ellipse')
	// is written verbatim as prst="oval", which PowerPoint cannot parse - it silently drops the shape during
	// repair. TypeScript's `SHAPE_NAME` catches this at compile time; throw for untyped JS callers. (issue #1449)
	if (typeof shapeName === 'string' && !VALID_SHAPE_PRESETS.has(shapeName)) {
		throw new Error(`Invalid shape "${shapeName}" - use a \`pptxgen.ShapeType\` value (ex: \`pptxgen.ShapeType.oval\` = 'ellipse')`)
	}

	// 1: ShapeLineProps defaults
	const hasConnectorEnds = options.line.sourceId != null || options.line.targetId != null
	const newLineOpts: ShapeLineProps = {
		type: options.line.type || (options.line.gradient ? 'gradient' : 'solid'),
		color: options.line.color || DEF_SHAPE_LINE_COLOR,
		transparency: options.line.transparency || 0,
		gradient: options.line.gradient,
		width: options.line.width || 1,
		dashType: options.line.dashType || 'solid',
		cap: options.line.cap,
		beginArrowType: options.line.beginArrowType || undefined,
		endArrowType: options.line.endArrowType || undefined,
		// ZentoSoft connectors
		sourceId: options.line.sourceId,
		targetId: options.line.targetId,
		sourceAnchorPos: options.line.sourceAnchorPos,
		targetAnchorPos: options.line.targetAnchorPos,
		isConnector: options.line.isConnector || hasConnectorEnds,
		curveadjust: options.line.curveadjust,
	}
	if (typeof options.line === 'object' && options.line.type !== 'none') options.line = newLineOpts

	// 2: Set options defaults
	const shapeOpts = options as ObjectOptions
	shapeOpts._explicitX = options.x !== undefined && options.x !== null
	shapeOpts._explicitY = options.y !== undefined && options.y !== null
	shapeOpts._explicitW = options.w !== undefined && options.w !== null
	shapeOpts._explicitH = options.h !== undefined && options.h !== null
	options.x = options.x || (options.x === 0 ? 0 : 1)
	options.y = options.y || (options.y === 0 ? 0 : 1)
	options.w = options.w || (options.w === 0 ? 0 : 1)
	options.h = options.h || (options.h === 0 ? 0 : 1)
	options.objectName = options.objectName
		? encodeXmlEntities(options.objectName)
		: `Shape ${target._slideObjects.filter(obj => obj._type === SLIDE_OBJECT_TYPES.text).length}`

	// 3: Handle line (lots of deprecated opts)
	normalizeDeprecatedLineProps(options, newLineOpts)

	// 4: Create hyperlink rels
	createHyperlinkRels(target, newObject)

	// LAST: Add object to slide
	target._slideObjects.push(newObject)
}

/**
 * Adds a table object to a slide definition.
 * @param {PresSlide} target - slide object that the table should be added to
 * @param {TableRow[]} tableRows - table data
 * @param {TableProps} options - table options
 * @param {SlideLayout} slideLayout - Slide layout
 * @param {PresLayout} presLayout - Presentation layout
 * @param {Function} addSlide - method
 * @param {Function} getSlide - method
 */
export function addTableDefinition(
	target: PresSlide,
	tableRows: TableRow[],
	options: TableProps,
	slideLayout: SlideLayout,
	presLayout: PresLayout,
	addSlide: (options?: AddSlideProps) => PresSlide,
	getSlide: (slideNumber: number) => PresSlide
): PresSlide[] {
	const slides: PresSlide[] = [target] // Create array of Slides as more may be added by auto-paging
	const opt: TableProps = options && typeof options === 'object' ? { ...options } : {}
	opt.objectName = opt.objectName ? encodeXmlEntities(opt.objectName) : `Table ${target._slideObjects.filter(obj => obj._type === SLIDE_OBJECT_TYPES.table).length}`

	// STEP 1: REALITY-CHECK
	{
		// A: check for empty
		if (tableRows === null || tableRows.length === 0 || !Array.isArray(tableRows)) {
			throw new Error('addTable: Array expected! EX: \'slide.addTable( [rows], {options} );\' (https://gitbrent.github.io/PptxGenJS/docs/api-tables.html)')
		}

		// B: check for non-well-formatted array (ex: rows=['a','b'] instead of [['a','b']])
		if (!tableRows[0] || !Array.isArray(tableRows[0])) {
			throw new Error(
				'addTable: \'rows\' should be an array of cells! EX: \'slide.addTable( [ [\'A\'], [\'B\'], {text:\'C\',options:{align:\'center\'}} ] );\' (https://gitbrent.github.io/PptxGenJS/docs/api-tables.html)'
			)
		}

		// TODO: FUTURE: This is wacky and wont function right (shows .w value when there is none from demo.js?!) 20191219
		/*
		if (opt.w && opt.colW) {
			console.warn('addTable: please use either `colW` or `w` - not both (table will use `colW` and ignore `w`)')
			console.log(`${opt.w} ${opt.colW}`)
		}
		*/
	}

	// STEP 2: Transform `tableRows` into well-formatted TableCell's
	// tableRows can be object or plain text array: `[{text:'cell 1'}, {text:'cell 2', options:{color:'ff0000'}}]` | `["cell 1", "cell 2"]`
	const arrRows: TableCell[][] = []
	const tableFillRid = registerTableCellImageFill(target, opt.fill)
	tableRows.forEach(row => {
		const newRow: TableCell[] = []

		if (Array.isArray(row)) {
			row.forEach((cell: number | string | TableCell) => {
				// A:
				const newCell: TableCell = {
					_type: SLIDE_OBJECT_TYPES.tablecell,
					text: '',
					options: typeof cell === 'object' && cell.options ? cell.options : {},
				}

				// B:
				if (typeof cell === 'string' || typeof cell === 'number') newCell.text = cell.toString()
				else if (cell.text) {
					// Cell can contain complex text type, or string, or number
					if (typeof cell.text === 'string' || typeof cell.text === 'number') newCell.text = cell.text.toString()
					else if (cell.text) newCell.text = cell.text
					// Capture options
					if (cell.options && typeof cell.options === 'object') newCell.options = cell.options
				}

				// C: Set cell borders
				const cellOpts = newCell.options ?? {}
				newCell.options = cellOpts
				const initialBorder = cellOpts.border || opt.border || [{ type: 'none' }, { type: 'none' }, { type: 'none' }, { type: 'none' }]

				// CASE 1: border interface is: BorderOptions | [BorderOptions, BorderOptions, BorderOptions, BorderOptions]
				const cellBorder: BorderProps | [BorderProps, BorderProps, BorderProps, BorderProps] =
					!Array.isArray(initialBorder) && typeof initialBorder === 'object' ? [initialBorder, initialBorder, initialBorder, initialBorder] : initialBorder
				cellOpts.border = cellBorder
				// Handle: [null, null, {type:'solid'}, null]
				if (!cellBorder[0]) cellBorder[0] = { type: 'none' }
				if (!cellBorder[1]) cellBorder[1] = { type: 'none' }
				if (!cellBorder[2]) cellBorder[2] = { type: 'none' }
				if (!cellBorder[3]) cellBorder[3] = { type: 'none' }

				// set complete BorderOptions for all sides
				const arrSides = [0, 1, 2, 3]
				arrSides.forEach(idx => {
					cellBorder[idx] = {
						type: cellBorder[idx].type || DEF_CELL_BORDER.type,
						color: cellBorder[idx].color || DEF_CELL_BORDER.color,
						pt: cellBorder[idx].width ?? (typeof cellBorder[idx].pt === 'number' ? cellBorder[idx].pt : DEF_CELL_BORDER.pt),
						transparency: typeof cellBorder[idx].transparency === 'number' ? cellBorder[idx].transparency : undefined,
					}
				})

				const fillRid = registerTableCellImageFill(target, cellOpts.fill) ??
					(cellOpts.fill ? undefined : tableFillRid)
				if (fillRid) cellOpts._fillRid = fillRid

				// LAST:
				newRow.push(newCell)
			})
		} else {
			console.log('addTable: tableRows has a bad row. A row should be an array of cells. You provided:')
			console.log(row)
		}

		arrRows.push(newRow)
	})

	// STEP 3: Set options
	const tableOpts = opt as ObjectOptions
	tableOpts._explicitX = opt.x !== undefined && opt.x !== null
	tableOpts._explicitY = opt.y !== undefined && opt.y !== null
	tableOpts._explicitW = opt.w !== undefined && opt.w !== null
	tableOpts._explicitH = opt.h !== undefined && opt.h !== null
	opt.x = getSmartParseNumber(opt.x || (opt.x === 0 ? 0 : EMU / 2), 'X', presLayout)
	opt.y = getSmartParseNumber(opt.y || (opt.y === 0 ? 0 : EMU / 2), 'Y', presLayout)
	if (opt.h) opt.h = getSmartParseNumber(opt.h, 'Y', presLayout) // NOTE: Dont set default `h` - leaving it null triggers auto-rowH in `makeXMLSlide()`
	opt.fontSize = opt.fontSize || DEF_FONT_SIZE
	opt.margin = opt.margin === 0 || opt.margin ? opt.margin : DEF_CELL_MARGIN_IN
	if (typeof opt.margin === 'number') opt.margin = [Number(opt.margin), Number(opt.margin), Number(opt.margin), Number(opt.margin)]
	// NOTE: dont add default color on tables with hyperlinks! (it causes any textObj's with hyperlinks to have subsequent words to be black)
	if (JSON.stringify({ arrRows: arrRows }).indexOf('hyperlink') === -1) {
		if (!opt.color) opt.color = opt.color || DEF_FONT_COLOR // Set default color if needed (table option > inherit from Slide > default to black)
	}
	if (typeof opt.border === 'string') {
		console.warn('addTable `border` option must be an object. Ex: `{border: {type:\'none\'}}`')
		opt.border = undefined
	} else if (Array.isArray(opt.border)) {
		const optBorder = opt.border
		;[0, 1, 2, 3].forEach(idx => {
			optBorder[idx] = optBorder[idx]
				? { type: optBorder[idx].type || DEF_CELL_BORDER.type, color: optBorder[idx].color || DEF_CELL_BORDER.color, pt: optBorder[idx].width ?? optBorder[idx].pt ?? DEF_CELL_BORDER.pt, transparency: optBorder[idx].transparency }
				: { type: 'none' }
		})
	}

	opt.autoPage = typeof opt.autoPage === 'boolean' ? opt.autoPage : false
	opt.autoPageRepeatHeader = typeof opt.autoPageRepeatHeader === 'boolean' ? opt.autoPageRepeatHeader : false
	opt.autoPageHeaderRows = typeof opt.autoPageHeaderRows !== 'undefined' && !isNaN(Number(opt.autoPageHeaderRows)) ? Number(opt.autoPageHeaderRows) : 1
	opt.autoPageLineWeight = typeof opt.autoPageLineWeight !== 'undefined' && !isNaN(Number(opt.autoPageLineWeight)) ? Number(opt.autoPageLineWeight) : 0
	if (opt.autoPageLineWeight) {
		if (opt.autoPageLineWeight > 1) opt.autoPageLineWeight = 1
		else if (opt.autoPageLineWeight < -1) opt.autoPageLineWeight = -1
	}
	// autoPage ^^^

	// Set/Calc table width
	// Get slide margins - start with default values, then adjust if master or slide margins exist
	let arrTableMargin = DEF_SLIDE_MARGIN_IN
	// Case 1: Master margins
	if (slideLayout && typeof slideLayout._margin !== 'undefined') {
		if (Array.isArray(slideLayout._margin)) arrTableMargin = slideLayout._margin
		else if (!isNaN(Number(slideLayout._margin))) { arrTableMargin = [Number(slideLayout._margin), Number(slideLayout._margin), Number(slideLayout._margin), Number(slideLayout._margin)] }
	}
	// Case 2: Table margins
	/* FIXME: add `_margin` option to slide options
		else if ( addNewSlide._margin ) {
			if ( Array.isArray(addNewSlide._margin) ) arrTableMargin = addNewSlide._margin;
			else if ( !isNaN(Number(addNewSlide._margin)) ) arrTableMargin = [Number(addNewSlide._margin), Number(addNewSlide._margin), Number(addNewSlide._margin), Number(addNewSlide._margin)];
		}
	*/

	/**
	 * Calc table width depending upon what data we have - several scenarios exist (including bad data, eg: colW doesnt match col count)
	 * The API does not require a `w` value, but XML generation does, hence, code to calc a width below using colW value(s)
	 */
	if (opt.colW) {
		const firstRowColCnt = arrRows[0].reduce((totalLen, c) => {
			if (c?.options?.colspan && typeof c.options.colspan === 'number') {
				totalLen += c.options.colspan
			} else {
				totalLen += 1
			}
			return totalLen
		}, 0)

		if (typeof opt.colW === 'string' || typeof opt.colW === 'number') {
			// Ex: `colW = 3` or `colW = '3'`
			opt.w = Math.floor(Number(opt.colW) * firstRowColCnt)
			opt.colW = undefined // IMPORTANT: Unset `colW` so table is created using `opt.w`, which will evenly divide cols
		} else if (opt.colW && Array.isArray(opt.colW) && opt.colW.length === 1 && firstRowColCnt > 1) {
			// Ex: `colW=[3]` but with >1 cols (same as above, user is saying "use this width for all")
			opt.w = Math.floor(Number(opt.colW) * firstRowColCnt)
			opt.colW = undefined // IMPORTANT: Unset `colW` so table is created using `opt.w`, which will evenly divide cols
		} else if (opt.colW && Array.isArray(opt.colW) && opt.colW.length !== firstRowColCnt) {
			// Err: Mismatched colW and cols count
			console.warn('addTable: mismatch: (colW.length != data.length) Therefore, defaulting to evenly distributed col widths.')
			opt.colW = undefined
		}
	} else if (opt.w) {
		opt.w = getSmartParseNumber(opt.w, 'X', presLayout)
	} else {
		opt.w = Math.floor(Number(presLayout._sizeW) / EMU - arrTableMargin[1] - arrTableMargin[3])
	}

	// STEP 4: Convert units to EMU now (we use different logic in makeSlide->table - smartCalc is not used)
	if (opt.x && opt.x < 20) opt.x = inch2Emu(opt.x)
	if (opt.y && opt.y < 20) opt.y = inch2Emu(opt.y)
	if (opt.w && typeof opt.w === 'number' && opt.w < 20) opt.w = inch2Emu(opt.w)
	if (opt.h && typeof opt.h === 'number' && opt.h < 20) opt.h = inch2Emu(opt.h)

	// STEP 5: Loop over cells: transform each to ITableCell; check to see whether to unset `autoPage` while here
	arrRows.forEach(row => {
		row.forEach((cell, idy) => {
			// A: Transform cell data if needed
			/* Table rows can be an object or plain text - transform into object when needed
				// EX:
				var arrTabRows1 = [
					[ { text:'A1\nA2', options:{rowspan:2, fill:'99FFCC'} } ]
					,[ 'B2', 'C2', 'D2', 'E2' ]
				]
			*/
			if (typeof cell === 'number' || typeof cell === 'string') {
				// Grab table formatting `opts` to use here so text style/format inherits as it should
				row[idy] = { _type: SLIDE_OBJECT_TYPES.tablecell, text: String(row[idy]), options: opt }
			} else if (typeof cell === 'object') {
				// ARG0: `text`
				if (typeof cell.text === 'number') row[idy].text = String(row[idy].text)
				else if (typeof cell.text === 'undefined' || cell.text === null) row[idy].text = ''

				// ARG1: `options`: ensure options exists
				row[idy].options = cell.options || {}

				// Set type to tabelcell
				row[idy]._type = SLIDE_OBJECT_TYPES.tablecell
			}

			// B: Check for fine-grained formatting, disable auto-page when found
			// Since genXmlTextBody already checks for text array ( text:[{},..{}] ) we're done!
			// Text in individual cells will be formatted as they are added by calls to genXmlTextBody within table builder
			// if (cell.text && Array.isArray(cell.text)) opt.autoPage = false
			// TODO: FIXME: WIP: 20210807: We cant do this anymore
		})
	})

	// If autoPage = true, we need to return references to newly created slides if any
	const newAutoPagedSlides: PresSlide[] = []

	// STEP 6: Auto-Paging: (via {options} and used internally)
	// (used internally by `tableToSlides()` to not engage recursion - we've already paged the table data, just add this one)
	if (opt && !opt.autoPage) {
		// Create hyperlink rels (IMPORTANT: Wait until table has been shredded across Slides or all rels will end-up on Slide 1!)
		createHyperlinkRels(target, arrRows)

		// Add slideObjects (NOTE: Use `extend` to avoid mutation)
		target._slideObjects.push({
			_type: SLIDE_OBJECT_TYPES.table,
			arrTabRows: arrRows,
			options: Object.assign({}, opt),
		})
	} else {
		if (opt.autoPageRepeatHeader) opt._arrObjTabHeadRows = arrRows.filter((_row, idx) => idx < (opt.autoPageHeaderRows ?? 1))

		// Loop over rows and create 1-N tables as needed (ISSUE#21)
		getSlidesForTableRows(arrRows, opt, presLayout, slideLayout).forEach((slide, idx) => {
			// A: Create new Slide when needed, otherwise, use existing (NOTE: More than 1 table can be on a Slide, so we will go up AND down the Slide chain)
			if (!getSlide((target._slideNum ?? 0) + idx)) slides.push(addSlide({ masterName: slideLayout?._name || undefined }))

			// B: Reset opt.y to `option`/`margin` after first Slide (ISSUE#43, ISSUE#47, ISSUE#48)
			if (idx > 0) opt.y = inch2Emu(opt.autoPageSlideStartY || opt.newSlideStartY || arrTableMargin[0])

			// C: Add this table to new Slide
			{
				const newSlide: PresSlide = getSlide((target._slideNum ?? 0) + idx)

				opt.autoPage = false

				// Create hyperlink rels (IMPORTANT: Wait until table has been shredded across Slides or all rels will end-up on Slide 1!)
				createHyperlinkRels(newSlide, slide.rows)

				// Add rows to new slide
				newSlide.addTable(slide.rows, Object.assign({}, opt))

				// Add reference to the new slide so it can be returned, but don't add the first one because the user already has a reference to that one.
				if (idx > 0) newAutoPagedSlides.push(newSlide)
			}
		})
	}
	return newAutoPagedSlides
}

/**
 * Adds a text object to a slide definition.
 * @param {PresSlide} target - slide object that the text should be added to
 * @param {string|TextProps[]} text text string or object
 * @param {TextPropsOptions} opts text options
 * @param {boolean} isPlaceholder whether this a placeholder object
 * @since: 1.0.0
 */
export function addTextDefinition(target: PresSlide | SlideLayout, text: TextProps[], opts: TextPropsOptions, isPlaceholder: boolean): void {
	const newObject: ISlideObject = {
		_type: isPlaceholder ? SLIDE_OBJECT_TYPES.placeholder : SLIDE_OBJECT_TYPES.text,
		shape: (opts?.shape) || SHAPE_TYPE.RECTANGLE,
		text: !text || text.length === 0 ? [{ text: '' }] : text,
		options: opts || {},
	}

	function cleanOpts(itemOpts: ObjectOptions): TextPropsOptions {
		// STEP 1: Set some options
		{
			// A.1: Color (placeholders should inherit their colors or override them, so don't default them)
			if (!itemOpts.placeholder) {
				itemOpts.color = itemOpts.color || newObject.options?.color || ('color' in target ? target.color : undefined) || DEF_FONT_COLOR
			}

			// A.2: Placeholder should inherit their bullets or override them, so don't default them
			if (itemOpts.placeholder || isPlaceholder) {
				itemOpts.bullet = itemOpts.bullet || false
			}

			// A.3: Text targeting a placeholder inherits layout defaults (margin, valign, fill, …) (Issue #640).
			// Caller options win so a user fill/spacing is not wiped by the placeholder (fujita-h 11c4cb6).
			// Do not share the layout `_bodyProp` object — rebuild it below from the merged public options.
			itemOpts._explicitX = itemOpts.x !== undefined && itemOpts.x !== null
			itemOpts._explicitY = itemOpts.y !== undefined && itemOpts.y !== null
			itemOpts._explicitW = itemOpts.w !== undefined && itemOpts.w !== null
			itemOpts._explicitH = itemOpts.h !== undefined && itemOpts.h !== null
			if (itemOpts.placeholder && '_slideLayout' in target && target._slideLayout && target._slideLayout._slideObjects) {
				const placeHold = target._slideLayout._slideObjects.filter(
					item => item._type === 'placeholder' && item.options && item.options.placeholder && item.options.placeholder === itemOpts.placeholder
				)[0]
				if (placeHold?.options) {
					const { _bodyProp: _ignored, ...layoutOpts } = placeHold.options
					itemOpts = { ...layoutOpts, ...itemOpts, _bodyProp: {} }
				}
			}

			// A.4: Other options
			itemOpts.objectName = itemOpts.objectName
				? encodeXmlEntities(itemOpts.objectName)
				: `Text ${target._slideObjects.filter(obj => obj._type === SLIDE_OBJECT_TYPES.text).length}`

			// B:
			if (itemOpts.shape === SHAPE_TYPE.LINE) {
				// ShapeLineProps defaults
				const newLineOpts: ShapeLineProps = {
					type: itemOpts.line?.type || (itemOpts.line?.gradient ? 'gradient' : 'solid'),
					color: itemOpts.line?.color || DEF_SHAPE_LINE_COLOR,
					transparency: itemOpts.line?.transparency || 0,
					gradient: itemOpts.line?.gradient,
					width: itemOpts.line?.width || 1,
					dashType: itemOpts.line?.dashType || 'solid',
					cap: itemOpts.line?.cap,
					beginArrowType: itemOpts.line?.beginArrowType || undefined,
					endArrowType: itemOpts.line?.endArrowType || undefined,
				}
				if (typeof itemOpts.line === 'object') itemOpts.line = newLineOpts

				// 3: Handle line (lots of deprecated opts)
				normalizeDeprecatedLineProps(itemOpts, newLineOpts)
			}

			// C: Line opts
			itemOpts.line = itemOpts.line || {}
			itemOpts.lineSpacing = itemOpts.lineSpacing && !isNaN(itemOpts.lineSpacing) ? itemOpts.lineSpacing : undefined
			itemOpts.lineSpacingMultiple = itemOpts.lineSpacingMultiple && !isNaN(itemOpts.lineSpacingMultiple) ? itemOpts.lineSpacingMultiple : undefined

			// D: Transform text options to bodyProperties as thats how we build XML
			itemOpts._bodyProp = itemOpts._bodyProp || {}
			itemOpts._bodyProp.autoFit = itemOpts.autoFit || false // DEPRECATED: (3.3.0) If true, shape will collapse to text size (Fit To shape)
			itemOpts._bodyProp.anchor = !itemOpts.placeholder ? TEXT_VALIGN.ctr : undefined // VALS: [t,ctr,b]
			itemOpts._bodyProp.vert = itemOpts.vert || undefined // VALS: [eaVert,horz,mongolianVert,vert,vert270,wordArtVert,wordArtVertRtl]
			itemOpts._bodyProp.wrap = typeof itemOpts.wrap === 'boolean' ? itemOpts.wrap : true
			if (itemOpts.vertOverflow === 'overflow' || itemOpts.vertOverflow === 'ellipsis' || itemOpts.vertOverflow === 'clip') {
				itemOpts._bodyProp.vertOverflow = itemOpts.vertOverflow
			}
			if (itemOpts.horzOverflow === 'overflow' || itemOpts.horzOverflow === 'clip') {
				itemOpts._bodyProp.horzOverflow = itemOpts.horzOverflow
			}

			// Text columns: `columns` (count) + optional `columnGap` (inches) -> numCol/spcCol on <a:bodyPr>
			// ECMA-376 §5.1.5.1.4 CT_TextBodyProperties@numCol/@spcCol (issue #1320)
			if (itemOpts.columns && itemOpts.columns > 1) {
				itemOpts._bodyProp.numCol = Math.floor(itemOpts.columns)
				itemOpts._bodyProp.spcCol = inch2Emu(itemOpts.columnGap || 0)
			}

			// E: Inset
			// @deprecated 3.10.0 (`inset` - use `margin`)
			if ((itemOpts.inset && !isNaN(Number(itemOpts.inset))) || itemOpts.inset === 0) {
				itemOpts._bodyProp.lIns = inch2Emu(itemOpts.inset)
				itemOpts._bodyProp.rIns = inch2Emu(itemOpts.inset)
				itemOpts._bodyProp.tIns = inch2Emu(itemOpts.inset)
				itemOpts._bodyProp.bIns = inch2Emu(itemOpts.inset)
			}

			// F: Margin → ECMA-376 Part 1 §5.1.5.1.1 `CT_TextBodyProperties` insets
			// (`lIns`/`tIns`/`rIns`/`bIns`, `ST_Coordinate32` EMUs). Public API is TRBL
			// [top, right, bottom, left], matching table `marT`/`marR`/`marB`/`marL`.
			// Dual-unit: >=1 points, <1 inches (same rule as table cell margins).
			// mikemeerschaert/fix-inconsistent-margins + guiwoda/fix/text-margin-trbl
			if (itemOpts.margin && Array.isArray(itemOpts.margin)) {
				itemOpts._bodyProp.tIns = marginToEmu(itemOpts.margin[0] || 0)
				itemOpts._bodyProp.rIns = marginToEmu(itemOpts.margin[1] || 0)
				itemOpts._bodyProp.bIns = marginToEmu(itemOpts.margin[2] || 0)
				itemOpts._bodyProp.lIns = marginToEmu(itemOpts.margin[3] || 0)
			} else if (typeof itemOpts.margin === 'number') {
				const emu = marginToEmu(itemOpts.margin)
				itemOpts._bodyProp.tIns = emu
				itemOpts._bodyProp.rIns = emu
				itemOpts._bodyProp.bIns = emu
				itemOpts._bodyProp.lIns = emu
			}

			// G: Transform @deprecated props
			if (typeof itemOpts.underline === 'boolean' && itemOpts.underline === true) itemOpts.underline = { style: 'sng' }
		}

		// STEP 2: Transform `align`/`valign` to XML values, store in _bodyProp for XML gen
		{
			if ((itemOpts.align || '').toLowerCase().indexOf('c') === 0) itemOpts._bodyProp.align = TEXT_HALIGN.center
			else if ((itemOpts.align || '').toLowerCase().indexOf('l') === 0) itemOpts._bodyProp.align = TEXT_HALIGN.left
			else if ((itemOpts.align || '').toLowerCase().indexOf('r') === 0) itemOpts._bodyProp.align = TEXT_HALIGN.right
			else if ((itemOpts.align || '').toLowerCase().indexOf('j') === 0) itemOpts._bodyProp.align = TEXT_HALIGN.justify

			if ((itemOpts.valign || '').toLowerCase().indexOf('b') === 0) itemOpts._bodyProp.anchor = TEXT_VALIGN.b
			else if ((itemOpts.valign || '').toLowerCase().indexOf('m') === 0) itemOpts._bodyProp.anchor = TEXT_VALIGN.ctr
			else if ((itemOpts.valign || '').toLowerCase().indexOf('t') === 0) itemOpts._bodyProp.anchor = TEXT_VALIGN.t
		}

		// STEP 3: ROBUST: Set rational values for some shadow props if needed
		if (itemOpts.shadow) itemOpts.shadow = correctShadowOptions(itemOpts.shadow)

		return itemOpts
	}

	// STEP 1: Create/Clean object options
	newObject.options = cleanOpts(newObject.options ?? {})

	// STEP 2: Create/Clean text options
	newObject.text?.forEach(item => (item.options = cleanOpts(item.options || {})))

	// STEP 3: Create hyperlinks
	createHyperlinkRels(target, newObject.text || '')

	// LAST: Add object to Slide
	target._slideObjects.push(newObject)
}

/**
 * Adds placeholder objects to slide
 * @param {PresSlide} slide - slide object containing layouts
 */
export function addPlaceholdersToSlideLayouts(slide: PresSlide): void {
	// Add all placeholders on this Slide that dont already exist
	(slide._slideLayout._slideObjects || []).forEach(slideLayoutObj => {
		if (slideLayoutObj._type === SLIDE_OBJECT_TYPES.placeholder) {
			// A: Search for this placeholder on Slide before we add
			// NOTE: Check to ensure a placeholder does not already exist on the Slide
			// They are created when they have been populated with text (ex: `slide.addText('Hi', { placeholder:'title' });`)
			if (slide._slideObjects.filter(slideObj => slideObj.options && slideObj.options.placeholder === slideLayoutObj.options?.placeholder).length === 0) {
				// Must be a PLACEHOLDER (not TEXT), else PowerPoint shows a duplicate empty "Click to add text" box (issue #1453)
				addTextDefinition(slide, [{ text: '' }], slideLayoutObj.options ?? {}, true)
			}
		}
	})
}

/* -------------------------------------------------------------------------------- */

/**
 * Adds a background image or color to a slide definition.
 * @param {BackgroundProps} props - color string or an object with image definition
 * @param {PresSlide} target - slide object that the background is set to
 */
export function addBackgroundDefinition(props: BackgroundProps, target: SlideLayout): void {
	// A: @deprecated
	if (target.bkgd) {
		warnDeprecatedOnce('bkgd', '`bkgd` is deprecated - use `background`')
		if (!target.background) target.background = {}

		if (typeof target.bkgd === 'string') target.background.color = target.bkgd
		else {
			if (target.bkgd.data) target.background.data = target.bkgd.data
			if (target.bkgd.path) target.background.path = target.bkgd.path
			if (target.bkgd.src) target.background.path = target.bkgd.src // @deprecated (drop in 4.x)
		}
	}
	if (target.background?.fill) target.background.color = target.background.fill

	// B: Handle media
	if (props && (props.path || props.data)) {
		// Allow the use of only the data key (`path` isnt reqd)
		props.path = props.path || 'preencoded.png'
		let strImgExtn = (props.path.split('.').pop() || 'png').split('?')[0] // Handle "blah.jpg?width=540" etc.
		if (strImgExtn === 'jpg') strImgExtn = 'jpeg' // base64-encoded jpg's come out as "data:image/jpeg;base64,/9j/[...]", so correct exttnesion to avoid content warnings at PPT startup

		target._relsMedia = target._relsMedia || []
		const intRels = target._relsMedia.length + 1
		// NOTE: `Target` cannot have spaces (eg:"Slide 1-image-1.jpg") or a "presentation is corrupt" warning comes up
		target._relsMedia.push({
			path: props.path,
			type: SLIDE_OBJECT_TYPES.image,
			extn: strImgExtn,
			data: props.data || undefined,
			rId: intRels,
			Target: `../media/${(target._name || '').replace(/\s+/gi, '-')}-image-${target._relsMedia.length + 1}.${strImgExtn}`,
		})
		target._bkgdImgRid = intRels
	}
}

/**
 * Parses text/text-objects from `addText()` and `addTable()` methods; creates 'hyperlink'-type Slide Rels for each hyperlink found
 * @param {PresSlide} target - slide object that any hyperlinks will be be added to
 * @param {number | string | TextProps | TextProps[] | ITableCell[][]} text - text to parse
 */
function createHyperlinkRels(
	target: PresSlide | SlideLayout,
	text: number | string | ISlideObject | TextProps | TextProps[] | TableCell[] | TableCell[][],
	options?: TextPropsOptions[],
): void {
	let textObjs: Array<TextProps | TableCell | TableCell[]> = []

	// Only text objects can have hyperlinks, bail when text param is plain text
	if (typeof text === 'string' || typeof text === 'number') return
	// IMPORTANT: "else if" Array.isArray must come before typeof===object! Otherwise, code will exhaust recursion!
	else if (Array.isArray(text)) textObjs = text
	// A full ISlideObject (has `_type`) carries no inline hyperlink text; only bare text objects do
	else if (typeof text === 'object' && !('_type' in text)) textObjs = [text]

	textObjs.forEach((text, idx: number) => {
		// IMPORTANT: `options` are lost due to recursion/copy!
		if (!Array.isArray(text) && options && options[idx] && options[idx].hyperlink) text.options = { ...text.options, ...options[idx] }

		// NOTE: `text` can be an array of other `text` objects (table cell word-level formatting), continue parsing using recursion
		if (Array.isArray(text)) {
			const cellOpts: TextPropsOptions[] = []
			text.forEach((tablecell) => {
				// NB: `tablecell.text.options` never exists (text is string|TableCell[]), so this is just a presence check
				if (tablecell.options) {
					cellOpts.push(tablecell.options)
				}
			})
			createHyperlinkRels(target, text, cellOpts)
		} else if (Array.isArray(text.text)) {
			createHyperlinkRels(target, text.text, options && options[idx] ? [options[idx]] : undefined)
		} else if (text && typeof text === 'object' && text.options && text.options.hyperlink && !text.options.hyperlink._rId) {
			if (typeof text.options.hyperlink !== 'object') {
				console.log('ERROR: text `hyperlink` option should be an object. Ex: `hyperlink: {url:\'https://github.com\'}` ')
			}
			else if (!text.options.hyperlink.url && !text.options.hyperlink.slide) {
				console.log('ERROR: \'hyperlink requires either: `url` or `slide`\'')
			}
			else {
				const relId = getNewRelId(target)

				target._rels.push({
					type: SLIDE_OBJECT_TYPES.hyperlink,
					data: text.options.hyperlink.slide ? 'slide' : 'dummy',
					rId: relId,
					Target: encodeXmlEntities(text.options.hyperlink.url) || text.options.hyperlink.slide?.toString() || '',
				})

				text.options.hyperlink._rId = relId
			}
		}
		else if (text && typeof text === 'object' && text.options && text.options.hyperlink && text.options.hyperlink._rId) {
			// NOTE: auto-paging will create new slides, but skip above as _rId exists, BUT this is a new slide, so add rels!
			if (target._rels.filter(rel => rel.rId === text.options?.hyperlink?._rId).length === 0) {
				target._rels.push({
					type: SLIDE_OBJECT_TYPES.hyperlink,
					data: text.options.hyperlink.slide ? 'slide' : 'dummy',
					rId: text.options.hyperlink._rId,
					Target: encodeXmlEntities(text.options.hyperlink.url) || text.options.hyperlink.slide?.toString() || '',
				})
			}
		}
	})
}

function makeGroupCollector (parent: PresSlide): PresSlide {
	return {
		_slideObjects: [],
		_rels: parent._rels,
		_relsMedia: parent._relsMedia,
		_relsChart: parent._relsChart,
		_presLayout: parent._presLayout,
		_slideNum: parent._slideNum,
		_allocChartId: parent._allocChartId,
		_rId: '_rId' in parent ? parent._rId : 0,
	} as unknown as PresSlide
}

/**
 * Adds a PresentationML group (`p:grpSp`, ECMA-376 §19.3.1.22).
 * Child coordinates are relative to the group origin (`chOff` 0,0 / `chExt` = group size).
 */
export function addGroupDefinition (target: PresSlide, opts: GroupProps, objects: ISlideObject[]): ISlideObject {
	const groupCount = target._slideObjects.filter(obj => obj._type === SLIDE_OBJECT_TYPES.group).length
	const newObject: ISlideObject = {
		_type: SLIDE_OBJECT_TYPES.group,
		options: {
			x: opts.x ?? 0,
			y: opts.y ?? 0,
			w: opts.w ?? 1,
			h: opts.h ?? 1,
			rotate: opts.rotate ?? 0,
			shadow: opts.shadow ? correctShadowOptions(opts.shadow) : undefined,
			objectName: opts.objectName ? encodeXmlEntities(opts.objectName) : `Group ${groupCount + 1}`,
		},
		_objects: objects,
	}
	target._slideObjects.push(newObject)
	return newObject
}

export function createGroupBuilder (parent: PresSlide, objects: ISlideObject[]): Group {
	const builder: Group = {
		addShape (shapeName, options) {
			const collector = makeGroupCollector(parent)
			addShapeDefinition(collector, shapeName, options ?? {})
			objects.push(...collector._slideObjects)
			return builder
		},
		addText (text, options) {
			const collector = makeGroupCollector(parent)
			const textParam = typeof text === 'string' || typeof text === 'number' ? [{ text, options }] : text
			addTextDefinition(collector, textParam, options ?? {}, false)
			objects.push(...collector._slideObjects)
			return builder
		},
		addImage (options) {
			const collector = makeGroupCollector(parent)
			addImageDefinition(collector, options)
			objects.push(...collector._slideObjects)
			return builder
		},
		addGroup (options, build) {
			const children: ISlideObject[] = []
			if (build) build(createGroupBuilder(parent, children))
			const collector = makeGroupCollector(parent)
			addGroupDefinition(collector, options, children)
			objects.push(...collector._slideObjects)
			return builder
		},
	}
	return builder
}
