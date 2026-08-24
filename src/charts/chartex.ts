/**
 * ChartEx (`cx:chartSpace`) rendering - MS-ODRAWXML 2.1.
 *
 * PowerPoint 2016 introduced chart layouts that ECMA-376 has no markup for. They are written to their
 * own part, in the `cx` namespace, and share nothing with the `c:` emitter in `xml.ts` beyond the
 * embedded workbook: the data lives in `cx:chartData` (one `cx:data` per series) and the layout is
 * selected by `cx:series@layoutId` rather than by the element name.
 *
 * The cell references point at the same worksheet `workbook.ts` writes for classic charts - column A
 * holds the category labels, columns B.. hold one series each, row 1 holds the series names - so no
 * chartex-specific workbook layout is needed.
 */

import { CHARTEX_LAYOUT_ID, CHARTEX_NAME, DEF_FONT_COLOR, DEF_FONT_SIZE, DEF_FONT_TITLE_SIZE, OOXML_CHARTEX, isChartexType } from '../core-enums'
import { IChartOptsLib, IOptsChartData, ISlideRelChart } from '../core-interfaces'
import { createColorElement, encodeXmlEntities, getUuid } from '../gen-utils'
import { getExcelColName } from './utils'

/** Sheet name written by `workbook.ts`; chartex cell references are absolute and always single-sheet */
const SHEET = 'Sheet1'

/**
 * Category-axis gap width per layout, matching what PowerPoint writes for a freshly inserted chart.
 * A layout absent here (treemap, sunburst) draws no axes at all.
 */
const CAT_GAP_WIDTH: Partial<Record<CHARTEX_NAME, string>> = { boxWhisker: '1', funnel: '0.0599999987', histogram: '0', waterfall: '0.5' }

function firstLabelLevel (data: IOptsChartData[]): string[] {
	const labels = data[0]?.labels
	if (!labels || labels.length === 0) return []
	const level = labels[0]
	return Array.isArray(level) ? level : []
}

/**
 * `<cx:pt idx="N">value</cx:pt>` for every point of one dimension
 * - a missing point is written empty: `cx:pt` inside a `cx:numDim` must be a number, so the string
 *   "null" would be a schema violation (the ECMA-376 emitter blanks the same case)
 */
function makePoints (values: Array<string | number | null | undefined>): string {
	return values
		.map((value, idx) => `<cx:pt idx="${idx}">${value === null || value === undefined ? '' : encodeXmlEntities(String(value))}</cx:pt>`)
		.join('')
}

/**
 * Absolute A1 reference for one column of the embedded worksheet, excluding the header row.
 * The range length must equal the dimension's own `ptCount`, so each dimension passes its own.
 */
function columnRef (colIdx: number, pointCount: number): string {
	const col = getExcelColName(colIdx)
	return `${SHEET}!$${col}$2:$${col}$${pointCount + 1}`
}

/**
 * One `cx:data` block per series.
 * - hierarchical layouts (treemap, sunburst) size their segments through a `size` dimension
 * - a histogram bins raw values itself, so it carries no category dimension at all
 */
function makeChartData (rel: ISlideRelChart, chartType: CHARTEX_NAME): string {
	const labels = firstLabelLevel(rel.data)
	const dimType = chartType === 'treemap' || chartType === 'sunburst' ? 'size' : 'val'

	return rel.data
		.map((series, seriesIdx) => {
			const values = series.values ?? []
			const catDim =
				chartType === 'histogram' || labels.length === 0
					? ''
					: `<cx:strDim type="cat"><cx:f>${columnRef(1, labels.length)}</cx:f><cx:lvl ptCount="${labels.length}">${makePoints(labels)}</cx:lvl></cx:strDim>`
			const valDim =
				`<cx:numDim type="${dimType}"><cx:f>${columnRef(seriesIdx + 2, values.length)}</cx:f>` +
				`<cx:lvl ptCount="${values.length}" formatCode="General">${makePoints(values)}</cx:lvl></cx:numDim>`

			return `<cx:data id="${seriesIdx}">${catDim}${valDim}</cx:data>`
		})
		.join('')
}

/** `cx:dataLabels` - chartex has no `c:dLbls`, visibility is a single element with three flags */
function makeDataLabels (opts: IChartOptsLib, chartType: CHARTEX_NAME): string {
	const showValue = opts.showValue ?? false
	const showCategory = opts.showLabel ?? false
	if (!showValue && !showCategory) return ''

	// `outEnd` is invalid on the layouts that draw labels inside their segments
	const defaultPos = chartType === 'treemap' || chartType === 'sunburst' || chartType === 'funnel' ? 'ctr' : 'outEnd'
	const pos = opts.dataLabelPosition ?? defaultPos
	const size = Math.round((opts.dataLabelFontSize || DEF_FONT_SIZE) * 100)
	const txPr =
		'<cx:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>' +
		`<a:defRPr sz="${size}" b="${opts.dataLabelFontBold ? 1 : 0}" i="${opts.dataLabelFontItalic ? 1 : 0}">` +
		`<a:solidFill>${createColorElement(opts.dataLabelColor || DEF_FONT_COLOR)}</a:solidFill>` +
		`<a:latin typeface="${encodeXmlEntities(opts.dataLabelFontFace || 'Arial')}"/>` +
		'</a:defRPr></a:pPr></a:p></cx:txPr>'
	const numFmt = opts.dataLabelFormatCode ? `<cx:numFmt formatCode="${encodeXmlEntities(opts.dataLabelFormatCode)}" sourceLinked="0"/>` : ''

	return (
		`<cx:dataLabels pos="${pos}">` +
		`<cx:visibility seriesName="0" categoryName="${showCategory ? 1 : 0}" value="${showValue ? 1 : 0}"/>` +
		numFmt +
		txPr +
		'</cx:dataLabels>'
	)
}

/**
 * `cx:layoutPr` carries everything that is layout-specific rather than series-generic.
 * Anything absent here is left to PowerPoint's own default for that layout.
 */
function makeLayoutProps (opts: IChartOptsLib, chartType: CHARTEX_NAME): string {
	switch (chartType) {
		case 'waterfall': {
			const subtotals = opts.chartExSubtotals ?? []
			if (subtotals.length === 0) return ''
			return `<cx:layoutPr><cx:subtotals>${subtotals.map(idx => `<cx:idx val="${idx}"/>`).join('')}</cx:subtotals></cx:layoutPr>`
		}
		case 'histogram': {
			const binSize = opts.chartExBinSize
			const binCount = opts.chartExBinCount
			// PowerPoint accepts exactly one binning rule; an explicit size wins over a count, and neither
			// means "let PowerPoint choose" (its automatic Scott's-rule binning)
			const rule = binSize !== undefined ? `<cx:binSize val="${binSize}"/>` : binCount !== undefined ? `<cx:binCount val="${binCount}"/>` : ''
			const binning = rule ? `<cx:binning intervalClosed="r">${rule}</cx:binning>` : '<cx:binning intervalClosed="r"/>'
			return `<cx:layoutPr>${binning}</cx:layoutPr>`
		}
		case 'boxWhisker':
			return (
				'<cx:layoutPr>' +
				`<cx:visibility meanLine="${opts.chartExMeanLine ? 1 : 0}" meanMarker="1" nonoutliers="0" outliers="1"/>` +
				'<cx:statistics quartileMethod="exclusive"/>' +
				'</cx:layoutPr>'
			)
		case 'treemap':
			return `<cx:layoutPr><cx:parentLabelLayout val="${opts.chartExParentLabels ?? 'overlapping'}"/></cx:layoutPr>`
		default:
			return ''
	}
}

/** One `cx:series` per data row. Order follows CT_Series: tx, spPr, dataLabels, dataId, layoutPr, axisId. */
function makeSeries (rel: ISlideRelChart, chartType: CHARTEX_NAME): string {
	const layoutId = CHARTEX_LAYOUT_ID[chartType]
	const dataLabels = makeDataLabels(rel.opts, chartType)
	const layoutPr = makeLayoutProps(rel.opts, chartType)

	return rel.data
		.map((series, idx) => {
			const nameCell = `${SHEET}!$${getExcelColName(idx + 2)}$1`
			// These layouts color each data point from the chart color style, so `chartColors` does not
			// apply - a series-level fill would flatten a treemap to one color. Only `color` overrides it.
			const spPr = series.color && series.color !== 'transparent' ? `<cx:spPr><a:solidFill>${createColorElement(series.color)}</a:solidFill></cx:spPr>` : ''

			return (
				`<cx:series layoutId="${layoutId}" uniqueId="{${getUuid('xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx').toUpperCase()}}">` +
				`<cx:tx><cx:txData><cx:f>${nameCell}</cx:f><cx:v>${encodeXmlEntities(series.name ?? `Series ${idx + 1}`)}</cx:v></cx:txData></cx:tx>` +
				spPr +
				dataLabels +
				`<cx:dataId val="${idx}"/>` +
				layoutPr +
				'</cx:series>'
			)
		})
		.join('')
}

/**
 * `cx:axis` elements for the layouts that have axes. CT_Axis order: scaling, gridlines, tickLabels.
 * - a funnel is drawn along a single category axis, which PowerPoint numbers `1`
 * - treemap and sunburst have no axes; the series carries no `cx:axisId` in either case
 */
function makeAxes (rel: ISlideRelChart, chartType: CHARTEX_NAME): string {
	const gapWidth = CAT_GAP_WIDTH[chartType]
	if (gapWidth === undefined) return ''

	if (chartType === 'funnel') return `<cx:axis id="1"><cx:catScaling gapWidth="${gapWidth}"/><cx:tickLabels/></cx:axis>`

	const valGridlines = rel.opts.valGridLine?.style === 'none' ? '' : '<cx:majorGridlines/>'
	return (
		`<cx:axis id="0"><cx:catScaling gapWidth="${gapWidth}"/><cx:tickLabels/></cx:axis>` +
		`<cx:axis id="1"><cx:valScaling/>${valGridlines}<cx:tickLabels/></cx:axis>`
	)
}

function makeTitle (opts: IChartOptsLib): string {
	if (!opts.showTitle) return ''
	const size = Math.round((opts.titleFontSize || DEF_FONT_TITLE_SIZE) * 100)
	const color = opts.titleColor ? `<a:solidFill>${createColorElement(opts.titleColor)}</a:solidFill>` : ''
	const face = opts.titleFontFace ? `<a:latin typeface="${encodeXmlEntities(opts.titleFontFace)}"/>` : ''

	return (
		'<cx:title pos="t" align="ctr" overlay="0"><cx:tx><cx:rich><a:bodyPr/><a:lstStyle/><a:p><a:pPr algn="ctr"/><a:r>' +
		`<a:rPr lang="${opts.lang || 'en-US'}" sz="${size}" b="${opts.titleBold ? 1 : 0}">${color}${face}</a:rPr>` +
		`<a:t>${encodeXmlEntities(opts.title || 'Chart Title')}</a:t>` +
		'</a:r></a:p></cx:rich></cx:tx></cx:title>'
	)
}

function makeLegend (opts: IChartOptsLib): string {
	if (!opts.showLegend) return ''
	// `cx:legend@pos` takes l/r/t/b only; PowerPoint stores the library's `tr` as a top-aligned right legend
	const isTopRight = opts.legendPos === 'tr'
	return `<cx:legend pos="${isTopRight ? 'r' : opts.legendPos || 'r'}" align="${isTopRight ? 'min' : 'ctr'}" overlay="0"/>`
}

/**
 * Build the `cx:chartSpace` part for one chartex chart.
 * @param {ISlideRelChart} rel - chart relationship, with `opts._type` narrowed to a chartex type
 * @returns {string} chartex part XML
 */
export function makeXmlChartEx (rel: ISlideRelChart): string {
	const type = rel.opts._type
	if (!isChartexType(type)) throw new Error(`pptxgenjs: "${String(type)}" is not a chartex chart type`)

	return (
		'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
		'<cx:chartSpace xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main"' +
		' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
		` xmlns:cx="${OOXML_CHARTEX.ns}">` +
		'<cx:chartData>' +
		'<cx:externalData r:id="rId1" cx:autoUpdate="0"/>' +
		makeChartData(rel, type) +
		'</cx:chartData>' +
		'<cx:chart>' +
		makeTitle(rel.opts) +
		'<cx:plotArea>' +
		`<cx:plotAreaRegion>${makeSeries(rel, type)}</cx:plotAreaRegion>` +
		makeAxes(rel, type) +
		'</cx:plotArea>' +
		makeLegend(rel.opts) +
		'</cx:chart>' +
		'</cx:chartSpace>'
	)
}
