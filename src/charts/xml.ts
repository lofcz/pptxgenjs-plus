/**
 * Chart XML rendering.
 */

import {
	AXIS_ID_CATEGORY_PRIMARY,
	AXIS_ID_CATEGORY_SECONDARY,
	AXIS_ID_SERIES_PRIMARY,
	AXIS_ID_VALUE_PRIMARY,
	AXIS_ID_VALUE_SECONDARY,
	BARCHART_COLORS,
	CHART_NAME,
	CHART_TYPE,
	isChartexType,
	DEF_FONT_COLOR,
	DEF_FONT_SIZE,
	DEF_FONT_TITLE_SIZE,
} from '../core-enums'
import { IChartOptsLib, IOptsChartData, ISlideRelChart } from '../core-interfaces'
import { createColorElement, encodeXmlEntities, genXmlColorSelection, getUuid, valToPts } from '../gen-utils'
import { makeCatAxis, makeSerAxis, makeValAxis } from './axes'
import { genXmlTitle } from './title'
import {
	createLineCap,
	createShadowElement,
	genXmlErrBars,
	getErrorrateExcelCol,
	getExcelColName,
	resolveAreaGrouping,
	resolveSeriesLineDash,
	resolveShadowOptions,
	seriesHasErrorrate,
} from './utils'

/**
 * Main entry point method for create charts
 * @see: http://www.datypic.com/sc/ooxml/s-dml-chart.xsd.html
 * @param {ISlideRelChart} rel - chart object
 * @return {string} XML
 */
export function makeXmlCharts (rel: ISlideRelChart): string {
	let strXml = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
	let usesSecondaryValAxis = false

	// STEP 1: Create chart
	{
		// CHARTSPACE: BEGIN vvv
		strXml +=
			'<c:chartSpace xmlns:c="http://schemas.openxmlformats.org/drawingml/2006/chart" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
		strXml += '<c:date1904 val="0"/>' // ppt defaults to 1904 dates, excel to 1900
		strXml += `<c:roundedCorners val="${rel.opts.chartArea?.roundedCorners ? '1' : '0'}"/>`
		strXml += '<c:chart>'

		// OPTION: Title
		if (rel.opts.showTitle) {
			strXml += genXmlTitle(
				{
					title: rel.opts.title || 'Chart Title',
					color: rel.opts.titleColor,
					fontFace: rel.opts.titleFontFace,
					fontSize: rel.opts.titleFontSize || DEF_FONT_TITLE_SIZE,
					titleAlign: rel.opts.titleAlign,
					titleBold: rel.opts.titleBold,
					titleItalic: rel.opts.titleItalic,
					titlePos: rel.opts.titlePos,
					titleRotate: rel.opts.titleRotate,
				},
				typeof rel.opts.x === 'number' ? rel.opts.x : undefined,
				typeof rel.opts.y === 'number' ? rel.opts.y : undefined
			)
			strXml += '<c:autoTitleDeleted val="0"/>'
		} else {
			// NOTE: Add autoTitleDeleted tag in else to prevent default creation of chart title even when showTitle is set to false
			strXml += '<c:autoTitleDeleted val="1"/>'
		}
		/** Add 3D view tag
         * @see: https://c-rex.net/projects/samples/ooxml/e1/Part4/OOXML_P4_DOCX_perspective_topic_ID0E6BUQB.html
         */
		if (rel.opts._type === CHART_TYPE.BAR3D) {
			strXml += `<c:view3D><c:rotX val="${rel.opts.v3DRotX}"/><c:rotY val="${rel.opts.v3DRotY}"/><c:rAngAx val="${!rel.opts.v3DRAngAx ? 0 : 1}"/><c:perspective val="${rel.opts.v3DPerspective}"/></c:view3D>`
		}

		strXml += '<c:plotArea>'
		// IMPORTANT: Dont specify layout to enable auto-fit: PPT does a great job maximizing space with all 4 TRBL locations
		if (rel.opts.layout) {
			strXml += '<c:layout>'
			strXml += ' <c:manualLayout>'
			strXml += '  <c:layoutTarget val="inner" />'
			strXml += '  <c:xMode val="edge" />'
			strXml += '  <c:yMode val="edge" />'
			strXml += '  <c:x val="' + (rel.opts.layout.x || 0) + '" />'
			strXml += '  <c:y val="' + (rel.opts.layout.y || 0) + '" />'
			strXml += '  <c:w val="' + (rel.opts.layout.w || 1) + '" />'
			strXml += '  <c:h val="' + (rel.opts.layout.h || 1) + '" />'
			strXml += ' </c:manualLayout>'
			strXml += '</c:layout>'
		} else {
			strXml += '<c:layout/>'
		}
	}

	// A: Create Chart XML -----------------------------------------------------------
	if (Array.isArray(rel.opts._type)) {
		rel.opts._type.forEach(type => {
			// TODO: FIXME: theres `options` on chart rels??
			const options = { ...rel.opts, ...type.options }
			// let options: IChartOptsLib = { type: type.type, }
			const valAxisId = options.secondaryValAxis ? AXIS_ID_VALUE_SECONDARY : AXIS_ID_VALUE_PRIMARY
			const catAxisId = options.secondaryCatAxis ? AXIS_ID_CATEGORY_SECONDARY : AXIS_ID_CATEGORY_PRIMARY
			usesSecondaryValAxis = usesSecondaryValAxis || (options.secondaryValAxis ?? false)
			strXml += makeChartType(type.type, type.data, options, valAxisId, catAxisId, true, rel.data)
		})
	} else if (rel.opts._type && !isChartexType(rel.opts._type)) {
		strXml += makeChartType(rel.opts._type, rel.data, rel.opts, AXIS_ID_VALUE_PRIMARY, AXIS_ID_CATEGORY_PRIMARY, false, rel.data)
	}

	// B: Axes -----------------------------------------------------------
	const axisOptions = (axis: 'secondaryCatAxis' | 'secondaryValAxis', secondary: boolean): IChartOptsLib => {
		if (!Array.isArray(rel.opts._type)) return rel.opts
		const chart = rel.opts._type.find(type => Boolean(type.options?.[axis]) === secondary)
		return chart ? { ...rel.opts, ...chart.options, _type: chart.type } : rel.opts
	}

	if (rel.opts._type !== CHART_TYPE.PIE && rel.opts._type !== CHART_TYPE.DOUGHNUT) {
		// Param check
		if (rel.opts.valAxes && rel.opts.valAxes.length > 1 && !usesSecondaryValAxis) {
			throw new Error('Secondary axis must be used by one of the multiple charts')
		}

		if (rel.opts.catAxes) {
			if (!rel.opts.valAxes || rel.opts.valAxes.length !== rel.opts.catAxes.length) {
				throw new Error('There must be the same number of value and category axes.')
			}
			strXml += makeCatAxis({ ...axisOptions('secondaryCatAxis', false), ...rel.opts.catAxes[0] }, AXIS_ID_CATEGORY_PRIMARY, AXIS_ID_VALUE_PRIMARY)
		} else {
			strXml += makeCatAxis(axisOptions('secondaryCatAxis', false), AXIS_ID_CATEGORY_PRIMARY, AXIS_ID_VALUE_PRIMARY)
		}

		if (rel.opts.valAxes) {
			strXml += makeValAxis({ ...axisOptions('secondaryValAxis', false), ...rel.opts.valAxes[0] }, AXIS_ID_VALUE_PRIMARY)
			if (rel.opts.valAxes[1]) {
				strXml += makeValAxis({ ...axisOptions('secondaryValAxis', true), ...rel.opts.valAxes[1] }, AXIS_ID_VALUE_SECONDARY)
			}
		} else {
			strXml += makeValAxis(axisOptions('secondaryValAxis', false), AXIS_ID_VALUE_PRIMARY)

			// Add series axis for 3D bar
			if (rel.opts._type === CHART_TYPE.BAR3D) {
				strXml += makeSerAxis(rel.opts, AXIS_ID_SERIES_PRIMARY, AXIS_ID_VALUE_PRIMARY)
			}
		}

		// Combo Charts: Add secondary axes after all vals
		if (rel.opts?.catAxes && rel.opts?.catAxes[1]) {
			strXml += makeCatAxis({ ...axisOptions('secondaryCatAxis', true), ...rel.opts.catAxes[1] }, AXIS_ID_CATEGORY_SECONDARY, AXIS_ID_VALUE_SECONDARY)
		}
	}

	// C: Chart Properties and plotArea Options: Border, Data Table, Fill, Legend
	{
		// NOTE: DataTable goes between '</c:valAx>' and '<c:spPr>'
		if (rel.opts.showDataTable) {
			strXml += '<c:dTable>'
			strXml += `  <c:showHorzBorder val="${!rel.opts.showDataTableHorzBorder ? 0 : 1}"/>`
			strXml += `  <c:showVertBorder val="${!rel.opts.showDataTableVertBorder ? 0 : 1}"/>`
			strXml += `  <c:showOutline    val="${!rel.opts.showDataTableOutline ? 0 : 1}"/>`
			strXml += `  <c:showKeys       val="${!rel.opts.showDataTableKeys ? 0 : 1}"/>`
			strXml += '  <c:spPr>'
			strXml += '    <a:noFill/>'
			strXml += '    <a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="tx1"><a:lumMod val="15000"/><a:lumOff val="85000"/></a:schemeClr></a:solidFill><a:round/></a:ln>'
			strXml += '    <a:effectLst/>'
			strXml += '  </c:spPr>'
			strXml += '  <c:txPr>'
			strXml += '   <a:bodyPr rot="0" spcFirstLastPara="1" vertOverflow="ellipsis" vert="horz" wrap="square" anchor="ctr" anchorCtr="1"/>'
			strXml += '   <a:lstStyle/>'
			strXml += '   <a:p>'
			strXml += '     <a:pPr rtl="0">'
			strXml += `       <a:defRPr sz="${Math.round((rel.opts.dataTableFontSize || DEF_FONT_SIZE) * 100)}" b="0" i="0" u="none" strike="noStrike" kern="1200" baseline="0">`
			strXml += '         <a:solidFill><a:schemeClr val="tx1"><a:lumMod val="65000"/><a:lumOff val="35000"/></a:schemeClr></a:solidFill>'
			strXml += '         <a:latin typeface="+mn-lt"/>'
			strXml += '         <a:ea typeface="+mn-ea"/>'
			strXml += '         <a:cs typeface="+mn-cs"/>'
			strXml += '       </a:defRPr>'
			strXml += '     </a:pPr>'
			strXml += '    <a:endParaRPr lang="en-US"/>'
			strXml += '   </a:p>'
			strXml += ' </c:txPr>'
			strXml += '</c:dTable>'
		}

		strXml += '  <c:spPr>'

		// OPTION: Fill
		strXml += rel.opts.plotArea?.fill && (rel.opts.plotArea.fill.type === 'gradient' || rel.opts.plotArea.fill.type === 'linearGradient' || rel.opts.plotArea.fill.type === 'pattern' || rel.opts.plotArea.fill.color)
			? genXmlColorSelection(rel.opts.plotArea.fill)
			: '<a:noFill/>'

		// OPTION: Border
		strXml += rel.opts.plotArea?.border
			? `<a:ln w="${valToPts(rel.opts.plotArea?.border?.pt)}" cap="flat">${genXmlColorSelection(rel.opts.plotArea?.border?.color)}</a:ln>`
			: '<a:ln><a:noFill/></a:ln>'

		// Close shapeProp/plotArea before Legend
		strXml += '    <a:effectLst/>'
		strXml += '  </c:spPr>'
		strXml += '</c:plotArea>'

		// OPTION: Legend
		// IMPORTANT: Dont specify layout to enable auto-fit: PPT does a great job maximizing space with all 4 TRBL locations
		if (rel.opts.showLegend) {
			strXml += '<c:legend>'
			strXml += '<c:legendPos val="' + rel.opts.legendPos + '"/>'
			// strXml += '<c:layout/>'
			strXml += '<c:overlay val="0"/>'
			if (rel.opts.legendFontFace || rel.opts.legendFontSize || rel.opts.legendColor) {
				strXml += '<c:txPr>'
				strXml += '  <a:bodyPr/>'
				strXml += '  <a:lstStyle/>'
				strXml += '  <a:p>'
				strXml += '    <a:pPr>'
				strXml += rel.opts.legendFontSize ? `<a:defRPr sz="${Math.round(Number(rel.opts.legendFontSize) * 100)}">` : '<a:defRPr>'
				if (rel.opts.legendColor) strXml += genXmlColorSelection(rel.opts.legendColor)
				if (rel.opts.legendFontFace) strXml += '<a:latin typeface="' + encodeXmlEntities(rel.opts.legendFontFace) + '"/>'
				if (rel.opts.legendFontFace) strXml += '<a:ea    typeface="' + encodeXmlEntities(rel.opts.legendFontFace) + '"/>'
				if (rel.opts.legendFontFace) strXml += '<a:cs    typeface="' + encodeXmlEntities(rel.opts.legendFontFace) + '"/>'
				strXml += '      </a:defRPr>'
				strXml += '    </a:pPr>'
				strXml += '    <a:endParaRPr lang="en-US"/>'
				strXml += '  </a:p>'
				strXml += '</c:txPr>'
			}
			strXml += '</c:legend>'
		}
	}

	strXml += '  <c:plotVisOnly val="1"/>'
	strXml += '  <c:dispBlanksAs val="' + rel.opts.displayBlanksAs + '"/>'
	if (rel.opts._type === CHART_TYPE.SCATTER) strXml += '<c:showDLblsOverMax val="1"/>'

	strXml += '</c:chart>'

	// D: CHARTSPACE SHAPE PROPS
	strXml += '<c:spPr>'
	strXml += rel.opts.chartArea?.fill && (rel.opts.chartArea.fill.type === 'gradient' || rel.opts.chartArea.fill.type === 'linearGradient' || rel.opts.chartArea.fill.type === 'pattern' || rel.opts.chartArea.fill.color)
		? genXmlColorSelection(rel.opts.chartArea.fill)
		: '<a:noFill/>'
	strXml += rel.opts.chartArea?.border
		? `<a:ln w="${valToPts(rel.opts.chartArea?.border?.pt)}" cap="flat">${genXmlColorSelection(rel.opts.chartArea?.border?.color)}</a:ln>`
		: '<a:ln><a:noFill/></a:ln>'
	strXml += '  <a:effectLst/>'
	strXml += '</c:spPr>'

	// E: DATA (Add relID)
	strXml += '<c:externalData r:id="rId1"><c:autoUpdate val="0"/></c:externalData>'

	// LAST: chartSpace end
	strXml += '</c:chartSpace>'

	return strXml
}

/**
 * Create XML string for any given chart type
 * @param {CHART_NAME} chartType chart type name
 * @param {IOptsChartData[]} data chart data
 * @param {IChartOptsLib} opts chart options
 * @param {string} valAxisId chart val axis id
 * @param {string} catAxisId chart cat axis id
 * @param {boolean} _isMultiTypeChart is this a multi-type chart? (reserved; not currently used)
 * @example 'bubble' returns <c:bubbleChart></c>
 * @example '<c:lineChart>'
 * @return {string} XML chart
 */
/**
 * Default run properties for scatter `c:dLbls` / `c:txPr` (`a:defRPr`).
 * ECMA-376 Part 1 §5.7.2.49 `CT_DLbls` + §5.1.5.3.5 `CT_TextCharacterProperties`
 * (`standards/ecma/part-23_drawingml-reference-material-drawingml-charts.txt`).
 * An empty `<a:defRPr/>` drops dataLabelFont* on scatter labels (gitbrent #1348).
 */
function genXmlDataLabelDefRPr (opts: IChartOptsLib): string {
	return (
		`<a:defRPr sz="${Math.round((opts.dataLabelFontSize || DEF_FONT_SIZE) * 100)}" b="${opts.dataLabelFontBold ? 1 : 0}" i="${opts.dataLabelFontItalic ? 1 : 0}" u="none" strike="noStrike">` +
		`<a:solidFill>${createColorElement(opts.dataLabelColor || DEF_FONT_COLOR)}</a:solidFill>` +
		`<a:latin typeface="${encodeXmlEntities(opts.dataLabelFontFace) || 'Arial'}"/>` +
		'</a:defRPr>'
	)
}

/**
 * Rich-text body for a custom chart data label (`c:tx` / `c:rich`).
 * Escapes text and applies data-label font options.
 */
function genXmlDataLabelRichText (text: string, opts: IChartOptsLib): string {
	const sz = Math.round((opts.dataLabelFontSize || DEF_FONT_SIZE) * 100)
	const lang = opts.lang || 'en-US'
	let xml = '<c:tx><c:rich><a:bodyPr/><a:lstStyle/><a:p><a:r>'
	xml += `<a:rPr lang="${lang}" dirty="0" sz="${sz}" b="${opts.dataLabelFontBold ? 1 : 0}" i="${opts.dataLabelFontItalic ? 1 : 0}">`
	xml += `<a:solidFill>${createColorElement(opts.dataLabelColor || DEF_FONT_COLOR)}</a:solidFill>`
	xml += `<a:latin typeface="${encodeXmlEntities(opts.dataLabelFontFace) || 'Arial'}"/>`
	xml += '</a:rPr>'
	xml += `<a:t>${encodeXmlEntities(text)}</a:t>`
	xml += '</a:r></a:p></c:rich></c:tx>'
	return xml
}

/**
 * Per-point custom data label (`c:dLbl`) — used when series `dataLabels[i]` replaces the numeric label.
 * showVal is off so the custom text is not concatenated with the calculated value.
 */
function genXmlCustomDataLabel (text: string, idx: number, opts: IChartOptsLib): string {
	let xml = '<c:dLbl>'
	xml += `<c:idx val="${idx}"/>`
	xml += genXmlDataLabelRichText(text, opts)
	if (opts.dataLabelPosition) xml += `<c:dLblPos val="${opts.dataLabelPosition}"/>`
	xml += '<c:showLegendKey val="0"/><c:showVal val="0"/><c:showCatName val="0"/><c:showSerName val="0"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>'
	xml += '</c:dLbl>'
	return xml
}

/**
 * Sparse per-point `c:dLbl` overrides. Undefined holes keep series-level `c:dLbls` defaults.
 */
function genXmlCustomDataLabels (dataLabels: Array<string | undefined> | undefined, opts: IChartOptsLib): string {
	if (!dataLabels) return ''
	let xml = ''
	dataLabels.forEach((value, i) => {
		if (typeof value === 'string') xml += genXmlCustomDataLabel(value, i, opts)
	})
	return xml
}

/**
 * `c15:showDataLabelsRange` inside series `<c:dLbls>` (Toukyh/fix-custom-label).
 * Only emit when Value-From-Cells labels are in play.
 */
function genXmlShowDataLabelsRangeExt (enabled: boolean): string {
	return (
		'<c:extLst>' +
		'<c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">' +
		`<c15:showDataLabelsRange val="${enabled ? '1' : '0'}"/>` +
		'<c15:showLeaderLines val="0"/>' +
		'</c:ext>' +
		'</c:extLst>'
	)
}

/**
 * Series-level Value-From-Cells extension (`c15:datalabelsRange` + `c16:uniqueId`).
 * Fixes vs upstream Toukyh:
 * - XML-encode label text
 * - unique `c16:uniqueId` per series (upstream hard-coded one GUID for all)
 * - placed after cat/val (CT_*Ser schema), not mid-series after dLbls
 */
function genXmlDataLabelsRangeSerExt (labelsRange: string[], formula: string): string {
	const pts = labelsRange
		.map((value, idx) => `<c:pt idx="${idx}"><c:v>${encodeXmlEntities(String(value ?? ''))}</c:v></c:pt>`)
		.join('')
	const uniqueId = `{${getUuid('xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx').toUpperCase()}}`
	return (
		'<c:extLst>' +
		'<c:ext uri="{02D57815-91ED-43cb-92C2-25804820EDAC}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">' +
		'<c15:datalabelsRange>' +
		`<c15:f>${formula}</c15:f>` +
		'<c15:dlblRangeCache>' +
		`<c:ptCount val="${labelsRange.length}"/>` +
		pts +
		'</c15:dlblRangeCache>' +
		'</c15:datalabelsRange>' +
		'</c:ext>' +
		'<c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">' +
		`<c16:uniqueId val="${uniqueId}"/>` +
		'</c:ext>' +
		'</c:extLst>'
	)
}

function makeChartType (
	chartType: CHART_NAME,
	data: IOptsChartData[],
	opts: IChartOptsLib,
	valAxisId: string,
	catAxisId: string,
	_isMultiTypeChart: boolean,
	/** Full workbook series list (may be wider than `data` for multi-type charts) */
	sheetData: IOptsChartData[],
): string {
	// NOTE: "Chart Range" (as shown in "select Chart Area dialog") is calculated.
	// ....: Ensure each X/Y Axis/Col has same row height (esp. applicable to XY Scatter where X can often be larger than Y's)
	let colorIndex = -1 // Maintain the color index by region
	let idxColLtr = 1
	let strXml = ''
	// NOTE: chart data arrays and colors are always populated at render time; guarded locals keep this null-safe without changing output
	const firstValues = data[0]?.values ?? []
	const chartColors = opts.chartColors ?? []

	switch (chartType) {
		case CHART_TYPE.AREA:
		case CHART_TYPE.BAR:
		case CHART_TYPE.BAR3D:
		case CHART_TYPE.LINE:
		case CHART_TYPE.RADAR:
			// 1: Start Chart
			strXml += `<c:${chartType}Chart>`
			// Area charts: `c:grouping` / ST_Grouping (ECMA-376 §5.7.2.76 / §5.7.3.17).
			// `barGrouping` is the public option name; the element is not `c:barGrouping` (that is ST_BarGrouping on bar charts).
			if (chartType === CHART_TYPE.AREA) {
				strXml += '<c:grouping val="' + resolveAreaGrouping(opts.barGrouping) + '"/>'
			}

			if (chartType === CHART_TYPE.BAR || chartType === CHART_TYPE.BAR3D) {
				strXml += '<c:barDir val="' + opts.barDir + '"/>'
				strXml += '<c:grouping val="' + (opts.barGrouping || 'clustered') + '"/>'
			}

			if (chartType === CHART_TYPE.RADAR) {
				strXml += '<c:radarStyle val="' + opts.radarStyle + '"/>'
			}

			strXml += '<c:varyColors val="0"/>'

			// 2: "Series" block for every data row
			/* EX1:
				data: [
				 {
				   name: 'Region 1',
				   labels: [['April', 'May', 'June', 'July']],
				   values: [17, 26, 53, 96]
				 },
				 {
				   name: 'Region 2',
				   labels: [['April', 'May', 'June', 'July']],
				   values: [55, 43, 70, 58]
				 }
				]
            */
			/* EX2:
				data: [
				 {
				   name: 'Region 1',
				   labels: [
					   ['April', 'May', 'June', 'April', 'May', 'June'],
					   ['2020',     '',     '', '2021',     '',     '']
				   ],
				   values: [17, 26, 53, 96, 40, 33]
				 },
				 {
				   name: 'Region 2',
				   labels: [
					   ['April', 'May', 'June', 'April', 'May', 'June'],
					   ['2020',     '',     '', '2021',     '',     '']
				   ],
				   values: [55, 43, 70, 58, 78, 63]
				 }
				]
             */
			data.forEach(obj => {
				colorIndex++
				const dataIndex = obj._dataIndex ?? 0
				const objLabels = obj.labels ?? []
				const objValues = obj.values ?? []
				strXml += '<c:ser>'
				strXml += `  <c:idx val="${dataIndex}"/><c:order val="${dataIndex}"/>`
				strXml += '  <c:tx>'
				strXml += '    <c:strRef>'
				strXml += '      <c:f>Sheet1!$' + getExcelColName(dataIndex + objLabels.length + 1) + '$1</c:f>'
				strXml += '      <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>' + encodeXmlEntities(obj.name) + '</c:v></c:pt></c:strCache>'
				strXml += '    </c:strRef>'
				strXml += '  </c:tx>'

				// Fill and Border
				// NOTE: a per-series `color` wins over the shared `chartColors` cycle (issue #37)
				const seriesColor = obj.color ?? (opts.chartColors ? opts.chartColors[colorIndex % opts.chartColors.length] : undefined)

				strXml += '  <c:spPr>'
				if (seriesColor === 'transparent') {
					strXml += '<a:noFill/>'
				} else if (opts.chartColorsOpacity) {
					strXml += '<a:solidFill>' + createColorElement(seriesColor, `<a:alpha val="${Math.round(opts.chartColorsOpacity * 1000)}"/>`) + '</a:solidFill>'
				} else {
					strXml += '<a:solidFill>' + createColorElement(seriesColor) + '</a:solidFill>'
				}

				if (chartType === CHART_TYPE.LINE || chartType === CHART_TYPE.RADAR) {
					if (opts.lineSize === 0) {
						strXml += '<a:ln><a:noFill/></a:ln>'
					} else {
						strXml += `<a:ln w="${valToPts(opts.lineSize)}" cap="${createLineCap(opts.lineCap)}"><a:solidFill>${createColorElement(seriesColor)}</a:solidFill>`
						strXml += '<a:prstDash val="' + resolveSeriesLineDash(obj.lineDash, opts.lineDash) + '"/><a:round/></a:ln>'
					}
				} else if (opts.dataBorder) {
					strXml += `<a:ln w="${valToPts(opts.dataBorder.pt)}" cap="${createLineCap(opts.lineCap)}"><a:solidFill>${createColorElement(opts.dataBorder.color)}</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>`
				}

				strXml += createShadowElement(resolveShadowOptions(opts.shadow))

				strXml += '  </c:spPr>'
				strXml += '  <c:invertIfNegative val="0"/>'

				// Data Labels per series
				// NOTE: [20190117] Adding these to RADAR chart causes unrecoverable corruption!
				const labelsRange = obj.labelsRange?.length ? obj.labelsRange : undefined
				const useDataLabelsRange = Boolean(labelsRange) || Boolean(opts.showDataLabelsRange)
				if (chartType !== CHART_TYPE.RADAR) {
					strXml += '<c:dLbls>'
					// Per-point custom labels (series.dataLabels) — sparse arrays OK; other points keep series defaults
					strXml += genXmlCustomDataLabels(obj.dataLabels, opts)
					strXml += `<c:numFmt formatCode="${encodeXmlEntities(opts.dataLabelFormatCode) || 'General'}" sourceLinked="0"/>`
					if (opts.dataLabelBkgrdColors) strXml += `<c:spPr><a:solidFill>${createColorElement(seriesColor)}</a:solidFill></c:spPr>`
					strXml += '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>'
					strXml += `<a:defRPr b="${opts.dataLabelFontBold ? 1 : 0}" i="${opts.dataLabelFontItalic ? 1 : 0}" strike="noStrike" sz="${Math.round(
						(opts.dataLabelFontSize || DEF_FONT_SIZE) * 100
					)}" u="none">`
					strXml += `<a:solidFill>${createColorElement(opts.dataLabelColor || DEF_FONT_COLOR)}</a:solidFill>`
					strXml += `<a:latin typeface="${encodeXmlEntities(opts.dataLabelFontFace) || 'Arial'}"/>`
					strXml += '</a:defRPr></a:pPr></a:p></c:txPr>'
					if (opts.dataLabelPosition) strXml += `<c:dLblPos val="${opts.dataLabelPosition}"/>`
					strXml += '<c:showLegendKey val="0"/>'
					strXml += `<c:showVal val="${opts.showValue ? '1' : '0'}"/>`
					strXml += `<c:showCatName val="0"/><c:showSerName val="${opts.showSerName ? '1' : '0'}"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>`
					strXml += `<c:showLeaderLines val="${opts.showLeaderLines ? '1' : '0'}"/>`
					// Toukyh/fix-custom-label: Value From Cells toggle (only when in use)
					if (useDataLabelsRange) strXml += genXmlShowDataLabelsRangeExt(Boolean(labelsRange) || Boolean(opts.showDataLabelsRange))
					strXml += '</c:dLbls>'
				}

				// 'c:marker' tag: `lineDataSymbol`
				if (chartType === CHART_TYPE.LINE || chartType === CHART_TYPE.RADAR) {
					strXml += '<c:marker>'
					strXml += '  <c:symbol val="' + opts.lineDataSymbol + '"/>'
					if (opts.lineDataSymbolSize) strXml += `<c:size val="${opts.lineDataSymbolSize}"/>` // Defaults to "auto" otherwise (but this is usually too small, so there is a default)
					strXml += '  <c:spPr>'
					// NateRadebaugh/transparent-markers: `'transparent'` → noFill (same as series fill)
					{
						const markerColor = obj.color ?? chartColors[dataIndex + 1 > chartColors.length ? Math.floor(Math.random() * chartColors.length) : dataIndex]
						strXml += markerColor === 'transparent'
							? '    <a:noFill/>'
							: `    <a:solidFill>${createColorElement(markerColor)}</a:solidFill>`
					}
					strXml += `    <a:ln w="${opts.lineDataSymbolLineSize}" cap="flat"><a:solidFill>${createColorElement(opts.lineDataSymbolLineColor || seriesColor)}</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>`
					strXml += '    <a:effectLst/>'
					strXml += '  </c:spPr>'
					strXml += '</c:marker>'
				}

				// Allow users with a single data set to pass their own array of colors (check for this using != ours)
				// Color chart bars various colors when >1 color
				// NOTE: `<c:dPt>` created with various colors will change PPT legend by design so each dataPt/color is an legend item!
				if (
					(chartType === CHART_TYPE.BAR || chartType === CHART_TYPE.BAR3D) &&
					data.length === 1 &&
					((opts.chartColors && opts.chartColors !== BARCHART_COLORS && opts.chartColors.length > 1) || (opts.invertedColors?.length))
				) {
					// Series Data Point colors
					objValues.forEach((value, index) => {
						const arrColors = value < 0 ? opts.invertedColors || opts.chartColors || BARCHART_COLORS : opts.chartColors || []

						strXml += '  <c:dPt>'
						strXml += `    <c:idx val="${index}"/>`
						strXml += '      <c:invertIfNegative val="0"/>'
						strXml += '    <c:bubble3D val="0"/>'
						strXml += '    <c:spPr>'
						if (opts.lineSize === 0) {
							strXml += '<a:ln><a:noFill/></a:ln>'
						} else if (chartType === CHART_TYPE.BAR || chartType === CHART_TYPE.BAR3D) {
							// CT_DPt.spPr is CT_ShapeProperties: series color is a fill (`a:solidFill`), not an outline (`a:ln`).
							strXml += `<a:solidFill>${createColorElement(arrColors[index % arrColors.length])}</a:solidFill>`
						} else {
							strXml += `<a:ln><a:solidFill>${createColorElement(arrColors[index % arrColors.length])}</a:solidFill></a:ln>`
						}
						strXml += createShadowElement(resolveShadowOptions(opts.shadow))
						strXml += '    </c:spPr>'
						strXml += '  </c:dPt>'
					})
				}

				// 2: "Categories"
				{
					strXml += '<c:cat>'
					if (opts.catLabelFormatCode) {
						// Use 'numRef' as catLabelFormatCode implies that we are expecting numbers here
						strXml += '  <c:numRef>'
						strXml += `    <c:f>Sheet1!$A$2:$A$${objLabels[0].length + 1}</c:f>`
						strXml += '    <c:numCache>'
						strXml += '      <c:formatCode>' + (opts.catLabelFormatCode || 'General') + '</c:formatCode>'
						strXml += `      <c:ptCount val="${objLabels[0].length}"/>`
						objLabels[0].forEach((label, idx) => (strXml += `<c:pt idx="${idx}"><c:v>${encodeXmlEntities(label)}</c:v></c:pt>`))
						strXml += '    </c:numCache>'
						strXml += '  </c:numRef>'
					} else if (objLabels.length === 1) {
						// Single-level cats: use strRef (Google Sheets/Slides ignore multiLvlStrRef with one level)
						// christiankiely/fix-categories-google-slides
						strXml += '  <c:strRef>'
						strXml += `    <c:f>Sheet1!$A$2:$${getExcelColName(objLabels.length)}$${objLabels[0].length + 1}</c:f>`
						strXml += '    <c:strCache>'
						strXml += `      <c:ptCount val="${objLabels[0].length}"/>`
						objLabels[0].forEach((label, idx) => (strXml += `<c:pt idx="${idx}"><c:v>${encodeXmlEntities(label)}</c:v></c:pt>`))
						strXml += '    </c:strCache>'
						strXml += '  </c:strRef>'
					} else {
						strXml += '  <c:multiLvlStrRef>'
						strXml += `    <c:f>Sheet1!$A$2:$${getExcelColName(objLabels.length)}$${objLabels[0].length + 1}</c:f>`
						strXml += '    <c:multiLvlStrCache>'
						strXml += `      <c:ptCount val="${objLabels[0].length}"/>`
						objLabels.forEach(labelsGroup => {
							strXml += '<c:lvl>'
							labelsGroup.forEach((label, idx) => (strXml += `<c:pt idx="${idx}"><c:v>${encodeXmlEntities(label)}</c:v></c:pt>`))
							strXml += '</c:lvl>'
						})
						strXml += '    </c:multiLvlStrCache>'
						strXml += '  </c:multiLvlStrRef>'
					}
					strXml += '</c:cat>'
				}

				// 3: "Values"
				{
					strXml += '<c:val>'
					strXml += '  <c:numRef>'
					strXml += `<c:f>Sheet1!$${getExcelColName(dataIndex + objLabels.length + 1)}$2:$${getExcelColName(dataIndex + objLabels.length + 1)}$${objLabels[0].length + 1}</c:f>`
					strXml += '    <c:numCache>'
					strXml += '      <c:formatCode>' + (opts.valLabelFormatCode || opts.dataTableFormatCode || 'General') + '</c:formatCode>'
					strXml += `      <c:ptCount val="${objLabels[0].length}"/>`
					objValues.forEach((value, idx) => (strXml += `<c:pt idx="${idx}"><c:v>${value || value === 0 ? value : ''}</c:v></c:pt>`))
					strXml += '    </c:numCache>'
					strXml += '  </c:numRef>'
					strXml += '</c:val>'
				}

				// Option: `smooth`
				if (chartType === CHART_TYPE.LINE) strXml += '<c:smooth val="' + (opts.lineSmooth ? '1' : '0') + '"/>'

				// LanPodder/master: custom Y error bars (symmetric plus/minus from `errorrate`)
				if (seriesHasErrorrate(obj)) {
					const errCol = getErrorrateExcelCol(sheetData, objLabels.length, dataIndex)
					const lastRow = (objLabels[0]?.length ?? 0) + 1
					strXml += genXmlErrBars(obj.errorrate ?? [], errCol, lastRow)
				}

				// Toukyh/fix-custom-label: series extLst after cat/val(/smooth) per CT_*Ser
				// Formula matches the series values column (same as upstream); cache carries the display text.
				if (labelsRange) {
					const col = getExcelColName(dataIndex + objLabels.length + 1)
					const lastRow = (objLabels[0]?.length ?? labelsRange.length) + 1
					strXml += genXmlDataLabelsRangeSerExt(labelsRange, `Sheet1!$${col}$2:$${col}$${lastRow}`)
				}

				// 4: Close "SERIES"
				strXml += '</c:ser>'
			})

			// 3: "Data Labels"
			{
				strXml += '  <c:dLbls>'
				strXml += `    <c:numFmt formatCode="${encodeXmlEntities(opts.dataLabelFormatCode) || 'General'}" sourceLinked="0"/>`
				strXml += '    <c:txPr>'
				strXml += '      <a:bodyPr/>'
				strXml += '      <a:lstStyle/>'
				strXml += '      <a:p><a:pPr>'
				strXml += `        <a:defRPr b="${opts.dataLabelFontBold ? 1 : 0}" i="${opts.dataLabelFontItalic ? 1 : 0}" strike="noStrike" sz="${Math.round((opts.dataLabelFontSize || DEF_FONT_SIZE) * 100)}" u="none">`
				strXml += '          <a:solidFill>' + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + '</a:solidFill>'
				strXml += '          <a:latin typeface="' + (encodeXmlEntities(opts.dataLabelFontFace) || 'Arial') + '"/>'
				strXml += '        </a:defRPr>'
				strXml += '      </a:pPr></a:p>'
				strXml += '    </c:txPr>'
				if (opts.dataLabelPosition) strXml += `<c:dLblPos val="${opts.dataLabelPosition}"/>`
				strXml += '    <c:showLegendKey val="0"/>'
				strXml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>'
				strXml += '    <c:showCatName val="0"/>'
				strXml += '    <c:showSerName val="' + (opts.showSerName ? '1' : '0') + '"/>'
				strXml += '    <c:showPercent val="0"/>'
				strXml += '    <c:showBubbleSize val="0"/>'
				strXml += `    <c:showLeaderLines val="${opts.showLeaderLines ? '1' : '0'}"/>`
				strXml += '  </c:dLbls>'
			}

			// 4: Add more chart options (gapWidth, line Marker, etc.)
			if (chartType === CHART_TYPE.BAR) {
				strXml += `  <c:gapWidth val="${opts.barGapWidthPct}"/>`
				// istevkovski/prioritize-overlap: honor barOverlapPct even when grouping is stacked
				// (also treat 0 as explicit — upstream used a truthy check that skipped 0)
				{
					const overlap =
						typeof opts.barOverlapPct === 'number'
							? opts.barOverlapPct
							: (opts.barGrouping || '').includes('tacked')
								? 100
								: 0
					strXml += `  <c:overlap val="${overlap}"/>`
				}
			} else if (chartType === CHART_TYPE.BAR3D) {
				strXml += `  <c:gapWidth val="${opts.barGapWidthPct}"/>`
				strXml += `  <c:gapDepth val="${opts.barGapDepthPct}"/>`
				strXml += '  <c:shape val="' + opts.bar3DShape + '"/>'
			} else if (chartType === CHART_TYPE.LINE) {
				strXml += '  <c:marker val="1"/>'
			}

			// 5: Add axisId (NOTE: order matters - category comes first)
			strXml += `<c:axId val="${catAxisId}"/><c:axId val="${valAxisId}"/><c:axId val="${AXIS_ID_SERIES_PRIMARY}"/>`

			// 6: Close Chart tag
			strXml += `</c:${chartType}Chart>`

			// end switch
			break

		case CHART_TYPE.SCATTER:
			/*
				`data` = [
					{ name:'X-Axis',    values:[1,2,3,4,5,6,7,8,9,10,11,12] },
					{ name:'Y-Value 1', values:[13, 20, 21, 25] },
					{ name:'Y-Value 2', values:[ 1,  2,  5,  9] }
				];
            */

			// 1: Start Chart
			strXml += '<c:' + chartType + 'Chart>'
			strXml += '<c:scatterStyle val="lineMarker"/>'
			strXml += '<c:varyColors val="0"/>'

			// 2: Series: (One for each Y-Axis)
			colorIndex = -1
			data.filter((_obj, idx) => idx > 0).forEach((obj, idx) => {
				colorIndex++
				const objLabels = obj.labels ?? []
				const objValues = obj.values ?? []
				strXml += '<c:ser>'
				strXml += `  <c:idx val="${idx}"/>`
				strXml += `  <c:order val="${idx}"/>`
				strXml += '  <c:tx>'
				strXml += '    <c:strRef>'
				strXml += `      <c:f>Sheet1!$${getExcelColName(idx + 2)}$1</c:f>`
				strXml += '      <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>' + encodeXmlEntities(obj.name) + '</c:v></c:pt></c:strCache>'
				strXml += '    </c:strRef>'
				strXml += '  </c:tx>'

				// 'c:spPr': Fill, Border, Line, LineStyle (dash, etc.), Shadow
				strXml += '  <c:spPr>'
				{
					const tmpSerColor = obj.color ?? chartColors[colorIndex % chartColors.length]

					if (tmpSerColor === 'transparent') {
						strXml += '<a:noFill/>'
					} else if (opts.chartColorsOpacity) {
						strXml += '<a:solidFill>' + createColorElement(tmpSerColor, '<a:alpha val="' + Math.round(opts.chartColorsOpacity * 1000).toString() + '"/>') + '</a:solidFill>'
					} else {
						strXml += '<a:solidFill>' + createColorElement(tmpSerColor) + '</a:solidFill>'
					}

					if (opts.lineSize === 0) {
						strXml += '<a:ln><a:noFill/></a:ln>'
					} else {
						strXml += `<a:ln w="${valToPts(opts.lineSize)}" cap="${createLineCap(opts.lineCap)}"><a:solidFill>${createColorElement(tmpSerColor)}</a:solidFill>`
						strXml += `<a:prstDash val="${resolveSeriesLineDash(obj.lineDash, opts.lineDash)}"/><a:round/></a:ln>`
					}

					// Shadow
					strXml += createShadowElement(resolveShadowOptions(opts.shadow))
				}
				strXml += '  </c:spPr>'

				// 'c:marker' tag: `lineDataSymbol`
				{
					strXml += '<c:marker>'
					strXml += '  <c:symbol val="' + opts.lineDataSymbol + '"/>'
					if (opts.lineDataSymbolSize) {
						// Defaults to "auto" otherwise (but this is usually too small, so there is a default)
						strXml += `<c:size val="${opts.lineDataSymbolSize}"/>`
					}
					strXml += '<c:spPr>'
					// NateRadebaugh/transparent-markers: `'transparent'` → noFill (same as series fill)
					{
						const markerColor = obj.color ?? chartColors[idx + 1 > chartColors.length ? Math.floor(Math.random() * chartColors.length) : idx]
						strXml += markerColor === 'transparent'
							? '<a:noFill/>'
							: `<a:solidFill>${createColorElement(markerColor)}</a:solidFill>`
					}
					strXml += `<a:ln w="${opts.lineDataSymbolLineSize}" cap="flat"><a:solidFill>${createColorElement(opts.lineDataSymbolLineColor || obj.color || chartColors[colorIndex % chartColors.length])}</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>`
					strXml += '<a:effectLst/>'
					strXml += '</c:spPr>'
					strXml += '</c:marker>'
				}

				// Option: scatter data point labels
				if (opts.showLabel) {
					const chartUuid = getUuid('-xxxx-xxxx-xxxx-xxxxxxxxxxxx')
					if (objLabels[0] && (opts.dataLabelFormatScatter === 'custom' || opts.dataLabelFormatScatter === 'customXY')) {
						strXml += '<c:dLbls>'
						objLabels[0].forEach((label, idx) => {
							if (opts.dataLabelFormatScatter === 'custom' || opts.dataLabelFormatScatter === 'customXY') {
								strXml += '  <c:dLbl>'
								strXml += `    <c:idx val="${idx}"/>`
								strXml += '    <c:tx>'
								strXml += '      <c:rich>'
								strXml += '            <a:bodyPr>'
								strXml += '                <a:spAutoFit/>'
								strXml += '            </a:bodyPr>'
								strXml += '            <a:lstStyle/>'
								strXml += '            <a:p>'
								strXml += '                <a:pPr>'
								strXml += `                    ${genXmlDataLabelDefRPr(opts)}`
								strXml += '                </a:pPr>'
								strXml += '              <a:r>'
								strXml += '                    <a:rPr lang="' + (opts.lang || 'en-US') + '" dirty="0"/>'
								strXml += '                    <a:t>' + encodeXmlEntities(label) + '</a:t>'
								strXml += '              </a:r>'
								// Apply XY values at end of custom label
								// Do not apply the values if the label was empty or just spaces
								// This allows for selective labelling where required
								if (opts.dataLabelFormatScatter === 'customXY' && !/^ *$/.test(label)) {
									strXml += '              <a:r>'
									strXml += '                  <a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0" dirty="0"/>'
									strXml += '                  <a:t> (</a:t>'
									strXml += '              </a:r>'
									strXml += '              <a:fld id="{' + getUuid('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx') + '}" type="XVALUE">'
									strXml += '                  <a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0"/>'
									strXml += '                  <a:pPr>'
									strXml += `                      ${genXmlDataLabelDefRPr(opts)}`
									strXml += '                  </a:pPr>'
									strXml += '                  <a:t>[' + encodeXmlEntities(obj.name) + '</a:t>'
									strXml += '              </a:fld>'
									strXml += '              <a:r>'
									strXml += '                  <a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0" dirty="0"/>'
									strXml += '                  <a:t>, </a:t>'
									strXml += '              </a:r>'
									strXml += '              <a:fld id="{' + getUuid('xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx') + '}" type="YVALUE">'
									strXml += '                  <a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0"/>'
									strXml += '                  <a:pPr>'
									strXml += `                      ${genXmlDataLabelDefRPr(opts)}`
									strXml += '                  </a:pPr>'
									strXml += '                  <a:t>[' + encodeXmlEntities(obj.name) + ']</a:t>'
									strXml += '              </a:fld>'
									strXml += '              <a:r>'
									strXml += '                  <a:rPr lang="' + (opts.lang || 'en-US') + '" baseline="0" dirty="0"/>'
									strXml += '                  <a:t>)</a:t>'
									strXml += '              </a:r>'
									strXml += '              <a:endParaRPr lang="' + (opts.lang || 'en-US') + '" dirty="0"/>'
								}
								strXml += '            </a:p>'
								strXml += '      </c:rich>'
								strXml += '    </c:tx>'
								strXml += '    <c:spPr>'
								strXml += '        <a:noFill/>'
								strXml += '        <a:ln>'
								strXml += '            <a:noFill/>'
								strXml += '        </a:ln>'
								strXml += '        <a:effectLst/>'
								strXml += '    </c:spPr>'
								if (opts.dataLabelPosition) strXml += `<c:dLblPos val="${opts.dataLabelPosition}"/>`
								strXml += '    <c:showLegendKey val="0"/>'
								strXml += '    <c:showVal val="0"/>'
								strXml += '    <c:showCatName val="0"/>'
								strXml += '    <c:showSerName val="0"/>'
								strXml += '    <c:showPercent val="0"/>'
								strXml += '    <c:showBubbleSize val="0"/>'
								strXml += '       <c:showLeaderLines val="1"/>'
								strXml += '    <c:extLst>'
								strXml += '      <c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart"/>'
								strXml += '      <c:ext uri="{C3380CC4-5D6E-409C-BE32-E72D297353CC}" xmlns:c16="http://schemas.microsoft.com/office/drawing/2014/chart">'
								strXml += `            <c16:uniqueId val="{${'00000000'.substring(0, 8 - (idx + 1).toString().length).toString()}${idx + 1}${chartUuid}}"/>`
								strXml += '      </c:ext>'
								strXml += '        </c:extLst>'
								strXml += '</c:dLbl>'
							}
						})
						strXml += '</c:dLbls>'
					}
					if (opts.dataLabelFormatScatter === 'XY') {
						strXml += '<c:dLbls>'
						strXml += '    <c:spPr>'
						strXml += '        <a:noFill/>'
						strXml += '        <a:ln>'
						strXml += '            <a:noFill/>'
						strXml += '        </a:ln>'
						strXml += '          <a:effectLst/>'
						strXml += '    </c:spPr>'
						strXml += '    <c:txPr>'
						strXml += '        <a:bodyPr>'
						strXml += '            <a:spAutoFit/>'
						strXml += '        </a:bodyPr>'
						strXml += '        <a:lstStyle/>'
						strXml += '        <a:p>'
						strXml += '            <a:pPr>'
						strXml += `                ${genXmlDataLabelDefRPr(opts)}`
						strXml += '            </a:pPr>'
						strXml += '            <a:endParaRPr lang="en-US"/>'
						strXml += '        </a:p>'
						strXml += '    </c:txPr>'
						if (opts.dataLabelPosition) strXml += `<c:dLblPos val="${opts.dataLabelPosition}"/>`
						strXml += '    <c:showLegendKey val="0"/>'
						strXml += ` <c:showVal val="${opts.showLabel ? '1' : '0'}"/>`
						strXml += ` <c:showCatName val="${opts.showLabel ? '1' : '0'}"/>`
						strXml += ` <c:showSerName val="${opts.showSerName ? '1' : '0'}"/>`
						strXml += '    <c:showPercent val="0"/>'
						strXml += '    <c:showBubbleSize val="0"/>'
						strXml += '    <c:extLst>'
						strXml += '        <c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">'
						strXml += '            <c15:showLeaderLines val="1"/>'
						strXml += '        </c:ext>'
						strXml += '    </c:extLst>'
						strXml += '</c:dLbls>'
					}
				}

				// Color bar chart bars various colors
				// Allow users with a single data set to pass their own array of colors (check for this using != ours)
				if (data.length === 1 && opts.chartColors !== BARCHART_COLORS) {
					// Series Data Point colors
					objValues.forEach((value, index) => {
						const arrColors = value < 0 ? opts.invertedColors || opts.chartColors || BARCHART_COLORS : opts.chartColors || []

						strXml += '  <c:dPt>'
						strXml += `    <c:idx val="${index}"/>`
						strXml += '      <c:invertIfNegative val="0"/>'
						strXml += '    <c:bubble3D val="0"/>'
						strXml += '    <c:spPr>'
						if (opts.lineSize === 0) {
							strXml += '<a:ln><a:noFill/></a:ln>'
						} else {
							strXml += `<a:solidFill>${createColorElement(arrColors[index % arrColors.length])}</a:solidFill>`
						}
						strXml += createShadowElement(resolveShadowOptions(opts.shadow))
						strXml += '    </c:spPr>'
						strXml += '  </c:dPt>'
					})
				}

				// 3: "Values": Scatter Chart has 2: `xVal` and `yVal`
				{
					// X-Axis is always the same
					strXml += '<c:xVal>'
					strXml += '  <c:numRef>'
					strXml += `    <c:f>Sheet1!$A$2:$A$${firstValues.length + 1}</c:f>`
					strXml += '    <c:numCache>'
					strXml += '      <c:formatCode>General</c:formatCode>'
					strXml += `      <c:ptCount val="${firstValues.length}"/>`
					firstValues.forEach((value, idx) => {
						strXml += `<c:pt idx="${idx}"><c:v>${value || value === 0 ? value : ''}</c:v></c:pt>`
					})
					strXml += '    </c:numCache>'
					strXml += '  </c:numRef>'
					strXml += '</c:xVal>'

					// Y-Axis vals are this object's `values`
					strXml += '<c:yVal>'
					strXml += '  <c:numRef>'
					strXml += `    <c:f>Sheet1!$${getExcelColName(idx + 2)}$2:$${getExcelColName(idx + 2)}$${firstValues.length + 1}</c:f>`
					strXml += '    <c:numCache>'
					strXml += '      <c:formatCode>General</c:formatCode>'
					// NOTE: Use pt count and iterate over data[0] (X-Axis) as user can have more values than data (eg: timeline where only first few months are populated)
					strXml += `      <c:ptCount val="${firstValues.length}"/>`
					firstValues.forEach((_value, idx) => {
						strXml += `<c:pt idx="${idx}"><c:v>${objValues[idx] || objValues[idx] === 0 ? objValues[idx] : ''}</c:v></c:pt>`
					})
					strXml += '    </c:numCache>'
					strXml += '  </c:numRef>'
					strXml += '</c:yVal>'
				}

				// Option: `smooth`
				strXml += '<c:smooth val="' + (opts.lineSmooth ? '1' : '0') + '"/>'

				// 4: Close "SERIES"
				strXml += '</c:ser>'
			})

			// 3: Data Labels
			{
				strXml += '  <c:dLbls>'
				strXml += `    <c:numFmt formatCode="${encodeXmlEntities(opts.dataLabelFormatCode) || 'General'}" sourceLinked="0"/>`
				strXml += '    <c:txPr>'
				strXml += '      <a:bodyPr/>'
				strXml += '      <a:lstStyle/>'
				strXml += '      <a:p><a:pPr>'
				strXml += `        <a:defRPr b="${opts.dataLabelFontBold ? '1' : '0'}" i="${opts.dataLabelFontItalic ? '1' : '0'}" strike="noStrike" sz="${Math.round((opts.dataLabelFontSize || DEF_FONT_SIZE) * 100)}" u="none">`
				strXml += '          <a:solidFill>' + createColorElement(opts.dataLabelColor || DEF_FONT_COLOR) + '</a:solidFill>'
				strXml += '          <a:latin typeface="' + (encodeXmlEntities(opts.dataLabelFontFace) || 'Arial') + '"/>'
				strXml += '        </a:defRPr>'
				strXml += '      </a:pPr></a:p>'
				strXml += '    </c:txPr>'
				if (opts.dataLabelPosition) strXml += `<c:dLblPos val="${opts.dataLabelPosition}"/>`
				strXml += '    <c:showLegendKey val="0"/>'
				strXml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>'
				strXml += '    <c:showCatName val="0"/>'
				strXml += '    <c:showSerName val="' + (opts.showSerName ? '1' : '0') + '"/>'
				strXml += '    <c:showPercent val="0"/>'
				strXml += '    <c:showBubbleSize val="0"/>'
				strXml += '  </c:dLbls>'
			}

			// 4: Add axis Id (NOTE: order matters - category comes first)
			strXml += `<c:axId val="${catAxisId}"/><c:axId val="${valAxisId}"/>`

			// 5: Close Chart tag
			strXml += '</c:' + chartType + 'Chart>'

			// end switch
			break

		case CHART_TYPE.BUBBLE:
		case CHART_TYPE.BUBBLE3D:
			/*
				`data` = [
					{ name:'X-Axis',     values:[1,2,3,4,5,6,7,8,9,10,11,12] },
					{ name:'Y-Values 1', values:[13, 20, 21, 25], sizes:[10, 5, 20, 15] },
					{ name:'Y-Values 2', values:[ 1,  2,  5,  9], sizes:[ 5, 3,  9,  3] }
				];
            */

			// 1: Start Chart
			strXml += '<c:bubbleChart>'
			strXml += '<c:varyColors val="0"/>'

			// 2: Series: (One for each Y-Axis)
			colorIndex = -1
			data.filter((_obj, idx) => idx > 0).forEach((obj, idx) => {
				colorIndex++
				const objValues = obj.values ?? []
				const objSizes = obj.sizes ?? []
				strXml += '<c:ser>'
				strXml += `  <c:idx val="${idx}"/>`
				strXml += `  <c:order val="${idx}"/>`

				// A: `<c:tx>`
				strXml += '  <c:tx>'
				strXml += '    <c:strRef>'
				strXml += '      <c:f>Sheet1!$' + getExcelColName(idxColLtr + 1) + '$1</c:f>'
				strXml += '      <c:strCache><c:ptCount val="1"/><c:pt idx="0"><c:v>' + encodeXmlEntities(obj.name) + '</c:v></c:pt></c:strCache>'
				strXml += '    </c:strRef>'
				strXml += '  </c:tx>'

				// B: '<c:spPr>': Fill, Border, Line, LineStyle (dash, etc.), Shadow
				{
					strXml += '<c:spPr>'

					const tmpSerColor = obj.color ?? chartColors[colorIndex % chartColors.length]

					if (tmpSerColor === 'transparent') {
						strXml += '<a:noFill/>'
					} else if (opts.chartColorsOpacity) {
						strXml += `<a:solidFill>${createColorElement(tmpSerColor, '<a:alpha val="' + Math.round(opts.chartColorsOpacity * 1000).toString() + '"/>')}</a:solidFill>`
					} else {
						strXml += '<a:solidFill>' + createColorElement(tmpSerColor) + '</a:solidFill>'
					}

					if (opts.lineSize === 0) {
						strXml += '<a:ln><a:noFill/></a:ln>'
					} else if (opts.dataBorder) {
						strXml += `<a:ln w="${valToPts(opts.dataBorder.pt)}" cap="flat"><a:solidFill>${createColorElement(opts.dataBorder.color)}</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>`
					} else {
						strXml += `<a:ln w="${valToPts(opts.lineSize)}" cap="flat"><a:solidFill>${createColorElement(tmpSerColor)}</a:solidFill>`
						strXml += `<a:prstDash val="${resolveSeriesLineDash(obj.lineDash, opts.lineDash)}"/><a:round/></a:ln>`
					}

					// Shadow
					strXml += createShadowElement(resolveShadowOptions(opts.shadow))

					strXml += '</c:spPr>'
				}

				// C: '<c:dLbls>' "Data Labels"
				// Let it be defaulted for now

				// D: '<c:xVal>'/'<c:yVal>' "Values": Scatter Chart has 2: `xVal` and `yVal`
				{
					// X-Axis is always the same
					strXml += '<c:xVal>'
					strXml += '  <c:numRef>'
					strXml += `    <c:f>Sheet1!$A$2:$A$${firstValues.length + 1}</c:f>`
					strXml += '    <c:numCache>'
					strXml += '      <c:formatCode>General</c:formatCode>'
					strXml += `      <c:ptCount val="${firstValues.length}"/>`
					firstValues.forEach((value, idx) => {
						strXml += `<c:pt idx="${idx}"><c:v>${value || value === 0 ? value : ''}</c:v></c:pt>`
					})
					strXml += '    </c:numCache>'
					strXml += '  </c:numRef>'
					strXml += '</c:xVal>'

					// Y-Axis vals are this object's `values`
					strXml += '<c:yVal>'
					strXml += '  <c:numRef>'
					strXml += `<c:f>Sheet1!$${getExcelColName(idxColLtr + 1)}$2:$${getExcelColName(idxColLtr + 1)}$${firstValues.length + 1}</c:f>`
					idxColLtr++
					strXml += '    <c:numCache>'
					strXml += '      <c:formatCode>General</c:formatCode>'
					// NOTE: Use pt count and iterate over data[0] (X-Axis) as user can have more values than data (eg: timeline where only first few months are populated)
					strXml += `      <c:ptCount val="${firstValues.length}"/>`
					firstValues.forEach((_value, idx) => {
						strXml += `<c:pt idx="${idx}"><c:v>${objValues[idx] || objValues[idx] === 0 ? objValues[idx] : ''}</c:v></c:pt>`
					})
					strXml += '    </c:numCache>'
					strXml += '  </c:numRef>'
					strXml += '</c:yVal>'
				}

				// E: '<c:bubbleSize>'
				strXml += '  <c:bubbleSize>'
				strXml += '    <c:numRef>'
				strXml += `<c:f>Sheet1!$${getExcelColName(idxColLtr + 1)}$2:$${getExcelColName(idxColLtr + 1)}$${objSizes.length + 1}</c:f>`
				idxColLtr++
				strXml += '      <c:numCache>'
				strXml += '        <c:formatCode>General</c:formatCode>'
				strXml += `           <c:ptCount val="${objSizes.length}"/>`
				objSizes.forEach((value, idx) => {
					strXml += `<c:pt idx="${idx}"><c:v>${value || ''}</c:v></c:pt>`
				})
				strXml += '      </c:numCache>'
				strXml += '    </c:numRef>'
				strXml += '  </c:bubbleSize>'
				strXml += '  <c:bubble3D val="' + (chartType === CHART_TYPE.BUBBLE3D ? '1' : '0') + '"/>'

				// F: Close "SERIES"
				strXml += '</c:ser>'
			})

			// 3: Data Labels
			{
				strXml += '<c:dLbls>'
				strXml += `<c:numFmt formatCode="${encodeXmlEntities(opts.dataLabelFormatCode) || 'General'}" sourceLinked="0"/>`
				strXml += '<c:txPr><a:bodyPr/><a:lstStyle/><a:p><a:pPr>'
				strXml += `<a:defRPr b="${opts.dataLabelFontBold ? 1 : 0}" i="${opts.dataLabelFontItalic ? 1 : 0}" strike="noStrike" sz="${Math.round(
					Math.round(opts.dataLabelFontSize || DEF_FONT_SIZE) * 100
				)}" u="none">`
				strXml += `<a:solidFill>${createColorElement(opts.dataLabelColor || DEF_FONT_COLOR)}</a:solidFill>`
				strXml += `<a:latin typeface="${encodeXmlEntities(opts.dataLabelFontFace) || 'Arial'}"/>`
				strXml += '</a:defRPr></a:pPr></a:p></c:txPr>'
				if (opts.dataLabelPosition) strXml += `<c:dLblPos val="${opts.dataLabelPosition}"/>`
				strXml += '<c:showLegendKey val="0"/>'
				strXml += `<c:showVal val="${opts.showValue ? '1' : '0'}"/>`
				strXml += `<c:showCatName val="0"/><c:showSerName val="${opts.showSerName ? '1' : '0'}"/><c:showPercent val="0"/><c:showBubbleSize val="0"/>`
				strXml += '<c:extLst>'
				strXml += '  <c:ext uri="{CE6537A1-D6FC-4f65-9D91-7224C49458BB}" xmlns:c15="http://schemas.microsoft.com/office/drawing/2012/chart">'
				strXml += '    <c15:showLeaderLines val="' + (opts.showLeaderLines ? '1' : '0') + '"/>'
				strXml += '  </c:ext>'
				strXml += '</c:extLst>'
				strXml += '</c:dLbls>'
			}

			// 4: Bubble options
			// strXml += '  <c:bubbleScale val="100"/>';
			// strXml += '  <c:showNegBubbles val="0"/>';
			// Commented out to let it default to PPT until we create options

			// 5: AxisId (NOTE: order matters - category comes first)
			strXml += `<c:axId val="${catAxisId}"/><c:axId val="${valAxisId}"/>`

			// 6: Close Chart tag
			strXml += '</c:bubbleChart>'

			// end switch
			break

		case CHART_TYPE.DOUGHNUT:
		case CHART_TYPE.PIE: {
			const optsChartData: IOptsChartData = data[0]
			// NOTE: chart data arrays are always populated at render time; guarded locals keep this null-safe without changing output
			const chartDataLabels = optsChartData.labels ?? []
			const chartDataValues = optsChartData.values ?? []

			/* EX:
				data: [
				 {
				   name: 'Project Status',
				   labels: ['Red', 'Amber', 'Green', 'Unknown'],
				   values: [10, 20, 38, 2]
				 }
				]
            */

			// 1: Start Chart
			strXml += '<c:' + chartType + 'Chart>'
			strXml += '  <c:varyColors val="1"/>'
			strXml += '<c:ser>'
			strXml += '  <c:idx val="0"/>'
			strXml += '  <c:order val="0"/>'
			strXml += '  <c:tx>'
			strXml += '    <c:strRef>'
			strXml += '      <c:f>Sheet1!$B$1</c:f>'
			strXml += '      <c:strCache>'
			strXml += '        <c:ptCount val="1"/>'
			strXml += '        <c:pt idx="0"><c:v>' + encodeXmlEntities(optsChartData.name) + '</c:v></c:pt>'
			strXml += '      </c:strCache>'
			strXml += '    </c:strRef>'
			strXml += '  </c:tx>'
			strXml += '  <c:spPr>'
			strXml += '    <a:solidFill><a:schemeClr val="accent1"/></a:solidFill>'
			strXml += '    <a:ln w="9525" cap="flat"><a:solidFill><a:srgbClr val="F9F9F9"/></a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>'
			if (opts.dataNoEffects) {
				strXml += '<a:effectLst/>'
			} else {
				strXml += createShadowElement(resolveShadowOptions(opts.shadow))
			}
			strXml += '  </c:spPr>'
			// strXml += '<c:explosion val="0"/>'

			// 2: "Data Point" block for every data row
			chartDataLabels[0].forEach((_label, idx) => {
				strXml += '<c:dPt>'
				strXml += ` <c:idx val="${idx}"/>`
				strXml += ' <c:bubble3D val="0"/>'
				strXml += ' <c:spPr>'
				strXml += `<a:solidFill>${createColorElement(
					chartColors[idx + 1 > chartColors.length ? Math.floor(Math.random() * chartColors.length) : idx]
				)}</a:solidFill>`
				if (opts.dataBorder) {
					strXml += `<a:ln w="${valToPts(opts.dataBorder.pt)}" cap="flat"><a:solidFill>${createColorElement(
						opts.dataBorder.color
					)}</a:solidFill><a:prstDash val="solid"/><a:round/></a:ln>`
				}
				strXml += createShadowElement(resolveShadowOptions(opts.shadow))
				strXml += '  </c:spPr>'
				strXml += '</c:dPt>'
			})

			// 3: Series-level data labels. Per-point `c:dLbl` is additive (sparse dataLabels only).
			strXml += '<c:dLbls>'
			strXml += genXmlCustomDataLabels(optsChartData.dataLabels, opts)
			strXml += ` <c:numFmt formatCode="${encodeXmlEntities(opts.dataLabelFormatCode) || 'General'}" sourceLinked="0"/>`
			strXml += '    <c:txPr>'
			strXml += '      <a:bodyPr/>'
			strXml += '      <a:lstStyle/>'
			strXml += '      <a:p>'
			strXml += '        <a:pPr>'
			strXml += `          <a:defRPr sz="${Math.round((opts.dataLabelFontSize || DEF_FONT_SIZE) * 100)}" b="${opts.dataLabelFontBold ? '1' : '0'}" i="${opts.dataLabelFontItalic ? '1' : '0'}" u="none" strike="noStrike">`
			strXml += `            <a:solidFill><a:srgbClr val="000000"/></a:solidFill><a:latin typeface="${encodeXmlEntities(opts.dataLabelFontFace) || 'Arial'}"/>`
			strXml += '          </a:defRPr>'
			strXml += '        </a:pPr>'
			strXml += '      </a:p>'
			strXml += '    </c:txPr>'
			// Default pie labels to center; honor opts.dataLabelPosition when set (outEnd/inEnd/bestFit/ctr)
			strXml += chartType === CHART_TYPE.PIE ? `<c:dLblPos val="${opts.dataLabelPosition ?? 'ctr'}"/>` : ''
			strXml += '    <c:showLegendKey val="0"/>'
			strXml += '    <c:showVal val="' + (opts.showValue ? '1' : '0') + '"/>'
			strXml += '    <c:showCatName val="' + (opts.showLabel ? '1' : '0') + '"/>'
			strXml += '    <c:showSerName val="' + (opts.showSerName ? '1' : '0') + '"/>'
			strXml += '    <c:showPercent val="' + (opts.showPercent ? '1' : '0') + '"/>'
			strXml += '    <c:showBubbleSize val="0"/>'
			strXml += ` <c:showLeaderLines val="${opts.showLeaderLines ? '1' : '0'}"/>`
			strXml += '</c:dLbls>'

			// 2: "Categories"
			strXml += '<c:cat>'
			strXml += '  <c:strRef>'
			strXml += `    <c:f>Sheet1!$A$2:$A$${chartDataLabels[0].length + 1}</c:f>`
			strXml += '    <c:strCache>'
			strXml += `         <c:ptCount val="${chartDataLabels[0].length}"/>`
			chartDataLabels[0].forEach((label, idx) => {
				strXml += `<c:pt idx="${idx}"><c:v>${encodeXmlEntities(label)}</c:v></c:pt>`
			})
			strXml += '    </c:strCache>'
			strXml += '  </c:strRef>'
			strXml += '</c:cat>'

			// 3: Create vals
			strXml += '  <c:val>'
			strXml += '    <c:numRef>'
			strXml += `      <c:f>Sheet1!$B$2:$B$${chartDataLabels[0].length + 1}</c:f>`
			strXml += '      <c:numCache>'
			strXml += `           <c:ptCount val="${chartDataLabels[0].length}"/>`
			chartDataValues.forEach((value, idx) => {
				strXml += `<c:pt idx="${idx}"><c:v>${value || value === 0 ? value : ''}</c:v></c:pt>`
			})
			strXml += '      </c:numCache>'
			strXml += '    </c:numRef>'
			strXml += '  </c:val>'

			// 4: Close "SERIES"
			strXml += '  </c:ser>'
			strXml += `  <c:firstSliceAng val="${opts.firstSliceAng ? Math.round(opts.firstSliceAng) : 0}"/>`
			if (chartType === CHART_TYPE.DOUGHNUT) strXml += `<c:holeSize val="${typeof opts.holeSize === 'number' ? opts.holeSize : '50'}"/>`
			strXml += '</c:' + chartType + 'Chart>'

			// Done with Doughnut/Pie
			break
		}
		default:
			strXml += ''
			break
	}

	return strXml
}
