/**
 * Shared chart XML helpers.
 */

import { DEF_CHART_GRIDLINE, DEF_SHAPE_SHADOW } from '../core-enums'
import { ChartLineCap, IOptsChartData, OptsChartGridLine, ShadowProps } from '../core-interfaces'
import { createColorElement, valToPts } from '../gen-utils'

/**
 * Calc and return excel column name for a given 1-based column index
 * @param colIndex column index (1 = A)
 * @return column name
 * @example 1 returns 'A'
 * @example 27 returns 'AA'
 * @example 703 returns 'AAA'
 * @see LanPodder/master — previously capped at ZZ (52+)
 */
export function getExcelColName (colIndex: number): string {
	let n = colIndex
	let name = ''
	while (n > 0) {
		const rem = (n - 1) % 26
		name = String.fromCharCode(65 + rem) + name
		n = Math.floor((n - 1) / 26)
	}
	return name
}

export function seriesHasErrorrate (obj: IOptsChartData | undefined): boolean {
	return Array.isArray(obj?.errorrate) && (obj?.errorrate?.length ?? 0) > 0
}

export function countErrorrateSeries (data: IOptsChartData[]): number {
	let n = 0
	for (const obj of data) if (seriesHasErrorrate(obj)) n++
	return n
}

/**
 * 1-based Excel column for a series' packed errorrate column.
 * Layout: label levels | series values | packed errorrate columns (only series with errorrate).
 */
export function getErrorrateExcelCol (data: IOptsChartData[], labelLevels: number, serIdx: number): number {
	let packed = 0
	for (let i = 0; i < serIdx; i++) if (seriesHasErrorrate(data[i])) packed++
	return labelLevels + data.length + 1 + packed
}

/** Symmetric custom Y error bars (`c:errBars`) referencing a packed workbook column */
export function genXmlErrBars (errorrate: number[], sheetCol: number, lastRow: number): string {
	const col = getExcelColName(sheetCol)
	const numRef =
		'<c:numRef>' +
		`<c:f>Sheet1!$${col}$2:$${col}$${lastRow}</c:f>` +
		'<c:numCache>' +
		'<c:formatCode>General</c:formatCode>' +
		`<c:ptCount val="${errorrate.length}"/>` +
		errorrate.map((value, idx) => `<c:pt idx="${idx}"><c:v>${value || value === 0 ? value : ''}</c:v></c:pt>`).join('') +
		'</c:numCache>' +
		'</c:numRef>'
	return (
		'<c:errBars>' +
		'<c:errDir val="y"/>' +
		'<c:errBarType val="both"/>' +
		'<c:errValType val="cust"/>' +
		'<c:noEndCap val="0"/>' +
		`<c:plus>${numRef}</c:plus>` +
		`<c:minus>${numRef}</c:minus>` +
		'</c:errBars>'
	)
}

/**
 * Creates `a:innerShdw` or `a:outerShdw` depending on pass options `opts`.
 * @param {Object} opts optional shadow properties
 * @param {Object} defaults defaults for unspecified properties in `opts`
 * @see http://officeopenxml.com/drwSp-effects.php
 * @example { type: 'outer', blur: 3, offset: (23000 / 12700), angle: 90, color: '000000', opacity: 0.35, rotateWithShape: true };
 * @return {string} XML
 */
/**
 * Nominal ("coloured") brand for resolved shadow options. The symbol is module-private and unexported,
 * so `ResolvedShadowProps` values can ONLY be produced by `resolveShadowOptions` below - a hand-built
 * `Required<ShadowProps>` is not assignable. This statically guarantees anything reaching
 * `createShadowElement` has passed through the defaults-merge boundary.
 */
const shadowBrand: unique symbol = Symbol('resolvedShadow')
type ResolvedShadowProps = Omit<Required<ShadowProps>, 'preset'> & { preset?: ShadowProps['preset'], readonly [shadowBrand]: boolean }

/**
 * Resolve boundary: merge user shadow options over the documented defaults and brand the result.
 * Returns undefined for absent/invalid input. The only constructor of `ResolvedShadowProps` (no cast -
 * the brand is added by this factory).
 */
export function resolveShadowOptions (options: ShadowProps | undefined): ResolvedShadowProps | undefined {
	if (!options) return undefined
	if (typeof options !== 'object') {
		console.warn('`shadow` options must be an object. Ex: `{shadow: {type:\'none\'}}`')
		return undefined
	}
	return { ...DEF_SHAPE_SHADOW, ...options, [shadowBrand]: true }
}

export function createShadowElement (shadow: ResolvedShadowProps | undefined): string {
	if (!shadow) {
		return '<a:effectLst/>'
	}

	let strXml = '<a:effectLst>'
	const type = shadow.type
	const blur = valToPts(shadow.blur)
	const offset = valToPts(shadow.offset)
	const angle = Math.round(shadow.angle * 60000)
	const color = shadow.color
	const opacity = Math.round(shadow.opacity * 100000)
	const rotShape = shadow.rotateWithShape ? 1 : 0

	strXml += `<a:${type}Shdw sx="100000" sy="100000" kx="0" ky="0"  algn="bl" blurRad="${blur}" rotWithShape="${rotShape}" dist="${offset}" dir="${angle}">`
	strXml += `<a:srgbClr val="${color}">`
	strXml += `<a:alpha val="${opacity}"/></a:srgbClr>`
	strXml += `</a:${type}Shdw>`
	strXml += '</a:effectLst>'

	return strXml
}

/**
 * Create Grid Line Element
 * @param {OptsChartGridLine} glOpts {size, color, style}
 * @return {string} XML
 */
export function createGridLineElement (glOpts: OptsChartGridLine): string {
	let strXml = '<c:majorGridlines>'
	strXml += ' <c:spPr>'
	strXml += `  <a:ln w="${valToPts(glOpts.size || DEF_CHART_GRIDLINE.size)}" cap="${createLineCap(glOpts.cap || DEF_CHART_GRIDLINE.cap)}">`
	strXml += `  <a:solidFill>${createColorElement(glOpts.color || DEF_CHART_GRIDLINE.color)}</a:solidFill>`
	strXml += '   <a:prstDash val="' + (glOpts.style || DEF_CHART_GRIDLINE.style) + '"/><a:round/>'
	strXml += '  </a:ln>'
	strXml += ' </c:spPr>'
	strXml += '</c:majorGridlines>'

	return strXml
}

export function createLineCap (lineCap: ChartLineCap | undefined): string {
	if (!lineCap || lineCap === 'flat') {
		return 'flat'
	} else if (lineCap === 'square') {
		return 'sq'
	} else if (lineCap === 'round') {
		return 'rnd'
	} else {
		const neverLineCap: never = lineCap
		throw new Error(`Invalid chart line cap: ${neverLineCap}`)
	}
}

/** DrawingML `a:prstDash` / ST_PresetLineDashVal — same set as `line.dashType` / chart `lineDash`. */
const CHART_LINE_DASH_TYPES = ['solid', 'dash', 'dashDot', 'lgDash', 'lgDashDot', 'lgDashDotDot', 'sysDash', 'sysDot'] as const
export type ChartLineDashType = (typeof CHART_LINE_DASH_TYPES)[number]

/**
 * Per-series dash wins over chart-level `lineDash`. Unknown values fall back so we never
 * emit an illegal ST_PresetLineDashVal.
 */
export function resolveSeriesLineDash (seriesDash?: string, chartDash?: string): ChartLineDashType {
	if (seriesDash && (CHART_LINE_DASH_TYPES as readonly string[]).includes(seriesDash)) return seriesDash as ChartLineDashType
	if (chartDash && (CHART_LINE_DASH_TYPES as readonly string[]).includes(chartDash)) return chartDash as ChartLineDashType
	return 'solid'
}

/** Area charts use `c:grouping` / ST_Grouping (standard|stacked|percentStacked), not ST_BarGrouping. */
const AREA_GROUPING = ['standard', 'stacked', 'percentStacked'] as const
export type AreaGrouping = (typeof AREA_GROUPING)[number]

export function resolveAreaGrouping (barGrouping?: string): AreaGrouping {
	if (barGrouping && (AREA_GROUPING as readonly string[]).includes(barGrouping)) return barGrouping as AreaGrouping
	return 'standard'
}
