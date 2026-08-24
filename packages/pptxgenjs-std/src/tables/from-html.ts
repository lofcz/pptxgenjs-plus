import type PptxGenJS from 'pptxgenjs-plus'
import { paginateTable, type PaginateTableProps, type PaginateTableResult, type TablePresentation, type TableSlide } from './paginate'

/**
 * The slice of the DOM this helper reads.
 *
 * Structural rather than `HTMLTableElement` so it works without DOM lib types in scope and can be
 * driven by a plain object in tests. A real `<table>` conforms as-is.
 */
export interface HtmlCell {
	textContent?: string | null
	colSpan?: number
	rowSpan?: number
	tagName?: string
}

export interface HtmlRow {
	cells: ArrayLike<HtmlCell>
}

export interface HtmlTable {
	rows: ArrayLike<HtmlRow>
}

/** The cell styling this helper carries over to PowerPoint */
export interface CellStyle {
	color?: string
	fill?: string
	bold?: boolean
	italic?: boolean
	fontSize?: number
	fontFace?: string
	align?: 'left' | 'center' | 'right'
}

export interface TableFromHtmlProps extends PaginateTableProps {
	/**
	 * Style lookup per cell. Defaults to reading `window.getComputedStyle` when a DOM is present,
	 * and to no styling otherwise - pass this to style cells from anything else.
	 */
	styleOf?: (cell: HtmlCell) => CellStyle
	/**
	 * Treat leading rows made entirely of `<th>` as repeated headers.
	 * Ignored when `repeatHeaderRows` is set explicitly. @default true
	 */
	detectHeaderRows?: boolean
}

/** `rgb(r, g, b)` / `rgba(...)` / `#rgb` / `#rrggbb` to a bare pptx hex, or undefined */
export function cssColorToHex (value: string | null | undefined): string | undefined {
	if (!value) return undefined
	const text = value.trim()

	const rgb = /^rgba?\(\s*(\d+)[,\s]+(\d+)[,\s]+(\d+)\s*(?:[,/]\s*([\d.]+)\s*)?\)$/i.exec(text)
	if (rgb) {
		// Fully transparent is "no fill", not black
		if (rgb[4] !== undefined && Number(rgb[4]) === 0) return undefined
		return [rgb[1], rgb[2], rgb[3]].map(part => Number(part).toString(16).padStart(2, '0')).join('').toUpperCase()
	}

	const hex = /^#?([0-9a-f]{3}|[0-9a-f]{6})$/i.exec(text)
	if (!hex) return undefined
	const digits = hex[1]
	return (digits.length === 3 ? [...digits].map(d => d + d).join('') : digits).toUpperCase()
}

/** Reads the styles a browser computed, when there is a browser */
function computedStyleReader (): ((cell: HtmlCell) => CellStyle) | undefined {
	if (typeof window === 'undefined' || typeof window.getComputedStyle !== 'function') return undefined

	return function styleOf (cell: HtmlCell): CellStyle {
		// The one cast in this file: `HtmlCell` is deliberately not `Element` so the public signature
		// needs no DOM types, and the guard above already proved there is a real DOM behind it.
		const style = window.getComputedStyle(cell as unknown as Element)
		const weight = style.fontWeight
		const align = style.textAlign
		const size = Number.parseFloat(style.fontSize)
		return {
			color: cssColorToHex(style.color),
			fill: cssColorToHex(style.backgroundColor),
			bold: weight === 'bold' || Number(weight) >= 600,
			italic: style.fontStyle === 'italic',
			// CSS reports px; PowerPoint wants points
			fontSize: Number.isFinite(size) && size > 0 ? Math.round((size * 72) / 96) : undefined,
			fontFace: style.fontFamily ? style.fontFamily.split(',')[0].replace(/["']/g, '').trim() : undefined,
			align: align === 'center' || align === 'right' ? align : 'left',
		}
	}
}

function isHeaderRow (row: HtmlRow): boolean {
	const cells = Array.from({ length: row.cells.length }, (_, i) => row.cells[i])
	return cells.length > 0 && cells.every(cell => (cell.tagName ?? '').toUpperCase() === 'TH')
}

/**
 * Reproduce an HTML table as a PowerPoint table, across as many slides as it needs.
 *
 * A measured counterpart to the core `tableToSlides`: this one measures the text, and takes the
 * table as an object rather than an element id so it is not tied to a live document.
 *
 * @example
 * tableFromHtml(pres, document.getElementById('report'), { x: 0.5, y: 0.5, w: 9, fontSize: 10 })
 */
export function tableFromHtml<S extends TableSlide<PptxGenJS.TableRow>> (
	pres: TablePresentation<S, PptxGenJS.TableRow>,
	table: HtmlTable,
	props: TableFromHtmlProps = {}
): PaginateTableResult<S> {
	const { styleOf = computedStyleReader(), detectHeaderRows = true, ...paginateProps } = props

	if (!table || !table.rows || table.rows.length === 0) throw new Error('tableFromHtml: the table has no rows')

	const htmlRows = Array.from({ length: table.rows.length }, (_, i) => table.rows[i])

	const rows: PptxGenJS.TableRow[] = htmlRows.map(htmlRow => {
		const cells = Array.from({ length: htmlRow.cells.length }, (_, i) => htmlRow.cells[i])
		return cells.map((cell): PptxGenJS.TableCell => {
			const style = styleOf ? styleOf(cell) : {}
			const isHeader = (cell.tagName ?? '').toUpperCase() === 'TH'
			const options: PptxGenJS.TableCellProps = {
				...(style.color ? { color: style.color } : {}),
				...(style.fill ? { fill: { color: style.fill } } : {}),
				...(style.bold ?? isHeader ? { bold: true } : {}),
				...(style.italic ? { italic: true } : {}),
				...(style.fontSize ? { fontSize: style.fontSize } : {}),
				...(style.fontFace ? { fontFace: style.fontFace } : {}),
				...(style.align ? { align: style.align } : {}),
				...(cell.colSpan && cell.colSpan > 1 ? { colspan: cell.colSpan } : {}),
				...(cell.rowSpan && cell.rowSpan > 1 ? { rowspan: cell.rowSpan } : {}),
			}
			return { text: (cell.textContent ?? '').trim(), options }
		})
	})

	let repeatHeaderRows = paginateProps.repeatHeaderRows
	if (repeatHeaderRows === undefined && detectHeaderRows) {
		let leading = 0
		while (leading < htmlRows.length - 1 && isHeaderRow(htmlRows[leading])) leading++
		repeatHeaderRows = leading
	}

	return paginateTable(pres, rows, { ...paginateProps, repeatHeaderRows })
}
