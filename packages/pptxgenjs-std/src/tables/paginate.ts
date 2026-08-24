import type PptxGenJS from 'pptxgenjs-plus'
import { measureText } from '../text/measure'

/** Default slide size (LAYOUT_16x9), inches */
const DEFAULT_SLIDE = { w: 10, h: 5.625 }
/** The core's "Normal" cell margins in PowerPoint 2021, inches - TRBL */
const DEFAULT_CELL_MARGIN: [number, number, number, number] = [0.05, 0.1, 0.05, 0.1]
const EMU_PER_IN = 914400

/**
 * The slice of `Slide` this helper needs.
 *
 * Structural rather than `PptxGenJS.Slide` so any object with a conforming `addTable` works,
 * including a test double.
 */
export interface TableSlide<R = PptxGenJS.TableRow> {
	addTable: (rows: R[], options?: PptxGenJS.TableProps) => unknown
}

/**
 * The slice of `PptxGenJS` this helper needs: somewhere to put the slides it creates.
 *
 * The row type is a parameter so the rows come back out the way they went in, and so a slide whose
 * `addTable` is stricter than the published `TableRow` still conforms.
 */
export interface TablePresentation<S extends TableSlide<R>, R = PptxGenJS.TableRow> {
	addSlide: (options?: { masterName?: string, sectionTitle?: string }) => S
	presLayout?: { width: number, height: number }
}

export interface PaginateTableProps extends Omit<PptxGenJS.TableProps, 'h' | 'autoPage'> {
	/** Rows repeated at the top of every slide @default 0 */
	repeatHeaderRows?: number
	/** `y` for tables on the slides after the first, inches @default the initial `y` */
	continueY?: number
	/** Master slide applied to every slide this creates */
	masterName?: string
	/** Bottom of the usable area, measured from the slide bottom, inches @default 0.5 */
	bottomMargin?: number
	/** Slide height override, inches - taken from `presLayout` when available @default 5.625 */
	slideHeight?: number
}

export interface PaginateTableResult<S> {
	/** Every slide created, in order */
	slides: S[]
	/** Rows placed on each slide, header repeats included */
	rowsPerSlide: number[]
	/** Whether any measurement fell back to a guess - see `measureText`'s `source` */
	estimated: boolean
}

/** Cell text flattened to a string; nested runs concatenate, as they render on one line flow */
function cellText (cell: PptxGenJS.TableCell): string {
	if (typeof cell.text === 'string') return cell.text
	if (Array.isArray(cell.text)) return cell.text.map(cellText).join('')
	return cell.text === undefined ? '' : String(cell.text)
}

function resolveMargin (margin: PptxGenJS.TableProps['margin']): [number, number, number, number] {
	if (margin === undefined) return DEFAULT_CELL_MARGIN
	if (typeof margin === 'number') return [margin, margin, margin, margin]
	if (Array.isArray(margin) && margin.length === 4) return margin.map(Number) as [number, number, number, number]
	return DEFAULT_CELL_MARGIN
}

/** Column widths in inches: explicit `colW` if given, otherwise the table width split evenly */
function resolveColumnWidths (rows: PptxGenJS.TableRow[], tableW: number, colW: PptxGenJS.TableProps['colW']): number[] {
	const columns = rows.reduce((max, row) => Math.max(max, row.reduce((n, cell) => n + (cell.options?.colspan ?? 1), 0)), 0)
	if (columns === 0) throw new Error('paginateTable: rows have no cells')

	if (typeof colW === 'number') return Array.from({ length: columns }, () => colW)
	if (Array.isArray(colW) && colW.length > 0) {
		if (colW.length !== columns) throw new Error(`paginateTable: colW has ${colW.length} widths for ${columns} columns`)
		return colW.map(Number)
	}
	return Array.from({ length: columns }, () => tableW / columns)
}

/**
 * Add a table across as many slides as its rows need, measuring the text instead of estimating it.
 *
 * An additional pager alongside the core `autoPage`: that path guesses wraps from a per-character
 * constant (`autoPageCharWeight`). This wraps each cell with {@link measureText} - real advance
 * widths where they are known - so the row heights it adds up are the ones PowerPoint will lay out.
 *
 * Options other than the paging ones are passed straight through to `addTable`, once per slide.
 *
 * @example
 * paginateTable(pres, rows, { x: 0.5, y: 0.5, w: 9, repeatHeaderRows: 1, fontSize: 11 })
 */
export function paginateTable<R extends PptxGenJS.TableRow, S extends TableSlide<R>> (
	pres: TablePresentation<S, R>,
	rows: R[],
	props: PaginateTableProps = {}
): PaginateTableResult<S> {
	const { repeatHeaderRows = 0, continueY, masterName, bottomMargin = 0.5, slideHeight, ...tableProps } = props

	if (!Array.isArray(rows) || rows.length === 0) throw new Error('paginateTable: at least one row is required')
	if (!Number.isInteger(repeatHeaderRows) || repeatHeaderRows < 0) throw new Error(`paginateTable: repeatHeaderRows must be an integer >= 0 (got ${repeatHeaderRows})`)
	if (repeatHeaderRows >= rows.length) throw new Error(`paginateTable: repeatHeaderRows ${repeatHeaderRows} leaves no body rows`)

	const layoutH = slideHeight ?? (pres.presLayout ? pres.presLayout.height / EMU_PER_IN : DEFAULT_SLIDE.h)
	const layoutW = pres.presLayout ? pres.presLayout.width / EMU_PER_IN : DEFAULT_SLIDE.w
	const x = typeof tableProps.x === 'number' ? tableProps.x : 0.5
	const y = typeof tableProps.y === 'number' ? tableProps.y : 0.5
	const tableW = typeof tableProps.w === 'number' ? tableProps.w : layoutW - x * 2
	if (!(tableW > 0)) throw new Error(`paginateTable: table width must be > 0 (got ${tableW})`)

	const startY = continueY ?? y
	const columnWidths = resolveColumnWidths(rows, tableW, tableProps.colW)
	const [marginTop, marginRight, marginBottom, marginLeft] = resolveMargin(tableProps.margin)

	let estimated = false

	/** `rowH` is a table-level floor, per row or one value for all of them */
	const minHeightOf = (index: number): number => {
		const rowH = tableProps.rowH
		if (typeof rowH === 'number') return rowH
		if (Array.isArray(rowH)) return rowH[index] ?? 0
		return 0
	}

	/** Height one row occupies: its tallest cell, measured at that cell's column width */
	const heightOf = (row: R, index: number): number => {
		let tallest = minHeightOf(index)
		let column = 0
		for (const cell of row) {
			const span = cell.options?.colspan ?? 1
			const width = columnWidths.slice(column, column + span).reduce((a, b) => a + b, 0)
			column += span

			const usable = width - marginLeft - marginRight
			const measured = measureText(cellText(cell), {
				fontFace: cell.options?.fontFace ?? tableProps.fontFace,
				fontSize: cell.options?.fontSize ?? tableProps.fontSize ?? 12,
				bold: cell.options?.bold ?? tableProps.bold,
				italic: cell.options?.italic ?? tableProps.italic,
				// A column narrower than its own margins cannot be measured; fall back to one line
				w: usable > 0 ? usable : undefined,
			})
			if (measured.source === 'estimate') estimated = true
			tallest = Math.max(tallest, measured.h + marginTop + marginBottom)
		}
		return tallest
	}

	const header = rows.slice(0, repeatHeaderRows)
	const body = rows.slice(repeatHeaderRows)
	const headerHeight = header.reduce((total, row, index) => total + heightOf(row, index), 0)

	// Chunk body rows by measured height. The first slide starts at `y`, the rest at `continueY`.
	const chunks: R[][] = []
	let current: R[] = []
	let used = 0
	let available = layoutH - y - bottomMargin - headerHeight

	body.forEach((row, index) => {
		const height = heightOf(row, repeatHeaderRows + index)
		if (current.length > 0 && used + height > available) {
			chunks.push(current)
			current = []
			used = 0
			available = layoutH - startY - bottomMargin - headerHeight
		}
		current.push(row)
		used += height
	})
	if (current.length > 0) chunks.push(current)

	if (available <= 0 && chunks.length === 0) throw new Error(`paginateTable: no usable height on a ${layoutH}" slide below y ${y}`)

	const slides: S[] = []
	const rowsPerSlide: number[] = []
	chunks.forEach((chunk, index) => {
		const slide = pres.addSlide(masterName ? { masterName } : undefined)
		const slideRows = [...header, ...chunk]
		slide.addTable(slideRows, { ...tableProps, x, y: index === 0 ? y : startY, w: tableW, colW: columnWidths })
		slides.push(slide)
		rowsPerSlide.push(slideRows.length)
	})

	return { slides, rowsPerSlide, estimated }
}
