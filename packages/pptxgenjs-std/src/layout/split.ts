import type { GridArea } from './grid'

/**
 * Weights per slot, or a slot count for equal shares.
 * `3` and `[1, 1, 1]` mean the same thing.
 */
export type Slots = number | number[]

function resolve (slots: Slots, label: string): number[] {
	if (typeof slots === 'number') {
		if (!Number.isInteger(slots) || slots < 1) throw new Error(`${label}: slot count must be an integer >= 1 (got ${slots})`)
		return Array.from({ length: slots }, () => 1)
	}
	if (slots.length === 0) throw new Error(`${label}: weights must not be empty`)
	if (slots.some(weight => !(weight > 0))) throw new Error(`${label}: every weight must be > 0 (got [${slots.join(', ')}])`)
	return slots
}

function split (area: GridArea, slots: Slots, gap: number, axis: 'w' | 'h', label: string): GridArea[] {
	const weights = resolve(slots, label)
	if (gap < 0) throw new Error(`${label}: gap must be >= 0 (got ${gap})`)

	const total = area[axis] - gap * (weights.length - 1)
	if (!(total > 0)) throw new Error(`${label}: gap ${gap} leaves no room across ${area[axis]}`)

	const sum = weights.reduce((a, b) => a + b, 0)
	const offset = axis === 'w' ? 'x' : 'y'
	let at = area[offset]

	return weights.map(weight => {
		const size = (total * weight) / sum
		const slot = { ...area, [offset]: at, [axis]: size } as GridArea
		at += size + gap
		return slot
	})
}

/**
 * Divide an area into side-by-side slots.
 *
 * Takes the placement arithmetic off the caller: pass a count or weights, get areas ready to
 * spread into any `addX` options. Output is the same shape as the input, so slots nest.
 *
 * @example
 * const [left, right] = row(at(0, 0, 12, 6), [1, 2])
 * const [a, b, c] = row(left, 3, 0.1)
 */
export function row (area: GridArea, slots: Slots, gap = 0.2): GridArea[] {
	return split(area, slots, gap, 'w', 'row')
}

/**
 * Divide an area into stacked slots.
 *
 * @example
 * const [header, body] = column(at(0, 0, 12, 6), [1, 4])
 */
export function column (area: GridArea, slots: Slots, gap = 0.2): GridArea[] {
	return split(area, slots, gap, 'h', 'column')
}
