/** Centimetres per inch - the whole reason this file exists */
const CM_PER_IN = 2.54

/** Points per inch. Font sizes are already in points; positions are not. */
const PT_PER_IN = 72

function convert (value: number, per: number, label: string): number {
	if (!Number.isFinite(value)) throw new Error(`${label}: value must be a finite number (got ${value})`)
	return value / per
}

/**
 * Centimetres as inches.
 *
 * Every `addX` option is in inches, so a design specified in centimetres - which is most of
 * them outside the US - gets divided by 2.54 at every call site until someone gets it wrong.
 *
 * @example
 * slide.addText('Titel', { x: cm(2.5), y: cm(1.8), w: cm(20), h: cm(2) })
 * const at = grid({ w: cm(33.87), h: cm(19.05), margin: cm(1.27) })
 */
export function cm (value: number): number {
	return convert(value, CM_PER_IN, 'cm')
}

/**
 * Points as inches.
 *
 * For lengths given in points - line widths, optical offsets, anything sized to match a font -
 * where the option itself wants inches.
 *
 * @example
 * slide.addShape('line', { x: 1, y: pt(18), w: 4, h: 0 })
 */
export function pt (value: number): number {
	return convert(value, PT_PER_IN, 'pt')
}
