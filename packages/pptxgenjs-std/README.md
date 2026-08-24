# pptxgenjs-plus-std

Layout, text-measure, table-paging, and diagram helpers for [pptxgenjs-plus](https://github.com/lofcz/pptxgenjs-plus). These helpers compose the public `addChart` / `addTable` / position APIs — they do not replace the core library.

```ts
import pptxgen from 'pptxgenjs-plus'
import { grid, waterfall, measureText, paginateTable, cm } from 'pptxgenjs-plus-std'

const pptx = new pptxgen()
const slide = pptx.addSlide()
const at = grid()

slide.addText('Q3 bridge', { ...at(0, 0, 12, 1), fontSize: 28 })
waterfall(slide, {
  labels: ['Open', 'Sales', 'Churn'],
  values: [100, 30, -12],
  total: 'Close',
}, { ...at(0, 1, 12, 5), showValue: true })
```

## Helpers

- `grid()` / `gridFor(pres)` — place objects on a column/row grid; throws instead of drifting off-slide
- `row(area, slots)` / `column(area, slots)` — split any `{ x, y, w, h }` by count or weights
- `cm()` / `pt()` — centimetres and points as the inches every `addX` option takes
- `measureText` / `fitText` / `checkOverflow` / `registerFontMetrics` — wrap and size text from canvas, bundled Calibri (Carlito) widths, or a flat estimate
- `paginateTable` / `tableFromHtml` — page tables by measured cell height (an additional pager alongside core `autoPage` / `tableToSlides`)
- `waterfall(slide, props, options?)` — stacked-bar waterfall (transparent riser + signed per-point labels)

Category subpaths: `pptxgenjs-plus-std/layout`, `pptxgenjs-plus-std/charts`, `pptxgenjs-plus-std/text`, `pptxgenjs-plus-std/tables`.

`pptxgenjs-plus` is a peer dependency at the same version as this package, so helpers act on the caller's own presentation instance.
