# pptxgenjs-plus-jsx

A JSX runtime for building PowerPoint presentations with [pptxgenjs-plus](https://github.com/lofcz/pptxgenjs-plus). Write your slides as JSX components and render them to `.pptx` files.

```tsx
import { Deck, Slide, Text, TextRun, Rect } from "pptxgenjs-plus-jsx";
import { renderPptx } from "pptxgenjs-plus-jsx/render";

await renderPptx(
  <Deck title="My Deck">
    <Slide>
      <Rect x={0} y={0} w={13.333} h={7.5} fill={{ color: "1E1E2E" }} />
      <Text x={1} y={3} w={11} h={1.5}>
        <TextRun options={{ fontSize: 44, color: "FFFFFF", bold: true }}>
          Hello, PowerPoint!
        </TextRun>
      </Text>
    </Slide>
  </Deck>,
  { fileName: "output.pptx" },
);
```

## Table of Contents

- [Installation](#installation)
- [TypeScript Configuration](#typescript-configuration)
- [Component Reference](#component-reference)
  - [Deck & Presentation](#deck--presentation)
  - [Slide](#slide)
  - [Text & TextRun](#text--textrun)
  - [FixedBox](#fixedbox)
  - [Shapes](#shapes)
  - [Charts](#charts)
  - [Tables](#tables)
  - [Images & Media](#images--media)
  - [Group](#group)
  - [Raw (escape hatch)](#raw-escape-hatch)
  - [Fragment](#fragment)
- [Async Components](#async-components)
- [Context Hooks](#context-hooks)
  - [useSlideContext](#useslidecontext)
  - [useDeckContext](#usedeckcontext)
  - [useGroupContext](#usegroupcontext)
- [Lazy Slide Loading](#lazy-slide-loading)
- [Percentage Coordinates](#percentage-coordinates)
- [Sections & Masters](#sections--masters)
- [Validation](#validation)
- [Rendering](#rendering)
- [License](#license)

---

## Installation

```bash
npm install pptxgenjs-plus-jsx
```

Peer dependency:

- **[pptxgenjs-plus](https://github.com/lofcz/pptxgenjs-plus)** — workspace dependency in this monorepo.

## TypeScript Configuration

Set `jsx` and `jsxImportSource` in your `tsconfig.json`:

```json
{
  "compilerOptions": {
    "jsx": "react-jsx",
    "jsxImportSource": "pptxgenjs-plus-jsx"
  }
}
```

> **Note**: This is **not** React — the JSX transform produces `PptxNode` objects, not DOM elements. No React dependency is required.

---

## Component Reference

### Deck / Presentation

The root element of every presentation. Maps to `new PptxGenJS()`.

```tsx
<Deck>
  <Slide>...</Slide>
</Deck>
```

`<Presentation>` is an alias for `<Deck>`.

#### Layout

Use the `layout` prop to control slide dimensions. Two forms are supported:

**1. Built-in layout name** (string) — pptxgenjs provides four standard presets:

| Name             | Dimensions    | Aspect Ratio |
| ---------------- | ------------- | ------------ |
| `"LAYOUT_4x3"`   | 10" × 7.5"    | 4:3          |
| `"LAYOUT_16x9"`  | 10" × 5.625"  | 16:9         |
| `"LAYOUT_16x10"` | 10" × 6.25"   | 16:10        |
| `"LAYOUT_WIDE"`  | 13.33" × 7.5" | 16:9 (wide)  |

Default: `"LAYOUT_WIDE"` (13.33" × 7.5").

```tsx
<Deck layout="LAYOUT_16x9">
  <Slide>...</Slide>
</Deck>
```

**2. Custom layout** (object) — define arbitrary dimensions via a `PresLayout` object with `name`, `width`, and `height` (in inches):

```tsx
<Deck layout={{ name: "A4", width: 10.83, height: 7.82 }}>
  <Slide>...</Slide>
</Deck>
```

For multiple custom layouts, use the `layouts` prop (array of `PresLayout`):

```tsx
<Deck
  layout="A4"
  layouts={[
    { name: "A4", width: 10.83, height: 7.82 },
    { name: "Letter", width: 10, height: 7.5 },
  ]}
>
  <Slide>...</Slide>
</Deck>
```

The first matching layout name becomes the presentation's active layout.

### Slide

A single slide. Maps to `pptx.addSlide()`.

```tsx
<Slide>
  <Text>A simple slide</Text>
</Slide>
```

**Lazy-loaded slide** (see [Lazy Slide Loading](#lazy-slide-loading)):

```tsx
<Slide component={() => import("./slides/chart-slide")} />
```

### Text & TextRun

**`<Text>`** — a text box or rich text container. Maps to `slide.addText()`.

```tsx
// Simple text (string from children)
<Text x={1} y={1} w={8} h={1} fontSize={24} color="333333">
  Hello, World!
</Text>

// Rich text with multiple TextRun elements
<Text x={1} y={2.5} w={8} h={1.5} valign="middle">
  <TextRun options={{ fontSize: 18, color: "666666" }}>Normal text </TextRun>
  <TextRun options={{ fontSize: 18, color: "0066CC", bold: true }}>
    bold and blue
  </TextRun>
</Text>
```

**`<TextRun>`** — a single formatted run inside `<Text>`. `options` accepts pptxgenjs `TextProps` (fontSize, color, bold, italic, fontFace, etc.). The run's text comes from its string children (or the `text` prop):

```tsx
<TextRun options={{ fontSize: 18, color: "0066CC", bold: true }}>bold and blue</TextRun>
```

**Text input modes.** A `<Text>` element has four ways to receive content. When more than one is present, this priority applies (higher wins, the rest is ignored):

1. child `<TextRun />` nodes — plain string/number children mixed among them are converted to default-style runs (order preserved)
2. the `runs` prop (rich text array — same shape as pptxgenjs `TextProps[]`)
3. the `text` prop
4. plain-string children

Plain strings mix freely with `<TextRun />` children — each string segment becomes a default-style run in place:

```tsx
<Text x={1} y={4} w={8} h={1}>
  {"Normal segment "}
  <TextRun options={{ color: "0066CC", bold: true }}>highlighted</TextRun>
</Text>
```

Whitespace-only text between elements (e.g. from multi-line JSX formatting) is ignored — attach explicit spaces to a run's text (`<TextRun>A </TextRun>`) instead. Mixing the other modes (e.g. `runs` or `text` props alongside children) still drops the lower-priority content.

```tsx
// runs prop (imperative rich text)
<Text
  x={1}
  y={1}
  w={8}
  h={1}
  runs={[
    { text: "Normal ", options: { fontSize: 18 } },
    { text: "bold", options: { fontSize: 18, bold: true } },
  ]}
/>
```

**Text with a shape background.** `shape` is a valid pptxgenjs text option and is forwarded to `slide.addText()`:

```tsx
<Text
  x={1}
  y={1}
  w={4}
  h={2}
  shape="roundRect"
  fill={{ color: "EDE9FE" }}
  margin={18} // text margin uses POINTS (~0.25"), not inches
  valign="middle"
>
  Shaped text box
</Text>
```

> Note: pptxgenjs text `margin` values are in **points** (e.g. `18` ≈ 0.25"), not inches.

### FixedBox

A **fixed-size** text box. Pass CommonMark plus TeX delimiters; the box stays at the authored `x/y/w/h` and PowerPoint shrink-to-fit scales the content (`fit="shrink"`).

- Headings (`#`–`######` and setext), paragraphs, `**bold**` / `*italic*`, `~~strike~~`, `---` rules, `>` quotes, `[links](url)`, and lists become native PPTX runs (real bullets / numbered lists).
- `$…$` (inline) and `$$…$$` (display) — also `\(…\)`, `\[…\]`, `\begin{env}` — become editable OMML via MathLive + mathml2omml-plus. No extra helper, no image raster.

```tsx
<FixedBox
  x={0.7}
  y={1.5}
  w={6}
  h={4}
  fontSize={18}
  color="0F172A"
  markdown={`
# Title
- item one
- item with $a^2+b^2=c^2$

$$\\frac{1}{2}$$
`}
/>
```

`fromMarkdown`, `tokenizeMath`, `latexToOmml`, and `mathRun` are also exported from the package root when you need runs without the component.

### Shapes

All pptxgenjs shapes are available as JSX components. Each supports standard positioning props (`x`, `y`, `w`, `h`) plus shape-specific options via `options` or as top-level props.

> **Shapes are leaf elements** — they do not render children. Nested content
> (e.g. `<RoundRect><Text>…</Text></RoundRect>`) is a compile-time error in TSX
> and a hard runtime error otherwise; it is never silently dropped.
>
> To put text on a shape, use **one** of:
>
> 1. A single text box with a shape background: `<Text shape="roundRect" …>` (recommended when the text and shape share one box — see [Text & TextRun](#text--textrun)).
> 2. A sibling `<Text>` layered over the shape (when text and shape need independent styles/shadows/geometry).
> 3. A `<Group>` only when you need a relative coordinate system or want to move/scale the pair together.

```tsx
<Rect x={0} y={0} w={13.333} h={7.5} fill={{ color: "1E1E2E" }} />
<Ellipse x={1} y={1} w={3} h={2} fill={{ color: "FF6B6B" }} />
<Triangle x={5} y={1} w={2} h={2} fill={{ color: "4ECDC4" }} />
<RoundRect x={1} y={4} w={4} h={2} fill={{ color: "45B7D1" }} rectRadius={0.3} />
<Line x={1} y={1} w={5} h={0} line={{ color: "FF0000", width: 2 }} />
<Cloud x={1} y={1} w={3} h={2} fill={{ color: "E8F4FD" }} />
<Heart x={5} y={1} w={2} h={2} fill={{ color: "FF4081" }} />
```

**`<LineBetween>`** — a line connecting two absolute coordinates. Supports percentage coordinates.

```tsx
<LineBetween
  x1={0}
  y1={0}
  x2={13.333}
  y2={7.5}
  line={{ color: "999999", width: 1, dashType: "dash" }}
/>
```

Available shape components:

| Component                  | pptxgenjs shape                     |
| -------------------------- | ----------------------------------- |
| `Rect`                     | `rect`                              |
| `RoundRect`                | `roundRect`                         |
| `Ellipse` / `Oval`         | `ellipse`                           |
| `Triangle`                 | `triangle`                          |
| `RightTriangle`            | `rtTriangle`                        |
| `Diamond`                  | `diamond`                           |
| `Pentagon`                 | `pentagon`                          |
| `Hexagon`                  | `hexagon`                           |
| `Star` / `Star5`           | `star5`                             |
| `Star4`                    | `star4`                             |
| `Star6`                    | `star6`                             |
| `Star8`                    | `star8`                             |
| `Star10`                   | `star10`                            |
| `Line`                     | `line`                              |
| `LineBetween`              | `line` (with computed bounding box) |
| `Arc`                      | `arc`                               |
| `BlockArc`                 | `blockArc`                          |
| `PieShape`                 | `pie`                               |
| `CustomGeometry`           | `custGeom`                          |
| `LeftArrow` / `RightArrow` | `leftArrow` / `rightArrow`          |
| `UpArrow` / `DownArrow`    | `upArrow` / `downArrow`             |
| `LeftRightArrow`           | `leftRightArrow`                    |
| `UpDownArrow`              | `upDownArrow`                       |
| `Chevron`                  | `chevron`                           |
| `Cloud`                    | `cloud`                             |
| `Heart`                    | `heart`                             |
| `Donut`                    | `donut`                             |
| `Plus`                     | `plus`                              |

### Charts

```tsx
// Generic chart with explicit `type`
<Chart
  x={1} y={1} w={10} h={5}
  type="bar"
  data={[
    { name: "Q1", labels: ["Jan", "Feb", "Mar"], values: [100, 150, 200] },
    { name: "Q2", labels: ["Jan", "Feb", "Mar"], values: [120, 180, 160] },
  ]}
  showTitle="Quarterly Sales"
  showLegend={true}
/>

// Or use a typed chart component
<BarChart x={1} y={1} w={10} h={5} data={[...]} showValue={true} />
<LineChart x={1} y={1} w={10} h={5} data={[...]} lineSize={3} />
<PieChart x={1} y={1} w={6} h={5} data={[...]} showPercent={true} />
```

Available chart components: `AreaChart`, `BarChart`, `Bar3DChart`, `BubbleChart`, `DoughnutChart`, `LineChart`, `PieChart`, `RadarChart`, `ScatterChart`.

### Tables

```tsx
<Table x={1} y={1} w={10} h={3} fontSize={12} border={{ type: "solid", color: "CCCCCC" }}>
  <TableRow>
    <TableCell options={{ fill: { color: "4472C4" }, color: "FFFFFF", bold: true }}>Name</TableCell>
    <TableCell options={{ fill: { color: "4472C4" }, color: "FFFFFF", bold: true }}>
      Value
    </TableCell>
  </TableRow>
  <TableRow>
    <TableCell>Item A</TableCell>
    <TableCell>100</TableCell>
  </TableRow>
  <TableRow>
    <TableCell>Item B</TableCell>
    <TableCell>200</TableCell>
  </TableRow>
</Table>
```

`TableToSlides` splits an HTML table across multiple slides (browser runtime only).

### Images & Media

```tsx
<Image
  x={1} y={1} w={5} h={3}
  path="https://example.com/image.png"
  sizing={{ type: "contain", w: 5, h: 3 }}
/>

<Media
  x={1} y={1} w={6} h={4}
  path="https://example.com/video.mp4"
  mediaType="video"
/>
```

### Group

A PresentationML group (`p:grpSp`). Child coordinates are relative to the group's origin. Maps to `slide.addGroup()`.

```tsx
<Group x={1} y={1} w={10} h={5}>
  {/* (0, 0) inside group → group-relative; the group itself sits at (1, 1) */}
  <Rect x={0} y={0} w={10} h={5} fill={{ color: "F0F0F0" }} />
  {/* "50%" inside group → 5" from the group origin */}
  <Text x="50%" y="50%" w={4} h={1}>
    <TextRun options={{ fontSize: 18 }}>Centered in group</TextRun>
  </Text>
</Group>
```

Key features:

- **Native group shape**: children move and transform together in PowerPoint.
- **Coordinate transformation**: Child `x`, `y`, `w`, `h` are relative to the group origin. Percentage strings resolve against the group's `w` (for x/w) or `h` (for y/h).
- **Nested groups**: Groups can nest via `group.addGroup()`.
- **Valid children**: shapes, text, images, and nested groups. Tables, charts, media, notes, and WordArt must stay on the slide.
- **Context-aware**: Children can use `useGroupContext()` for the group's virtual canvas (`width` / `height`) and slide origin (`x` / `y`).

### Raw (escape hatch)

For pptxgenjs features not covered by a dedicated component.

```tsx
<Raw
  render={({ pptx, slide, node }) => {
    // Direct access to slide.addShape(), slide.addText(), etc.
    slide.addShape("rect", { x: 1, y: 1, w: 5, h: 3, fill: { color: "FF0000" } });
  }}
/>
```

> Raw children are **not** auto-rendered — they are only accessible via
> `context.node.children` inside the `render` callback. Use props for
> configuration; use children only when the callback reads them itself.

### Fragment

Groups multiple children without producing a wrapper element. Useful when a component needs to return multiple siblings (e.g., in a lazy-loaded slide or inside a map).

**Shorthand form `<>...</>`** — for simple grouping without props:

```tsx
// slides/title-slide.tsx
export default function TitleSlide() {
  return (
    <>
      <Text x={1} y={3} w={10} h={1.5} fontSize={44} bold>
        Welcome
      </Text>
      <Text x={1} y={4.5} w={10} h={1} fontSize={18} color="666666">
        Subtitle text
      </Text>
    </>
  );
}
```

**Explicit `<Fragment>`** — when you need a `key` prop (e.g., in a `.map()` loop):

```tsx
import { Fragment } from "pptxgenjs-plus-jsx";

<Slide>
  {items.map((item) => (
    <Fragment key={item.id}>
      <Text x={1} y={item.y}>
        {item.name}
      </Text>
      <Text x={5} y={item.y}>
        {item.value}
      </Text>
    </Fragment>
  ))}
</Slide>;
```

> **Note**: `<>...</>` is a JSX compile-time syntax — it does not support props like `key`. For dynamic lists, always use `<Fragment key={...}>`.

---

## Async Components

Components can be `async` functions — they are automatically detected and lazily resolved during rendering:

```tsx
// slides/data-slide.tsx
export default async function DataSlide() {
  const res = await fetch("https://api.example.com/data");
  const data = await res.json();

  return (
    <Slide>
      <Text x={1} y={1} w={8} h={1} fontSize={32} bold>
        {data.title}
      </Text>
      <Text x={1} y={2.5} w={8} h={4} fontSize={16}>
        {data.description}
      </Text>
    </Slide>
  );
}
```

```tsx
// main.tsx
await renderPptx(
  <Deck>
    <Slide>{/* ... */}</Slide>
    <DataSlide /> {/* async — resolves automatically */}
  </Deck>,
  { fileName: "output.pptx" },
);
```

This works because the JSX factory (`jsx`) wraps async component results in a `PptxNodePromise`, and the renderer resolves them during tree traversal.

---

## Context Hooks

Context hooks provide runtime information about the current rendering environment.

**How it works**: When TypeScript compiles your JSX, component factories are NOT called during JSX construction. Instead, they are wrapped in a deferred node and executed later during rendering — at which point the renderer has set up the context store via `AsyncLocalStorage`. This is why you can write components that call `useSlideContext()` as direct children of `<Slide>`, even though the JSX appears to be built "eagerly."

### useSlideContext

Exposes the current slide's index and total.

```tsx
import { useSlideContext } from "pptxgenjs-plus-jsx";

function SlideNumber() {
  const { index, total, sectionTitle } = useSlideContext();

  return (
    <Text x={1} y={6.5} w={10} h={0.5} fontSize={10} color="999999">
      Slide {index} of {total}
      {sectionTitle ? ` · ${sectionTitle}` : ""}
    </Text>
  );
}
```

```tsx
// Usage inside a Slide:
<Slide>
  {/* ... slide content ... */}
  <SlideNumber />
</Slide>
```

### useDeckContext

Exposes the deck's slide dimensions (width, height in inches).

```tsx
import { useDeckContext } from "pptxgenjs-plus-jsx";

function FullBleedBackground() {
  const { width, height } = useDeckContext();
  return <Rect x={0} y={0} w={width} h={height} fill={{ color: "1E1E2E" }} />;
}
```

### useGroupContext

Exposes the current group's origin on the slide and its virtual canvas. When called outside a `<Group>`, falls back to deck dimensions with zero offset (`relative: false`). Inside a group, child coordinates are group-relative (`relative: true`).

```tsx
import { useGroupContext } from "pptxgenjs-plus-jsx";

function ProgressBar() {
  const { width } = useGroupContext();
  return <Rect x={0} y={0} w={width * 0.7} h={0.4} fill={{ color: "4CAF50" }} />;
}
```

---

## Lazy Slide Loading

Use the `component` prop on `<Slide>` to defer loading of slide definitions — analogous to React Router's lazy route loading.

```tsx
// slides/title-slide.tsx
export default function TitleSlide() {
  return (
    <>
      <Text x={1} y={3} w={10} h={1.5} fontSize={44} bold>
        Welcome
      </Text>
    </>
  );
}
```

```tsx
// main.tsx
<Slide component={() => import("./slides/title-slide")} />
```

The component's return value is rendered **inside** the `<Slide>` that declares `component`, so it should provide slide **content only** — not another `<Slide>` element. Context hooks work inside lazy-loaded components.

---

## Percentage Coordinates

`x`, `y`, `w`, `h` values can be specified as percentage strings (e.g. `"50%"`, `"100%"`), which are resolved relative to the enclosing context:

- **Inside a `<Group>`**: resolved against the group's `w` (for x/w) or `h` (for y/h).
- **Directly inside a `<Slide>`**: resolved against the slide's dimensions from the deck layout.

```tsx
<Group x={1} y={1} w={10} h={5}>
  {/* 50% of group width (5"), 25% of group height (1.25") */}
  <Rect x="25%" y="25%" w="50%" h="50%" fill={{ color: "4ECDC4" }} />
</Group>
```

This works for all positioning props: `x`, `y`, `w`, `h` on all shape/text/image components, and `x1`, `y1`, `x2`, `y2` on `LineBetween`.

---

## Sections & Masters

### Sections

Group slides into named sections in the PowerPoint outline view.

```tsx
<Deck>
  <Slide>...</Slide>
  <Section title="Overview">
    <Slide>...</Slide>
    <Slide>...</Slide>
  </Section>
  <Section title="Details">
    <Slide>...</Slide>
  </Section>
</Deck>
```

### Masters

Define slide masters with reusable layout objects.

```tsx
<Deck>
  <Master name="myMaster" background={{ fill: "F5F5F5" }}>
    <Text x={1} y={0.3} w={10} h={0.5} fontSize={10} color="999999">
      Confidential
    </Text>
    <Rect x={0} y={7} w={13.333} h={0.5} fill={{ color: "4472C4" }} />
  </Master>
  <Slide masterName="myMaster">...</Slide>
</Deck>
```

You can also use `<Placeholder>` inside masters:

```tsx
<Master name="content">
  <Placeholder options={{ name: "Body", type: "body", x: 1, y: 1, w: 10, h: 5 }} />
</Master>
```

## Validation

The `validateDeck()` function checks your slide tree for common mistakes before rendering. It is **async** — always `await` it:

```tsx
import { validateDeck } from "pptxgenjs-plus-jsx/render";

const deck = <Deck>{/* ... */}</Deck>;
const issues = await validateDeck(deck);

if (issues.length > 0) {
  for (const issue of issues) {
    console.error(`[${issue.level}] ${issue.message}`);
  }
}
```

Validation catches:

- Invalid child types (e.g., a `<Text>` directly inside a `<Deck>`)
- **Leaf components with children** (`child.leaf`, e.g. `<RoundRect>` containing a `<Text>`)
- Missing required props (e.g., `CustomGeometry` without `points`)
- Suspicious prop usage (e.g., `angleRange` on `RoundRect`)
- Mixed `<Text>` input modes (`text.input.mixed`) — warns whenever a lower-priority source would be ignored and recommends `<TextRun />` children for rich text
- Invalid `LineBetween` endpoints

---

## Rendering

### To a File

```tsx
import { renderPptx } from "pptxgenjs-plus-jsx/render";

await renderPptx(<Deck>{/* ... */}</Deck>, {
  fileName: "output/presentation.pptx",
});
```

Additional pptxgenjs `writeFile` options are also supported:

```tsx
await renderPptx(<Deck>{/* ... */}</Deck>, {
  fileName: "output.pptx",
  compression: true, // Enable ZIP compression
  zipOptions: { level: 9 }, // Compression level
});
```

### To a Buffer / Blob

```tsx
import { writePptx } from "pptxgenjs-plus-jsx/render";

const buffer = await writePptx(<Deck>{/* ... */}</Deck>, {
  outputType: "arraybuffer", // "arraybuffer" | "blob" | "uint8array" | "base64" | "nodebuffer"
});
```

### Reuse an Existing PptxGenJS Instance

```tsx
import PptxGenJS from "pptxgenjs-plus";

const pptx = new PptxGenJS();
// ... configure the instance ...
await renderPptx(<Deck>{/* ... */}</Deck>, { pptx });
```

### Aliases

```tsx
import { render, write } from "pptxgenjs-plus-jsx/render";

await render(<Deck>{/* ... */}</Deck>, { fileName: "output.pptx" });
```

## License

MIT
