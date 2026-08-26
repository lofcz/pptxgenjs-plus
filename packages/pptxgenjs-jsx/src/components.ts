/**
 * Component factory definitions for pptxgenjsx.
 *
 * Each component factory creates a tagged PptxNode that the renderer later
 * maps to the corresponding PptxGenJS method.
 *
 * All positioning / sizing values are in inches.
 * Default canvas: 13.333 x 7.5 (WIDE).
 *
 * @module components
 */

import type PptxGenJSType from "pptxgenjs-plus";

import { fromMarkdown } from "./markdown/fromMarkdown.js";
import {
  type ComponentFactory,
  type ComponentProps,
  type PptxChildren,
  type PptxNode,
  type LeafProps,
  createNode,
} from "./node.js";

// ════════════════════════════════════════════════════════════════════
// COMPONENT PROPS TYPE DEFINITIONS
// ════════════════════════════════════════════════════════════════════

// ── Presentation ──────────────────────────────────────────────────

/** Props for the root `<Deck>` / `<Presentation>` component.
 *
 * Maps presentation-level metadata, layout, themes, and slide masters
 * to the underlying `new PptxGenJS()` constructor and `pptx.defineLayout`,
 * `pptx.defineSlideMaster`, etc.
 */
export type DeckProps = {
  /** Presentation title metadata. */
  title?: string;
  /** Presentation author metadata. */
  author?: string;
  /** Presentation company metadata. */
  company?: string;
  /** Presentation subject metadata. */
  subject?: string;
  /** Presentation revision metadata. Must be a whole-number string for PowerPoint compatibility. */
  revision?: string;
  /** Enable right-to-left mode for the whole presentation. */
  rtlMode?: boolean;
  /** OPC core `dcterms:created`. Maps to `pptx.created`. */
  created?: Date;
  /** OPC core `dcterms:modified`. Maps to `pptx.modified`. */
  modified?: Date;
  /** Failed media handling. Maps to `pptx.mediaOnError`. */
  mediaOnError?: PptxGenJSType.MediaOnError;
  /** Default theme fonts. Maps to `pptx.theme`. */
  theme?: PptxGenJSType.ThemeProps;
  /** Built-in layout name or custom layout object. Use `layouts` for additional custom layouts. */
  layout?: string | PptxGenJSType.PresLayout;
  /** Additional custom layouts passed to `pptx.defineLayout`. */
  layouts?: PptxGenJSType.PresLayout[];
  /** Master slides passed to `pptx.defineSlideMaster`. */
  masters?: PptxGenJSType.SlideMasterProps[];
  /** Presentation sections passed to `pptx.addSection`. */
  sections?: PptxGenJSType.SectionProps[];
  /** Child nodes — typically `<Slide />`, `<Layout />`, `<Master />`, `<Section />`. */
  children?: PptxChildren;
};

// ── Structure: Slide, Layout, Section, Master, Placeholder ───────

/** A lazy component import — analogous to React Router's `component={() => import('./path')}`.
 * The function returns a Promise resolving to a module whose `default` export is a
 * {@link ComponentFactory}. The component is loaded and rendered during `createPptx()`.
 */
export type ComponentImport<P extends ComponentProps = ComponentProps> = () => Promise<{
  default: ComponentFactory<P>;
}>;

/** Props for `<Slide>`. Maps to `pptx.addSlide()`.
 * Supports background, visibility, slide number format, layout selection,
 * and a lazy `component` prop for code-split slide definitions.
 */
export type SlideProps = Partial<
  Pick<PptxGenJSType.PresSlide, "background" | "color" | "hidden" | "slideNumber">
> &
  PptxGenJSType.AddSlideProps & {
    /** Lazy-loaded component factory. Accepts `() => import('./path')` — the loaded
     * module's `default` export is called as a component factory during rendering.
     * When provided, the component result replaces direct children for slide content.
     */
    component?: ComponentImport;
    children?: PptxChildren;
  };

/** Props for `<Layout>`. Maps to `pptx.defineLayout()`. */
export type LayoutProps = PptxGenJSType.PresLayout & LeafProps;

/** Props for `<Section>`. Maps to `pptx.addSection()`. */
export type SectionProps = PptxGenJSType.SectionProps & {
  children?: PptxChildren;
};

/** Props for `<Master>`. Maps to `pptx.defineSlideMaster()`. */
export type MasterProps = PptxGenJSType.SlideMasterProps & {
  children?: PptxChildren;
};

/** Props for `<Placeholder>` (inside `<Master>`).
 * Maps to `SlideMasterProps.objects[].placeholder` entries.
 */
export type PlaceholderProps = {
  /** Placeholder options used inside `defineSlideMaster({ objects })`. */
  options: PptxGenJSType.PlaceholderProps;
  /** Placeholder text shown until edited. Falls back to children strings. */
  text?: string;
  children?: PptxChildren;
};

// ── Text & Notes ─────────────────────────────────────────────────

/** Props for `<TextRun>` — a single formatted run inside `<Text>`.
 * Maps to an entry in the `TextProps[]` array passed to `slide.addText()`.
 *
 * The run's text comes from the `text` prop or, when absent, from the
 * element's string/number children: `<TextRun options={{ bold: true }}>Bold</TextRun>`.
 */
export type TextRunProps = PptxGenJSType.TextProps & {
  /** Run text. Falls back to string/number children. */
  children?: string | number | Array<string | number>;
};

/** Props for `<Text>` — a text box, rich text container, or text inside a shape.
 * Maps to `slide.addText(text, options)`.
 *
 * Text can come from one of four sources. When more than one is present,
 * the following priority applies (higher wins, the rest is ignored):
 *
 * 1. child `<TextRun />` nodes — plain string/number children mixed among
 *    them are converted to default-style runs (order preserved)
 * 2. the `runs` prop (rich text array)
 * 3. the `text` prop
 * 4. plain-string children
 *
 * Plain strings mix freely with `<TextRun />` children (each becomes a
 * default-style run in place; whitespace-only text between elements is
 * ignored). Mixing the other modes drops the lower-priority content.
 */
export type TextProps = Omit<PptxGenJSType.TextPropsOptions, "children"> & {
  /** Text string or rich-text array passed to `slide.addText`. Falls back to children strings. */
  text?: string | PptxGenJSType.TextProps[];
  /** Rich text runs passed to `slide.addText`. Same shape as PptxGenJS `TextProps[]`. */
  runs?: PptxGenJSType.TextProps[];
  /** Optional nested `<TextRun />` nodes or plain strings. */
  children?: PptxChildren;
  /** Options forwarded to the underlying PptxGenJS API call. */
  options?: PptxGenJSType.TextPropsOptions;
};

/**
 * A fixed-size text box. Markdown (headings, lists, bold/italic, `$…$` / `$$…$$`)
 * is parsed into native PPTX runs — real bullets and OMML equations — then
 * PowerPoint shrink-to-fit keeps the authored `x/y/w/h`.
 */
export type FixedBoxProps = Omit<TextProps, "text" | "runs" | "children"> & {
  x: PptxGenJSType.Coord;
  y: PptxGenJSType.Coord;
  w: PptxGenJSType.Coord;
  h: PptxGenJSType.Coord;
  /** CommonMark + TeX delimiters. Preferred over string children. */
  markdown?: string;
  /** Markdown source when `markdown` is omitted. */
  children?: string | number | readonly (string | number | null | undefined)[];
};

/** Props for `<Notes>` — speaker notes for a slide.
 * Maps to `slide.addNotes()`.
 */
export type NotesProps = {
  /** Speaker notes text. Falls back to children strings. */
  text?: string;
  children?: PptxChildren;
};

// ── Shapes ────────────────────────────────────────────────────────

/** Generic shape props for `<Shape>`. Use this for any PptxGenJS shape
 * that does not have a dedicated component (e.g. `<Rect />`, `<Ellipse />`).
 * Maps to `slide.addShape(shape, options)`.
 */
export type ShapeProps = Omit<PptxGenJSType.ShapeProps, "children"> & {
  /** Shape type passed as first argument to `slide.addShape()`. e.g. `"rect"`, `"ellipse"`, `"roundRect"`, `"line"`. */
  shape: PptxGenJSType.SHAPE_NAME;
  /** Options forwarded to the underlying PptxGenJS API call. */
  options?: PptxGenJSType.ShapeProps;
} & LeafProps;

/** Props for `<Line>` — a simple line shape.
 * Maps to `slide.addShape("line", options)`. For lines defined by two
 * endpoints, use `<LineBetween>` instead.
 */
export type LineProps = Omit<ShapeProps, "shape">;

/** Props for `<LineBetween>` — a line connecting two absolute coordinates.
 * Pass start/end points via `x1`/`y1`/`x2`/`y2`.
 */
export type LineBetweenProps = Omit<LineProps, "x" | "y" | "w" | "h" | "flipH" | "flipV"> & {
  /** Start x coordinate. Supports inches (number) or percentage of group/parent width (`${number}%`). */
  x1: PptxGenJSType.Coord;
  /** Start y coordinate. Supports inches (number) or percentage of group/parent height (`${number}%`). */
  y1: PptxGenJSType.Coord;
  /** End x coordinate. Supports inches (number) or percentage of group/parent width (`${number}%`). */
  x2: PptxGenJSType.Coord;
  /** End y coordinate. Supports inches (number) or percentage of group/parent height (`${number}%`). */
  y2: PptxGenJSType.Coord;
};

/** Shared shape-options base for all named shape components. */
export type ShapeOptionsProps = Omit<PptxGenJSType.ShapeProps, "children"> & {
  /** Options forwarded to the underlying PptxGenJS API call. */
  options?: PptxGenJSType.ShapeProps;
} & LeafProps;

/** Props for `<Rect>`. Maps to `slide.addShape("rect", options)`. */
export type RectProps = ShapeOptionsProps;
/** Props for `<Ellipse>` / `<Oval>`. Maps to `slide.addShape("ellipse", options)`. */
export type EllipseProps = ShapeOptionsProps;
/** Props for `<Triangle>`. Maps to `slide.addShape("triangle", options)`. */
export type TriangleProps = ShapeOptionsProps;
/** Props for `<Diamond>`. Maps to `slide.addShape("diamond", options)`. */
export type DiamondProps = ShapeOptionsProps;
/** Props for `<Pentagon>`. Maps to `slide.addShape("pentagon", options)`. */
export type PentagonProps = ShapeOptionsProps;
/** Props for `<Hexagon>`. Maps to `slide.addShape("hexagon", options)`. */
export type HexagonProps = ShapeOptionsProps;
/** Props for star shapes (`<Star>`, `<Star4>`–`<Star10>`). Maps to `slide.addShape("star*", options)`. */
export type StarProps = ShapeOptionsProps;
/** Props for arrow shapes (`<LeftArrow>`, `<RightArrow>`, etc.). Maps to `slide.addShape("*Arrow", options)`. */
export type ArrowProps = ShapeOptionsProps;

/** Props for `<RoundRect>`. Maps to `slide.addShape("roundRect", options)`.
 * @property rectRadius — corner radius; valid only for `roundRect`, range 0.0 (square) to 1.0 (fully rounded).
 */
export type RoundRectProps = ShapeOptionsProps & {
  /** Rounded rectangle radius. Valid only for `roundRect`; range 0.0 to 1.0. */
  rectRadius?: number;
};

/** Props for `<Arc>`. Maps to `slide.addShape("arc", options)`.
 * @property angleRange — start and sweep angles in degrees [start, sweep]; valid for `arc`, `pie`, `blockArc`.
 */
export type ArcProps = ShapeOptionsProps & {
  /** Arc angle range. Valid for `arc`, `pie`, and `blockArc`; range [0-359, 0-359]. */
  angleRange?: [number, number];
};

/** Props for `<BlockArc>`. Maps to `slide.addShape("blockArc", options)`.
 * Extends `<Arc>` with a thickness ratio.
 */
export type BlockArcProps = ArcProps & {
  /** Block arc thickness ratio. Valid only for `blockArc`; range 0.0 to 1.0. */
  arcThicknessRatio?: number;
};

/** Props for `<PieShape>` (visual pie-slice shape, NOT a data chart).
 * Maps to `slide.addShape("pie", options)`.
 */
export type PieShapeProps = ArcProps;

/** A single point or segment in a custom geometry path.
 * Supports move-to, line-to, close, and three curve types (arc, cubic, quadratic).
 */
export type CustomGeometryPoint =
  | { x: PptxGenJSType.Coord; y: PptxGenJSType.Coord; moveTo?: boolean }
  | {
      x: PptxGenJSType.Coord;
      y: PptxGenJSType.Coord;
      curve: {
        type: "arc";
        hR: PptxGenJSType.Coord;
        wR: PptxGenJSType.Coord;
        stAng: number;
        swAng: number;
      };
    }
  | {
      x: PptxGenJSType.Coord;
      y: PptxGenJSType.Coord;
      curve: {
        type: "cubic";
        x1: PptxGenJSType.Coord;
        y1: PptxGenJSType.Coord;
        x2: PptxGenJSType.Coord;
        y2: PptxGenJSType.Coord;
      };
    }
  | {
      x: PptxGenJSType.Coord;
      y: PptxGenJSType.Coord;
      curve: { type: "quadratic"; x1: PptxGenJSType.Coord; y1: PptxGenJSType.Coord };
    }
  | { close: true };

/** Props for `<CustomGeometry>`. Maps to `slide.addShape("custGeom", options)`.
 * Convert SVG path commands to points:
 * | SVG | CustomGeometryPoint |
 * |-----|---------------------|
 * | `M x y` | `{ x, y, moveTo:true }` |
 * | `L x y` | `{ x, y }` |
 * | `C x1 y1, x2 y2, x y` | cubic curve entry |
 * | `Q x1 y1, x y` | quadratic curve entry |
 * | `A ...` | arc curve entry |
 * | `Z` / `z` | `{ close:true }` |
 */
export type CustomGeometryProps = ShapeOptionsProps & {
  /** Custom geometry path points passed to PptxGenJS `custGeom`. */
  points: CustomGeometryPoint[];
};

// ── Media ─────────────────────────────────────────────────────────

/** Props for `<Image>`. Maps to `slide.addImage(options)`.
 * Accepts a file path, data URI, or base64 string via the `path` prop.
 */
export type ImageProps = PptxGenJSType.ImageProps & {
  /** Options forwarded to the underlying PptxGenJS API call. */
  options?: PptxGenJSType.ImageProps;
} & LeafProps;

/** Props for `<Media>` — audio, video, or online media.
 * Maps to `slide.addMedia(options)`.
 */
export type MediaProps = PptxGenJSType.MediaProps & {
  /** Options forwarded to the underlying PptxGenJS API call. */
  options?: PptxGenJSType.MediaProps;
} & LeafProps;

// ── Charts ────────────────────────────────────────────────────────

/** Props for `<Chart>` — generic chart with explicit `type` + `data`.
 * Maps to `slide.addChart(type, data, options)`.
 * Use typed chart components (`<BarChart />`, `<LineChart />`, etc.)
 * for a more concise syntax.
 */
export type ChartProps = Omit<PptxGenJSType.IChartOpts, "children"> & {
  /** Chart type or multi-chart descriptor. */
  type: PptxGenJSType.CHART_NAME | PptxGenJSType.IChartMulti[];
  /** Chart series data. */
  data: PptxGenJSType.OptsChartData[];
  /** Options forwarded to the underlying PptxGenJS API call. */
  options?: PptxGenJSType.IChartOpts;
} & LeafProps;

/** Props shared by typed chart components. */
export type TypedChartProps = Omit<ChartProps, "type">;

/** Props for `<AreaChart>`. Maps to `slide.addChart("area", data, options)`. */
export type AreaChartProps = TypedChartProps;
/** Props for `<BarChart>`. Maps to `slide.addChart("bar", data, options)`. */
export type BarChartProps = TypedChartProps;
/** Props for `<Bar3DChart>`. Maps to `slide.addChart("bar3D", data, options)`. */
export type Bar3DChartProps = TypedChartProps;
/** Props for `<BubbleChart>`. Maps to `slide.addChart("bubble", data, options)`. */
export type BubbleChartProps = TypedChartProps;
/** Props for `<DoughnutChart>`. Maps to `slide.addChart("doughnut", data, options)`. */
export type DoughnutChartProps = TypedChartProps;
/** Props for `<LineChart>`. Maps to `slide.addChart("line", data, options)`. */
export type LineChartProps = TypedChartProps;
/** Props for `<PieChart>`. Maps to `slide.addChart("pie", data, options)`. */
export type PieChartProps = TypedChartProps;
/** Props for `<RadarChart>`. Maps to `slide.addChart("radar", data, options)`. */
export type RadarChartProps = TypedChartProps;
/** Props for `<ScatterChart>`. Maps to `slide.addChart("scatter", data, options)`. */
export type ScatterChartProps = TypedChartProps;

// ── Table ─────────────────────────────────────────────────────────

/** Props for `<TableCell>` — a single cell in a table row.
 * Maps to an entry in the `PptxGenJS.TableCell[]` array.
 */
export type TableCellProps = PptxGenJSType.TableCellProps & {
  /** Cell text. Falls back to children strings. */
  text?: string | PptxGenJSType.TableCell[];
  children?: PptxChildren;
  /** Options forwarded to the underlying PptxGenJS API call. */
  options?: PptxGenJSType.TableCellProps;
};

/** Props for `<TableRow>` — a single row in a table.
 * Maps to `PptxGenJS.TableRow`.
 */
export type TableRowProps = {
  /** Row data. Falls back to child `<TableCell />` nodes. */
  cells?: PptxGenJSType.TableRow;
  children?: PptxChildren;
};

/** Props for `<Table>`. Maps to `slide.addTable(rows, options)`.
 * Rows can be provided via the `rows` prop or as child `<TableRow />` nodes.
 */
export type TableProps = Omit<PptxGenJSType.TableProps, "children"> & {
  /** Table rows passed as first argument to `slide.addTable`. */
  rows?: PptxGenJSType.TableRow[];
  children?: PptxChildren;
  /** Options forwarded to the underlying PptxGenJS API call. */
  options?: PptxGenJSType.TableProps;
};

/** Props for `<TableToSlides>` — split an HTML table across multiple slides.
 * Maps to `pptx.tableToSlides(eleId, options)` (browser / runtime DOM only).
 */
export type TableToSlidesProps = PptxGenJSType.TableToSlidesProps & {
  /** HTML table element id passed to `pptx.tableToSlides`. */
  eleId: string;
  /** Options forwarded to the underlying PptxGenJS API call. */
  options?: PptxGenJSType.TableToSlidesProps;
} & LeafProps;

// ── Group (`p:grpSp`) ──────────────────────────────────────────────

/** Props for `<Group>`. Maps to `slide.addGroup()`.
 * Child coordinates are relative to the group origin. Only shapes, text,
 * images, and nested groups are valid children (not tables, charts, or media).
 */
export type GroupProps = Omit<PptxGenJSType.GroupProps, "_objects"> & {
  children?: PptxChildren;
};

/** Props for `<WordArt>`. Maps to `slide.addWordArt()`. */
export type WordArtProps = TextProps;

// ── Escape hatch ──────────────────────────────────────────────────

/** Props for `<Raw>` — escape hatch for unsupported or newly added PptxGenJS APIs.
 * Provides direct access to the `pptx` instance and the current `slide`.
 */
export type RawProps = {
  /** Custom render callback invoked during rendering. Supports sync or async. */
  render: (context: RenderContext) => void | Promise<void>;
  /** Raw children are NOT auto-rendered. They are only accessible via
   * `context.node.children` inside the `render` callback — an escape
   * hatch for fully custom content. */
  children?: PptxChildren;
};

/** Context passed to the `<Raw>` render callback. */
export type RenderContext = {
  /** The root PptxGenJS instance. */
  pptx: PptxGenJSType;
  /** The current slide being rendered (undefined during presentation-level rendering). */
  slide?: PptxGenJSType.PresSlide;
  /** The current group builder when rendering inside `<Group>`. */
  group?: PptxGenJSType.Group;
  /** The raw node being rendered. */
  node: PptxNode<"Raw">;
};

// ════════════════════════════════════════════════════════════════════
// COMPONENT FACTORIES
// ════════════════════════════════════════════════════════════════════

// ── Generic factory helpers ───────────────────────────────────────

/** Create a generic component factory for the given node type tag. */
function component<T extends string>(type: T): ComponentFactory {
  return (props) => createNode(type, props as any);
}

/** Create a component factory for a specific PptxGenJS shape (e.g. `"rect"`, `"ellipse"`). */
function shapeComponent<T extends string>(type: T, shape: string): ComponentFactory {
  return (props) => createNode(type, { ...props, shape } as any);
}

/** Create a component factory for a specific chart type (e.g. `"bar"`, `"line"`). */
function chartComponent<T extends string>(type: T, chartType: string): ComponentFactory {
  return (props) => createNode(type, { ...props, type: chartType } as any);
}

// ── Deck / Presentation ──────────────────────────────────────────

/** Root presentation wrapper. Maps to `new PptxGenJS()`.
 *
 * Usage:
 * ```tsx
 * <Deck title="My Deck">
 *   <Slide>...</Slide>
 * </Deck>
 * ```
 */
export const Deck = component("Deck") as ComponentFactory<DeckProps>;
/** Alias for `<Deck />`. */
export const Presentation = Deck as ComponentFactory<DeckProps>;

// ── Slide ─────────────────────────────────────────────────────────

/** A single slide in the presentation. Maps to `pptx.addSlide()`.
 *
 * Usage:
 * ```tsx
 * <Slide layout="TITLE_ONLY">
 *   <Text>Hello</Text>
 * </Slide>
 * ```
 */
export const Slide = component("Slide") as ComponentFactory<SlideProps>;

// ── Layout / Section / Master / Placeholder ───────────────────────

/** Define an additional custom slide layout. Maps to `pptx.defineLayout()`.
 * Provide the layout object via props (name, width, height).
 */
export const Layout = component("Layout") as ComponentFactory<LayoutProps>;
/** Group one or more slides into a named PowerPoint section.
 * Maps to `pptx.addSection()`.
 *
 * Usage:
 * ```tsx
 * <Section name="Overview">
 *   <Slide>...</Slide>
 * </Section>
 * ```
 */
export const Section = component("Section") as ComponentFactory<SectionProps>;
/** Define a reusable slide master. Maps to `pptx.defineSlideMaster()`.
 * Supports placeholders via nested `<Placeholder />` nodes.
 */
export const Master = component("Master") as ComponentFactory<MasterProps>;
/** Declare a placeholder inside a slide master.
 * Maps to `SlideMasterProps.objects[].placeholder` entries.
 */
export const Placeholder = component("Placeholder") as ComponentFactory<PlaceholderProps>;

// ── Text ──────────────────────────────────────────────────────────

/** Text box or rich text block. Maps to `slide.addText(text, options)`.
 *
 * Usage:
 * ```tsx
 * <Text x={1} y={1} w={4} fontSize={24}>Hello World</Text>
 * <Text x={1} y={2} w={4}>
 *   <TextRun bold fontSize={24}>Bold</TextRun>
 *   <TextRun color="888">Regular</TextRun>
 * </Text>
 * ```
 *
 * Text content can come from the `text` prop, child `<TextRun />` nodes,
 * plain-string children, or the `runs` prop. */
export const Text = component("Text") as ComponentFactory<TextProps>;
/** A single formatted text run inside `<Text />`. Maps to `PptxGenJS.TextProps`.
 * Supports per-run font, color, bold, italic, etc.
 */
export const TextRun = component("TextRun") as ComponentFactory<TextRunProps>;
/** Speaker notes for a slide. Maps to `slide.addNotes()`.
 *
 * Usage:
 * ```tsx
 * <Slide>
 *   <Notes>Remember to mention the quarterly results.</Notes>
 * </Slide>
 * ```
 */
export const Notes = component("Notes") as ComponentFactory<NotesProps>;

function markdownFromProps(props: FixedBoxProps): string {
  if (typeof props.markdown === "string") return props.markdown;
  const children = props.children;
  if (typeof children === "string" || typeof children === "number") return String(children);
  if (Array.isArray(children)) {
    return children
      .filter((child): child is string | number => typeof child === "string" || typeof child === "number")
      .join("");
  }
  return "";
}

/**
 * Fixed box with autoscaled markdown. Lists become PPTX bullets; `$a^2$` and
 * `$$\\frac{1}{2}$$` become native OMML. Defaults to `fit="shrink"` and `valign="top"`.
 *
 * ```tsx
 * <FixedBox x={0.7} y={1.5} w={6} h={4} markdown={`
 * # Title
 * - item one
 * - item with $a^2+b^2=c^2$
 *
 * $$\\frac{1}{2}$$
 * `} />
 * ```
 */
export const FixedBox = ((props: FixedBoxProps) => {
  const { markdown: _markdown, children: _children, fit, valign, fontSize, color, fontFace, ...rest } =
    props;
  const colorHex = typeof color === "string" ? color : undefined;
  return createNode("Text", {
    ...rest,
    fontSize,
    color,
    fontFace,
    fit: fit ?? "shrink",
    valign: valign ?? "top",
    runs: fromMarkdown(markdownFromProps(props), { fontSize, color: colorHex, fontFace }),
  });
}) as ComponentFactory<FixedBoxProps>;

// ── Shapes ────────────────────────────────────────────────────────

/** Generic shape — for any PptxGenJS shape without a dedicated component.
 * Maps to `slide.addShape(shape, options)`. */
export const Shape = component("Shape") as ComponentFactory<ShapeProps>;

/** Rectangle. Maps to `slide.addShape("rect", options)`. */
export const Rect = shapeComponent("Rect", "rect") as ComponentFactory<RectProps>;
/** Rounded rectangle; supports `rectRadius` for corner rounding.
 * Maps to `slide.addShape("roundRect", options)`. */
export const RoundRect = shapeComponent(
  "RoundRect",
  "roundRect",
) as ComponentFactory<RoundRectProps>;
/** Ellipse / oval. Maps to `slide.addShape("ellipse", options)`. */
export const Ellipse = shapeComponent("Ellipse", "ellipse") as ComponentFactory<EllipseProps>;
/** Alias for `<Ellipse />`. */
export const Oval = Ellipse;
/** Triangle. Maps to `slide.addShape("triangle", options)`. */
export const Triangle = shapeComponent("Triangle", "triangle") as ComponentFactory<TriangleProps>;
/** Right-angle triangle. Maps to `slide.addShape("rtTriangle", options)`. */
export const RightTriangle = shapeComponent(
  "RightTriangle",
  "rtTriangle",
) as ComponentFactory<TriangleProps>;
/** Diamond. Maps to `slide.addShape("diamond", options)`. */
export const Diamond = shapeComponent("Diamond", "diamond") as ComponentFactory<DiamondProps>;
/** Pentagon. Maps to `slide.addShape("pentagon", options)`. */
export const Pentagon = shapeComponent(
  "Pentagon",
  "pentagon",
) as ComponentFactory<ShapeOptionsProps>;
/** Hexagon. Maps to `slide.addShape("hexagon", options)`. */
export const Hexagon = shapeComponent("Hexagon", "hexagon") as ComponentFactory<HexagonProps>;
/** Five-point star. Maps to `slide.addShape("star5", options)`. */
export const Star = shapeComponent("Star", "star5") as ComponentFactory<StarProps>;
/** Four-point star. Maps to `slide.addShape("star4", options)`. */
export const Star4 = shapeComponent("Star4", "star4") as ComponentFactory<StarProps>;
/** Alias for `<Star />` (five-point). */
export const Star5 = Star;
/** Six-point star. Maps to `slide.addShape("star6", options)`. */
export const Star6 = shapeComponent("Star6", "star6") as ComponentFactory<StarProps>;
/** Eight-point star. Maps to `slide.addShape("star8", options)`. */
export const Star8 = shapeComponent("Star8", "star8") as ComponentFactory<StarProps>;
/** Ten-point star. Maps to `slide.addShape("star10", options)`. */
export const Star10 = shapeComponent("Star10", "star10") as ComponentFactory<StarProps>;
/** Simple line. Maps to `slide.addShape("line", options)`.
 * For lines defined by start/end points, use `<LineBetween />` instead. */
export const Line = component("Line") as ComponentFactory<LineProps>;
/** Line connecting two absolute coordinates.
 * Pass start/end points via `x1`/`y1`/`x2`/`y2`. */
export const LineBetween = component("LineBetween") as ComponentFactory<LineBetweenProps>;
/** Arc shape; supports `angleRange`. Maps to `slide.addShape("arc", options)`. */
export const Arc = shapeComponent("Arc", "arc") as ComponentFactory<ArcProps>;
/** Thick arc; supports `angleRange` and `arcThicknessRatio`.
 * Maps to `slide.addShape("blockArc", options)`. */
export const BlockArc = shapeComponent("BlockArc", "blockArc") as ComponentFactory<BlockArcProps>;
/** Pie-slice shape (visual only — not a data chart).
 * Maps to `slide.addShape("pie", options)`. */
export const PieShape = shapeComponent("PieShape", "pie") as ComponentFactory<PieShapeProps>;
/** Editable custom path geometry. Maps to `slide.addShape("custGeom", options)`. */
export const CustomGeometry = shapeComponent(
  "CustomGeometry",
  "custGeom",
) as ComponentFactory<CustomGeometryProps>;
/** Left-pointing arrow. Maps to `slide.addShape("leftArrow", options)`. */
export const LeftArrow = shapeComponent("LeftArrow", "leftArrow") as ComponentFactory<ArrowProps>;
/** Right-pointing arrow. Maps to `slide.addShape("rightArrow", options)`. */
export const RightArrow = shapeComponent(
  "RightArrow",
  "rightArrow",
) as ComponentFactory<ArrowProps>;
/** Up-pointing arrow. Maps to `slide.addShape("upArrow", options)`. */
export const UpArrow = shapeComponent("UpArrow", "upArrow") as ComponentFactory<ArrowProps>;
/** Down-pointing arrow. Maps to `slide.addShape("downArrow", options)`. */
export const DownArrow = shapeComponent("DownArrow", "downArrow") as ComponentFactory<ArrowProps>;
/** Left-right bidirectional arrow. Maps to `slide.addShape("leftRightArrow", options)`. */
export const LeftRightArrow = shapeComponent(
  "LeftRightArrow",
  "leftRightArrow",
) as ComponentFactory<ArrowProps>;
/** Up-down bidirectional arrow. Maps to `slide.addShape("upDownArrow", options)`. */
export const UpDownArrow = shapeComponent(
  "UpDownArrow",
  "upDownArrow",
) as ComponentFactory<ArrowProps>;
/** Chevron / angle bracket. Maps to `slide.addShape("chevron", options)`. */
export const Chevron = shapeComponent("Chevron", "chevron") as ComponentFactory<ArrowProps>;
/** Cloud shape. Maps to `slide.addShape("cloud", options)`. */
export const Cloud = shapeComponent("Cloud", "cloud") as ComponentFactory<ShapeOptionsProps>;
/** Heart shape. Maps to `slide.addShape("heart", options)`. */
export const Heart = shapeComponent("Heart", "heart") as ComponentFactory<ShapeOptionsProps>;
/** Donut / ring shape. Maps to `slide.addShape("donut", options)`. */
export const Donut = shapeComponent("Donut", "donut") as ComponentFactory<ShapeOptionsProps>;
/** Plus / cross shape. Maps to `slide.addShape("plus", options)`. */
export const Plus = shapeComponent("Plus", "plus") as ComponentFactory<ShapeOptionsProps>;

// ── Media ─────────────────────────────────────────────────────────

/** Image (path, data URI, or base64). Maps to `slide.addImage(options)`. */
export const Image = component("Image") as ComponentFactory<ImageProps>;
/** Audio, video, or online media. Maps to `slide.addMedia(options)`. */
export const Media = component("Media") as ComponentFactory<MediaProps>;

// ── Charts ────────────────────────────────────────────────────────

/** Generic chart with explicit type + data. Maps to `slide.addChart(type, data, options)`.
 * For common chart types, prefer typed components (`<BarChart />`, `<LineChart />`, etc.). */
export const Chart = component("Chart") as ComponentFactory<ChartProps>;
/** Area chart. Maps to `slide.addChart("area", data, options)`. */
export const AreaChart = chartComponent("AreaChart", "area") as ComponentFactory<AreaChartProps>;
/** Bar chart. Maps to `slide.addChart("bar", data, options)`. */
export const BarChart = chartComponent("BarChart", "bar") as ComponentFactory<BarChartProps>;
/** 3D bar chart. Maps to `slide.addChart("bar3D", data, options)`. */
export const Bar3DChart = chartComponent(
  "Bar3DChart",
  "bar3D",
) as ComponentFactory<Bar3DChartProps>;
/** Bubble chart. Maps to `slide.addChart("bubble", data, options)`. */
export const BubbleChart = chartComponent(
  "BubbleChart",
  "bubble",
) as ComponentFactory<BubbleChartProps>;
/** Doughnut chart. Maps to `slide.addChart("doughnut", data, options)`. */
export const DoughnutChart = chartComponent(
  "DoughnutChart",
  "doughnut",
) as ComponentFactory<DoughnutChartProps>;
/** Line chart. Maps to `slide.addChart("line", data, options)`. */
export const LineChart = chartComponent("LineChart", "line") as ComponentFactory<LineChartProps>;
/** Pie chart. Maps to `slide.addChart("pie", data, options)`. */
export const PieChart = chartComponent("PieChart", "pie") as ComponentFactory<PieChartProps>;
/** Radar chart. Maps to `slide.addChart("radar", data, options)`. */
export const RadarChart = chartComponent(
  "RadarChart",
  "radar",
) as ComponentFactory<RadarChartProps>;
/** Scatter chart. Maps to `slide.addChart("scatter", data, options)`. */
export const ScatterChart = chartComponent(
  "ScatterChart",
  "scatter",
) as ComponentFactory<ScatterChartProps>;

// ── Table ─────────────────────────────────────────────────────────

/** Table of rows and cells. Maps to `slide.addTable(rows, options)`. */
export const Table = component("Table") as ComponentFactory<TableProps>;
/** A row inside `<Table />`. Maps to `PptxGenJS.TableRow`. */
export const TableRow = component("TableRow") as ComponentFactory<TableRowProps>;
/** A cell inside `<TableRow />`. Maps to `PptxGenJS.TableCell`. */
export const TableCell = component("TableCell") as ComponentFactory<TableCellProps>;
/** Split an HTML table across multiple slides (browser runtime only).
 * Maps to `pptx.tableToSlides(eleId, options)`. */
export const TableToSlides = component("TableToSlides") as ComponentFactory<TableToSlidesProps>;

// ── Escape hatch ──────────────────────────────────────────────────

/** Escape hatch for unsupported or new PptxGenJS APIs.
 * Provides direct access via a custom render callback.
 * Maps to the `<Raw>` custom render protocol. */
export const Raw = component("Raw") as ComponentFactory<RawProps>;

// ── Group ─────────────────────────────────────────────────────────

/** PresentationML group (`p:grpSp`). Child coordinates are relative to
 * the group's origin. Maps to `slide.addGroup()`. Nested groups use
 * `group.addGroup()`. Tables, charts, and media are not valid children. */
export const Group = component("Group") as ComponentFactory<GroupProps>;

/** WordArt text with a warp (`a:prstTxWarp`). Maps to `slide.addWordArt()`. */
export const WordArt = component("WordArt") as ComponentFactory<WordArtProps>;
