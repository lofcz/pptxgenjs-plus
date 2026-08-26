/**
 * pptxgenjsx — JSX runtime for building PowerPoint presentations with pptxgenjs.
 *
 * Entry point that re-exports all public API (components + context).
 *
 * @module pptxgenjsx
 */

export { Fragment } from "./node.js";
export type {
  ComponentFactory,
  ComponentProps,
  ContainerProps,
  LeafProps,
  PptxNode,
  PptxChild,
  PptxChildren,
  PptxNodePromise,
  PrimitiveChild,
} from "./node.js";
export { useSlideContext, useDeckContext, useGroupContext } from "./context.js";
export type { SlideContextInfo, DeckContextInfo, GroupContextInfo } from "./context.js";
export * from "./components.js";
export {
  applyOmmlRunStyle,
  fromMarkdown,
  latexToOmml,
  mathRun,
  tokenizeMath,
  tryLatexToOmml,
} from "./markdown/index.js";
export type { ContentSegment, FromMarkdownStyle, OmmlRunStyle } from "./markdown/index.js";
