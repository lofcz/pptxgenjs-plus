/**
 * CommonMark → pptxgenjs text runs.
 *
 * Math is lifted out first ({@link tokenizeMath}) so `$…$` / `$$…$$` become
 * native OMML runs. Lists become real PPTX bullets so importers (Fika) can
 * measure `<li>` and shrink-to-fit the authored box.
 */

import MarkdownIt from "markdown-it";
import type PptxGenJS from "pptxgenjs-plus";
import { tryLatexToOmml } from "./latexToOmml.js";
import { tokenizeMath } from "./tokenizeMath.js";

export type FromMarkdownStyle = {
  fontSize?: number;
  color?: string;
  fontFace?: string;
};

type TextRun = PptxGenJS.TextProps;

type MdToken = {
  type: string;
  tag: string;
  content: string;
  children: MdToken[] | null;
};

const PLACEHOLDER_START = "\uE000";
const PLACEHOLDER_END = "\uE001";
const PLACEHOLDER_RE = /\uE000(\d+)\uE001/g;
const HEADING_SCALE = [1.8, 1.45, 1.2, 1.1, 1.05, 1] as const;
const DEFAULT_FONT_SIZE = 18;

const markdownParser = new MarkdownIt({
  html: false,
  breaks: true,
  linkify: false,
  typographer: false,
});

function headingSize(level: number, body: number): number {
  const scale = HEADING_SCALE[Math.min(5, Math.max(0, level - 1))] ?? 1;
  return Math.round(body * scale);
}

function placeholderFor(index: number): string {
  return `${PLACEHOLDER_START}${index}${PLACEHOLDER_END}`;
}

export function fromMarkdown(source: string, style: FromMarkdownStyle = {}): TextRun[] {
  const fontSize = style.fontSize && style.fontSize > 0 ? style.fontSize : DEFAULT_FONT_SIZE;
  const color = style.color;
  const fontFace = style.fontFace;
  const math: { latex: string; display: boolean }[] = [];
  let rebuilt = "";
  for (const segment of tokenizeMath(source ?? "")) {
    if (segment.type === "math") {
      rebuilt += placeholderFor(math.length);
      math.push({ latex: segment.value.trim(), display: segment.display });
    } else {
      rebuilt += segment.value;
    }
  }

  const tokens = markdownParser.parse(rebuilt, {}) as MdToken[];
  const runs: TextRun[] = [];
  let pendingBreak = false;
  const listStack: { ordered: boolean }[] = [];
  let headingLevel: number | null = null;
  let inListItem = false;
  let itemFirstRun = false;
  let boldDepth = 0;
  let italicDepth = 0;

  const baseOptions = (): PptxGenJS.TextPropsOptions => {
    const options: PptxGenJS.TextPropsOptions = {};
    if (fontFace) options.fontFace = fontFace;
    if (color) options.color = color;
    options.fontSize = headingLevel ? headingSize(headingLevel, fontSize) : fontSize;
    if (boldDepth > 0 || headingLevel) options.bold = true;
    if (italicDepth > 0) options.italic = true;
    return options;
  };

  const attachList = (options: PptxGenJS.TextPropsOptions): PptxGenJS.TextPropsOptions => {
    if (!inListItem || !itemFirstRun) return options;
    const list = listStack[listStack.length - 1];
    if (!list) return options;
    itemFirstRun = false;
    options.bullet = list.ordered ? { type: "number" } : true;
    if (listStack.length > 1) options.indentLevel = listStack.length - 1;
    return options;
  };

  const flushBreak = () => {
    if (!pendingBreak || runs.length === 0) {
      pendingBreak = false;
      return;
    }
    const last = runs[runs.length - 1];
    last.options = { ...last.options, breakLine: true };
    pendingBreak = false;
  };

  const pushRun = (text: string, extra: PptxGenJS.TextPropsOptions = {}) => {
    if (!text) return;
    flushBreak();
    const options = attachList({ ...baseOptions(), ...extra });
    runs.push({ text, options });
  };

  const pushMath = (entry: { latex: string; display: boolean }) => {
    const size = headingLevel ? headingSize(headingLevel, fontSize) : fontSize;
    const omml = tryLatexToOmml(entry.latex, { fontSizePt: size, color });
    if (!omml) {
      pushRun(entry.latex, { italic: true });
      return;
    }
    if (entry.display) pendingBreak = true;
    flushBreak();
    const options = attachList({
      ...baseOptions(),
      fontSize: size,
      omml,
    });
    runs.push({ text: "", options });
    if (entry.display) pendingBreak = true;
  };

  const emitText = (text: string, extra: PptxGenJS.TextPropsOptions = {}) => {
    if (!text) return;
    let last = 0;
    PLACEHOLDER_RE.lastIndex = 0;
    for (const match of text.matchAll(PLACEHOLDER_RE)) {
      const index = match.index ?? 0;
      if (index > last) pushRun(text.slice(last, index), extra);
      const entry = math[Number(match[1])];
      if (entry) pushMath(entry);
      last = index + match[0].length;
    }
    if (last < text.length) pushRun(text.slice(last), extra);
  };

  const emitInline = (children: MdToken[] | null) => {
    if (!children) return;
    for (const child of children) {
      switch (child.type) {
        case "text":
          emitText(child.content);
          break;
        case "softbreak":
        case "hardbreak":
          pushRun(" ");
          break;
        case "strong_open":
        case "b_open":
          boldDepth += 1;
          break;
        case "strong_close":
        case "b_close":
          boldDepth = Math.max(0, boldDepth - 1);
          break;
        case "em_open":
        case "i_open":
          italicDepth += 1;
          break;
        case "em_close":
        case "i_close":
          italicDepth = Math.max(0, italicDepth - 1);
          break;
        case "code_inline":
          emitText(child.content, { fontFace: "Consolas" });
          break;
        case "image":
          if (child.content) emitText(child.content);
          break;
        default:
          if (child.children) emitInline(child.children);
          if (child.type === "text" || child.content) break;
          break;
      }
    }
  };

  const endBlock = () => {
    if (runs.length > 0) pendingBreak = true;
  };

  for (const token of tokens) {
    switch (token.type) {
      case "heading_open":
        headingLevel = Number(token.tag.slice(1)) || 1;
        break;
      case "heading_close":
        endBlock();
        headingLevel = null;
        break;
      case "paragraph_close":
        endBlock();
        break;
      case "bullet_list_open":
        listStack.push({ ordered: false });
        break;
      case "ordered_list_open":
        listStack.push({ ordered: true });
        break;
      case "bullet_list_close":
      case "ordered_list_close":
        listStack.pop();
        break;
      case "list_item_open":
        inListItem = true;
        itemFirstRun = true;
        break;
      case "list_item_close":
        endBlock();
        inListItem = false;
        itemFirstRun = false;
        break;
      case "inline":
        emitInline(token.children);
        break;
      case "fence":
      case "code_block": {
        const body = token.content.replace(/\n$/, "");
        if (body) {
          for (const line of body.split("\n")) {
            pushRun(line || " ", { fontFace: "Consolas" });
            pendingBreak = true;
          }
        }
        break;
      }
      default:
        break;
    }
  }

  return runs;
}
