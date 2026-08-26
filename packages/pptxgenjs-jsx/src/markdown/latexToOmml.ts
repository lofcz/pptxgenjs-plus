/**
 * LaTeX → MathML (MathLive SSR) → OMML (mathml2omml-plus).
 * No hand-rolled math grammar — conversion is entirely library-based.
 */

import { convertLatexToMathMl } from "@lofcz/mathlive/ssr";
import { mml2omml } from "mathml2omml-plus";
import type PptxGenJS from "pptxgenjs-plus";

const MATH_NS = "http://www.w3.org/1998/Math/MathML";

export interface OmmlRunStyle {
  color?: string;
  fontSizePt?: number;
}

function wrapMathMl(fragment: string): string {
  const trimmed = fragment.trim();
  if (!trimmed) throw new Error("empty MathML from MathLive");
  if (/^<math[\s/>]/i.test(trimmed)) return trimmed;
  return `<math xmlns="${MATH_NS}">${trimmed}</math>`;
}

function mathMlToOmml(mathmlFragment: string): string {
  const omml = mml2omml(wrapMathMl(mathmlFragment));
  if (!omml || !/<m:oMath[\s/>]/i.test(omml)) {
    throw new Error("mathml2omml produced no m:oMath");
  }
  return omml;
}

function toSrgbHex(color: string): string | null {
  const raw = color.trim();
  if (/^[0-9a-fA-F]{6}$/.test(raw)) return raw.toUpperCase();
  const hashed = raw.startsWith("#") ? raw.slice(1) : raw;
  if (/^[0-9a-fA-F]{6}$/.test(hashed)) return hashed.toUpperCase();
  if (/^[0-9a-fA-F]{3}$/.test(hashed)) {
    return hashed
      .split("")
      .map((ch) => ch + ch)
      .join("")
      .toUpperCase();
  }
  return null;
}

function buildArPr(style: OmmlRunStyle): string | null {
  const parts: string[] = [];
  const attrs: string[] = ['dirty="0"'];
  if (style.fontSizePt && Number.isFinite(style.fontSizePt) && style.fontSizePt > 0) {
    attrs.push(`sz="${Math.round(style.fontSizePt * 100)}"`);
  }
  if (style.color) {
    const hex = toSrgbHex(style.color);
    if (hex) parts.push(`<a:solidFill><a:srgbClr val="${hex}"/></a:solidFill>`);
  }
  if (!parts.length && attrs.length === 1) return null;
  return `<a:rPr ${attrs.join(" ")}>${parts.join("")}</a:rPr>`;
}

/** Stamp color / size onto every `m:r` and `m:ctrlPr` so PPTX contrast matches the box. */
export function applyOmmlRunStyle(omml: string, style: OmmlRunStyle): string {
  const rPr = buildArPr(style);
  if (!rPr || !omml) return omml;
  const withRunPr = omml.replace(/<m:r(\s[^>]*)?>([\s\S]*?)<\/m:r>/g, (_full, attrs = "", body: string) => {
    const next = /<a:rPr[\s/>]/.test(body)
      ? body.replace(/<a:rPr\b[^>]*\/>|<a:rPr\b[\s\S]*?<\/a:rPr>/, rPr)
      : /<m:rPr\b[\s\S]*?<\/m:rPr>/.test(body)
        ? body.replace(/<\/m:rPr>/, `</m:rPr>${rPr}`)
        : `${rPr}${body}`;
    return `<m:r${attrs}>${next}</m:r>`;
  });

  return withRunPr
    .replace(/<m:ctrlPr(\s[^>]*)?\/>/g, `<m:ctrlPr$1>${rPr}</m:ctrlPr>`)
    .replace(/<m:ctrlPr(\s[^>]*)?>([\s\S]*?)<\/m:ctrlPr>/g, (_full, attrs = "", body: string) => {
      const next = /<a:rPr[\s/>]/.test(body)
        ? body.replace(/<a:rPr\b[^>]*\/>|<a:rPr\b[\s\S]*?<\/a:rPr>/, rPr)
        : `${rPr}${body}`;
      return `<m:ctrlPr${attrs}>${next}</m:ctrlPr>`;
    });
}

export function latexToOmml(latex: string, style: OmmlRunStyle = {}): string {
  const source = latex.trim();
  if (!source) throw new Error("empty LaTeX");
  const omml = mathMlToOmml(convertLatexToMathMl(source));
  return applyOmmlRunStyle(omml, style);
}

export function tryLatexToOmml(latex: string, style: OmmlRunStyle = {}): string | null {
  try {
    return latexToOmml(latex, style);
  } catch {
    return null;
  }
}

/** Low-level run helper. Prefer `$…$` inside `<FixedBox markdown>` for authored decks. */
export function mathRun(
  latex: string,
  options: PptxGenJS.TextPropsOptions = {},
): PptxGenJS.TextProps {
  const color = typeof options.color === "string" ? options.color : undefined;
  return {
    text: "",
    options: {
      ...options,
      omml: latexToOmml(latex, { fontSizePt: options.fontSize, color }),
    },
  };
}
