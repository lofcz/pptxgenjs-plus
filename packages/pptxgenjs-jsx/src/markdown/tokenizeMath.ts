/**
 * Split source into plain text and math spans. Same delimiter rules as Fika
 * (`$…$`, `$$…$$`, `\(…\)`, `\[…\]`, `\begin{env}…\end{env}`): `$$` before `$`,
 * skip inline code, honor `\$`, and do not treat `$1` as math.
 */

export interface ContentSegment {
  type: "text" | "math";
  raw: string;
  value: string;
  display: boolean;
}

const DOLLAR_OPENERS = [
  { open: "$$", display: true, multiline: true },
  { open: "$", display: false, multiline: false },
] as const;

const BRACKET_OPENERS = [
  { open: "\\[", close: "\\]", display: true, multiline: true },
  { open: "\\(", close: "\\)", display: false, multiline: false },
] as const;

const BEGIN_RE = /\\begin\{([a-zA-Z*]+)\}/y;

function findPlainClose(text: string, from: number, close: string, multiline: boolean): number {
  const idx = text.indexOf(close, from);
  if (idx === -1) return -1;
  if (!multiline) {
    const nl = text.indexOf("\n", from);
    if (nl !== -1 && nl < idx) return -1;
  }
  return idx;
}

function findDollarClose(text: string, from: number, isDouble: boolean): number {
  for (let j = from; j < text.length; j++) {
    const ch = text[j];
    if (ch === "\\") {
      j += 1;
      continue;
    }
    if (!isDouble && ch === "\n") return -1;
    if (ch === "$") {
      if (!isDouble) return j;
      if (text[j + 1] === "$") return j;
    }
  }
  return -1;
}

export function tokenizeMath(source: string): ContentSegment[] {
  const text = String(source);
  const n = text.length;
  const out: ContentSegment[] = [];
  let textStart = 0;
  let i = 0;
  const flushText = (end: number) => {
    if (end > textStart) {
      const value = text.slice(textStart, end);
      out.push({ type: "text", raw: value, value, display: false });
    }
  };
  const pushMath = (
    start: number,
    innerStart: number,
    innerEnd: number,
    end: number,
    display: boolean,
  ) => {
    flushText(start);
    out.push({
      type: "math",
      raw: text.slice(start, end),
      value: text.slice(innerStart, innerEnd),
      display,
    });
    i = end;
    textStart = end;
  };
  while (i < n) {
    const ch = text[i];

    if (ch === "`") {
      let run = 1;
      while (text[i + run] === "`") run += 1;
      const close = text.indexOf("`".repeat(run), i + run);
      i = close === -1 ? i + run : close + run;
      continue;
    }
    if (ch === "\\") {
      BEGIN_RE.lastIndex = i;
      const begin = BEGIN_RE.exec(text);
      if (begin && begin.index === i) {
        const closer = `\\end{${begin[1]}}`;
        const innerStart = i + begin[0].length;
        const end = text.indexOf(closer, innerStart);
        if (end !== -1) {
          pushMath(i, innerStart, end, end + closer.length, true);
          continue;
        }
        i += begin[0].length;
        continue;
      }
      const bracket = BRACKET_OPENERS.find((d) => text.startsWith(d.open, i));
      if (bracket) {
        const innerStart = i + bracket.open.length;
        const end = findPlainClose(text, innerStart, bracket.close, bracket.multiline);
        if (end !== -1) {
          pushMath(i, innerStart, end, end + bracket.close.length, bracket.display);
          continue;
        }
      }
      i += 2;
      continue;
    }
    if (ch === "$") {
      const isDouble = text[i + 1] === "$";
      if (!isDouble && text[i + 1] >= "0" && text[i + 1] <= "9") {
        i += 1;
        continue;
      }
      const opener = isDouble ? DOLLAR_OPENERS[0] : DOLLAR_OPENERS[1];
      const innerStart = i + opener.open.length;
      const end = findDollarClose(text, innerStart, isDouble);
      if (end !== -1) {
        pushMath(i, innerStart, end, end + opener.open.length, opener.display);
        continue;
      }
      i += opener.open.length;
      continue;
    }
    i += 1;
  }
  flushText(n);
  return out;
}
