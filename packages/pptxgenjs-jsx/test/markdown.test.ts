import { describe, expect, test } from "bun:test";
import { JSZip } from "@node-projects/jszip";
import { Deck, FixedBox, Slide, fromMarkdown, latexToOmml, tokenizeMath } from "../src/index.ts";
import { pptxElement } from "../src/jsx-runtime.ts";
import { writePptx } from "../src/render.ts";

describe("tokenizeMath", () => {
  test("splits inline and display dollars, skips code and $1", () => {
    const parts = tokenizeMath("see `$x$` and $a^2$ then $$\\frac{1}{2}$$ cost $1");
    expect(parts.filter((p) => p.type === "math").map((p) => ({ value: p.value, display: p.display }))).toEqual([
      { value: "a^2", display: false },
      { value: "\\frac{1}{2}", display: true },
    ]);
    expect(parts.some((p) => p.type === "text" && p.value.includes("`$x$`"))).toBe(true);
    expect(parts.some((p) => p.type === "text" && p.value.includes("$1"))).toBe(true);
  });

  test("treats \\( \\) and \\[ \\] as math", () => {
    const parts = tokenizeMath("inline \\(x+1\\) and display \\[y=2\\]");
    expect(parts.filter((p) => p.type === "math")).toEqual([
      { type: "math", raw: "\\(x+1\\)", value: "x+1", display: false },
      { type: "math", raw: "\\[y=2\\]", value: "y=2", display: true },
    ]);
  });
});

describe("fromMarkdown", () => {
  test("emits bullets, bold, and OMML for $ / $$", () => {
    const runs = fromMarkdown(
      `# Title\n- item **one**\n- item with $a^2+b^2=c^2$\n\n$$\\frac{1}{2}$$`,
      { fontSize: 18, color: "0F172A" },
    );

    const title = runs.find((r) => r.text === "Title");
    expect(title?.options?.bold).toBe(true);
    expect((title?.options?.fontSize ?? 0) > 18).toBe(true);

    const firstBullet = runs.find((r) => r.options?.bullet === true);
    expect(firstBullet?.text).toContain("item");

    const numbered = fromMarkdown("1. first\n2. second");
    expect(numbered.some((r) => typeof r.options?.bullet === "object" && r.options.bullet.type === "number")).toBe(
      true,
    );

    const mathRuns = runs.filter((r) => r.options?.omml);
    expect(mathRuns.length).toBe(2);
    expect(mathRuns[0]?.options?.omml).toMatch(/<m:oMath[\s/>]/);
    expect(mathRuns[0]?.options?.omml).toMatch(/<m:sSup[\s/>]/);
    expect(mathRuns[1]?.options?.omml).toMatch(/<m:f[\s/>]/);
  });

  test("does not treat escaped dollars or currency as math", () => {
    const runs = fromMarkdown("price is $12 and not \\$x$");
    expect(runs.every((r) => !r.options?.omml)).toBe(true);
  });
});

describe("latexToOmml", () => {
  test("produces structured OMML for a fraction", () => {
    const omml = latexToOmml("\\frac{2}{7}");
    expect(omml).toContain("<m:f>");
    expect(omml).toContain("<m:num>");
    expect(omml).toContain("<m:den>");
  });
});

describe("FixedBox", () => {
  test("writes shrink-fit, bullets, and real OMML into the slide", async () => {
    const deck = pptxElement(
      Deck,
      { title: "FixedBox" },
      pptxElement(
        Slide,
        {},
        pptxElement(FixedBox, {
          x: 0.7,
          y: 1.5,
          w: 6,
          h: 4,
          fontSize: 18,
          color: "0F172A",
          markdown: "# Title\n- item one\n- $a^2+b^2=c^2$\n\n$$\\frac{1}{2}$$",
        }),
      ),
    );

    const buffer = await writePptx(deck, { outputType: "nodebuffer" });
    const zip = await JSZip.loadAsync(buffer as Buffer);
    const xml = await zip.file("ppt/slides/slide1.xml")?.async("string");
    expect(xml).toBeTruthy();
    expect(xml!).toContain("<a:normAutofit/>");
    expect(xml!).toMatch(/<a:bu[A-Za-z]+/);
    expect(xml!).toContain("<m:sSup>");
    expect(xml!).toContain("<m:f>");
    expect(xml!).toContain("<a14:m");
  });
});
