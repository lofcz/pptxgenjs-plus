import { describe, expect, test } from "bun:test";
import { JSZip } from "@node-projects/jszip";
import { Deck, FixedBox, Slide, fromMarkdown, headingSize, latexToOmml, tokenizeMath } from "../src/index.ts";
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

  test("sizes every ATX heading distinctly and bold", () => {
    const body = 18;
    const runs = fromMarkdown("# H1\n## H2\n### H3\n#### H4\n##### H5\n###### H6", { fontSize: body });
    const sizes = ["H1", "H2", "H3", "H4", "H5", "H6"].map((text) => {
      const run = runs.find((r) => r.text === text);
      expect(run?.options?.bold).toBe(true);
      expect(run?.options?.fontSize).toBe(headingSize(Number(text.slice(1)), body));
      return run?.options?.fontSize ?? 0;
    });
    expect(new Set(sizes).size).toBe(6);
    expect(sizes[0]).toBeGreaterThan(sizes[1]!);
    expect(sizes[5]).toBeGreaterThan(body);
  });

  test("parses setext headings as h1/h2", () => {
    const runs = fromMarkdown("Big title\n=========\n\nSmaller\n-------", { fontSize: 18 });
    const h1 = runs.find((r) => r.text === "Big title");
    const h2 = runs.find((r) => r.text === "Smaller");
    expect(h1?.options?.fontSize).toBe(headingSize(1, 18));
    expect(h2?.options?.fontSize).toBe(headingSize(2, 18));
    expect(h1?.options?.bold).toBe(true);
    expect(h2?.options?.bold).toBe(true);
  });

  test("applies bold, italic, nested emphasis, and strike", () => {
    const runs = fromMarkdown("**bold** *italic* ***both*** ~~gone~~ __also__ _em_");
    expect(runs.find((r) => r.text === "bold")?.options?.bold).toBe(true);
    expect(runs.find((r) => r.text === "italic")?.options?.italic).toBe(true);
    const both = runs.find((r) => r.text === "both");
    expect(both?.options?.bold).toBe(true);
    expect(both?.options?.italic).toBe(true);
    expect(runs.find((r) => r.text === "gone")?.options?.strike).toBe("sngStrike");
    expect(runs.find((r) => r.text === "also")?.options?.bold).toBe(true);
    expect(runs.find((r) => r.text === "em")?.options?.italic).toBe(true);
  });

  test("emits a horizontal rule for --- *** and ___", () => {
    for (const source of ["before\n\n---\n\nafter", "before\n\n***\n\nafter", "before\n\n___\n\nafter"]) {
      const runs = fromMarkdown(source);
      const rule = runs.find((r) => typeof r.text === "string" && r.text.includes("─"));
      expect(rule?.text?.length).toBeGreaterThan(10);
      expect((rule?.options?.fontSize ?? 18) < 18).toBe(true);
    }
  });

  test("underlines links and keeps the href", () => {
    const runs = fromMarkdown("see [docs](https://example.com/path)");
    const link = runs.find((r) => r.text === "docs");
    expect(link?.options?.underline).toEqual({ style: "sng" });
    expect(link?.options?.hyperlink).toEqual({ url: "https://example.com/path" });
  });

  test("italicizes and indents block quotes", () => {
    const runs = fromMarkdown("> quoted line");
    const quote = runs.find((r) => r.text === "quoted line");
    expect(quote?.options?.italic).toBe(true);
    expect(quote?.options?.indentLevel).toBe(1);
  });

  test("turns hard line breaks into breakLine", () => {
    const runs = fromMarkdown("line one  \nline two");
    const first = runs.find((r) => r.text === "line one");
    expect(first?.options?.breakLine).toBe(true);
    expect(runs.some((r) => r.text === "line two")).toBe(true);
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
          markdown:
            "# Title\n## Sub\n- item one\n- **bold** *italic* ~~strike~~ [docs](https://example.com)\n\n---\n\n> quote\n\n$a^2+b^2=c^2$\n\n$$\\frac{1}{2}$$",
        }),
      ),
    );

    const buffer = await writePptx(deck, { outputType: "nodebuffer" });
    const zip = await JSZip.loadAsync(buffer as Buffer);
    const xml = await zip.file("ppt/slides/slide1.xml")?.async("string");
    expect(xml).toBeTruthy();
    expect(xml!).toContain("<a:normAutofit/>");
    expect(xml!).toMatch(/<a:bu[A-Za-z]+/);
    expect(xml!).toContain('strike="sngStrike"');
    expect(xml!).toContain('u="sng"');
    expect(xml!).toContain("─");
    expect(xml!).toContain("<m:sSup>");
    expect(xml!).toContain("<m:f>");
    expect(xml!).toContain("<a14:m");
  });
});
