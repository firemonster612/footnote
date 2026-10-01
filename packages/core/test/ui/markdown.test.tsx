import { renderToStaticMarkup } from "react-dom/server";
import { describe, expect, test } from "vitest";
import { Markdown } from "../../src/ui/markdown/Markdown.tsx";
import { parseInline, parseMarkdown } from "../../src/ui/markdown/parseMarkdown.ts";

describe("parseMarkdown", () => {
  test("parses headings, paragraphs, fences and rules", () => {
    expect(
      parseMarkdown("## Plan\nfirst line\nsecond line\n\n```ts\nconst a = 1;\n```\n---"),
    ).toEqual([
      { type: "heading", level: 2, children: [{ type: "text", text: "Plan" }] },
      { type: "paragraph", children: [{ type: "text", text: "first line second line" }] },
      { type: "code", lang: "ts", text: "const a = 1;" },
      { type: "rule" },
    ]);
  });

  test("runs an unclosed fence to the end, as while streaming", () => {
    expect(parseMarkdown("```\nline 1\nline 2")).toEqual([
      { type: "code", lang: "", text: "line 1\nline 2" },
    ]);
  });

  test("nests indented lists inside items", () => {
    expect(parseMarkdown("1. one\n   - a\n   - b\n2. two")).toEqual([
      {
        type: "list",
        ordered: true,
        start: 1,
        items: [
          [
            { type: "paragraph", children: [{ type: "text", text: "one" }] },
            {
              type: "list",
              ordered: false,
              start: 1,
              items: [
                [{ type: "paragraph", children: [{ type: "text", text: "a" }] }],
                [{ type: "paragraph", children: [{ type: "text", text: "b" }] }],
              ],
            },
          ],
          [{ type: "paragraph", children: [{ type: "text", text: "two" }] }],
        ],
      },
    ]);
  });

  test("parses GFM tables", () => {
    expect(parseMarkdown("| A | B |\n| --- | ---: |\n| 1 | `x` |")).toEqual([
      {
        type: "table",
        header: [[{ type: "text", text: "A" }], [{ type: "text", text: "B" }]],
        rows: [[[{ type: "text", text: "1" }], [{ type: "code", text: "x" }]]],
      },
    ]);
  });

  test("parses inline emphasis, code and links", () => {
    expect(parseInline("a **b _c_** `d` [e](https://x.test/a_(b))")).toEqual([
      { type: "text", text: "a " },
      {
        type: "strong",
        children: [
          { type: "text", text: "b " },
          { type: "em", children: [{ type: "text", text: "c" }] },
        ],
      },
      { type: "text", text: " " },
      { type: "code", text: "d" },
      { type: "text", text: " " },
      { type: "link", href: "https://x.test/a_(b)", children: [{ type: "text", text: "e" }] },
    ]);
  });

  test("drops links with unsafe schemes", () => {
    expect(parseInline("[click](javascript:alert(1))")).toEqual([{ type: "text", text: "click" }]);
  });
});

describe("Markdown", () => {
  test("renders model HTML as text", () => {
    const html = renderToStaticMarkup(<Markdown text={'<img src=x onerror="alert(1)"> **hi**'} />);
    expect(html).toContain("&lt;img src=x onerror=&quot;alert(1)&quot;&gt;");
    expect(html).toContain("<strong>hi</strong>");
    expect(html).not.toContain("<img");
  });
});
