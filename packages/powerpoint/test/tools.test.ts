import { describe, expect, it } from "vitest";
import type { ShapeInfo } from "../src/ops/types.ts";
import { formatCodeRun } from "../src/tools/code.ts";
import { pageShapes } from "../src/tools/read.ts";

const shape = (id: number, text = "x"): ShapeInfo => ({
  id: String(id),
  name: `Shape ${id}`,
  type: "TextBox",
  left: 0,
  top: 0,
  width: 10,
  height: 10,
  text,
});

describe("pageShapes", () => {
  it("pages by limit and reports the next offset", () => {
    const shapes = Array.from({ length: 5 }, (_, i) => shape(i));
    expect(pageShapes(shapes, 0, 2)).toEqual({ page: shapes.slice(0, 2), nextOffset: 2 });
    expect(pageShapes(shapes, 4, 2)).toEqual({ page: shapes.slice(4) });
  });

  it("stops at the character budget and truncates long text", () => {
    const shapes = Array.from({ length: 20 }, (_, i) => shape(i, "y".repeat(5_000)));
    const { page, nextOffset } = pageShapes(shapes, 0, 40);
    expect(page.length).toBeLessThan(20);
    expect(nextOffset).toBe(page.length);
    expect(page[0]?.text).toMatch(/… \(5000 chars\)$/);
  });
});

describe("formatCodeRun", () => {
  it("tells the model to re-read when the outcome is unknown", () => {
    const text = formatCodeRun({
      ok: false,
      logs: ["started"],
      outcomeUnknown: true,
      error: { message: "timed out after 30s" },
    });
    expect(text).toContain("Outcome unknown (timed out after 30s)");
    expect(text).toContain("Do not rerun it: call get_deck or get_slide first");
    expect(text).toContain("started");
  });

  it("returns the JSON result and error codes", () => {
    expect(formatCodeRun({ ok: true, result: { count: 3 }, logs: [] })).toBe('Result: {"count":3}');
    const text = formatCodeRun({
      ok: false,
      logs: [],
      error: {
        message: "InvalidParam passed to GetItem(id)",
        code: "InvalidArgument",
        debugInfo: { errorLocation: "ShapeCollection.getItem" },
      },
    });
    expect(text).toContain(
      "Error: InvalidParam passed to GetItem(id) (code InvalidArgument; at ShapeCollection.getItem)",
    );
    expect(text).toContain("Edits synced before the failure stay applied");
  });
});
