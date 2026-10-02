import {
  type FootnoteTool,
  type OfficeHost,
  OfficeOpError,
  type ToolEnv,
} from "@footnote/core/contracts";
import { describe, expect, it } from "vitest";
import { createPowerPointModule } from "../src/index.ts";
import { createFootnoteHelpers, powerpointOps } from "../src/ops/index.ts";
import { type FakeDeck, fakeShape, fakeSlide, installFakeOffice } from "./fakeOffice.ts";

function fakeDeck(): FakeDeck {
  return {
    layouts: [
      { id: "layout-title", name: "Title Slide" },
      { id: "layout-content", name: "Title and Content" },
    ],
    selectedSlideIds: ["s2"],
    slides: ["s1", "s2", "s3"].map((id) =>
      fakeSlide(id, [
        fakeShape(`${id}-title`, { placeholder: "Title", text: `Title ${id}` }),
        fakeShape(`${id}-body`, { placeholder: "Body", text: "Body" }),
        fakeShape(`${id}-note`, { text: "Note" }),
        fakeShape(`${id}-line`, { type: "Line" }),
        fakeShape(`${id}-table`, {
          type: "Table",
          table: {
            rowCount: 2,
            columnCount: 2,
            values: [
              ["a", "b"],
              ["c", "d"],
            ],
            rows: [{}, {}],
            columns: [{}, {}],
          },
        }),
      ]),
    ),
  };
}

/** Tools wired to the real ops through an in-process host, over a fake Office that counts syncs. */
function setup(apiVersion?: string) {
  const deck = fakeDeck();
  const office = installFakeOffice(deck, apiVersion);
  const calls: string[] = [];
  const host: OfficeHost = {
    status: async () => ({ connected: true, host: "powerpoint" }),
    onStatusChange: () => () => {},
    async call<T>(op: string, args?: unknown): Promise<T> {
      calls.push(op);
      try {
        return structuredClone(await powerpointOps[op]!(structuredClone(args))) as T;
      } catch (error) {
        const message = error instanceof Error ? error.message : String(error);
        const code = (error as { code?: string }).code;
        throw new OfficeOpError(op, { message, ...(code && { code }) });
      }
    },
    runCode: async () => ({ ok: true, logs: [] }),
  };
  const env: ToolEnv = {
    host,
    attachments: { list: () => [], get: () => undefined },
    settings: () => ({}) as any,
  };
  const module = createPowerPointModule();
  module.undo.beginTurn("chat", "t1");
  const tools = module.createTools(env);
  const tool = (name: string): FootnoteTool => tools.find((candidate) => candidate.name === name)!;
  const read = async (...slideIds: string[]) => {
    for (const slideId of slideIds) await tool("get_slide").execute("read", { slideId });
  };

  /** Runs one tool call and counts its host calls and syncs. */
  async function measure(name: string, args: unknown) {
    calls.length = 0;
    office.reset();
    const result = await tool(name).execute("call", args);
    const { runs: _, ...syncs } = office.totals();
    return { counts: { hostCalls: calls.length, ...syncs }, result };
  }
  return { deck, office, module, env, measure, read, tool };
}

const text = (deck: FakeDeck, slideId: string, shapeId: string) =>
  deck.slides.find((slide) => slide.id === slideId)?.shapes.find((shape) => shape.id === shapeId)
    ?.textFrame?.textRange.text;

const receipt = (result: { content: { type: string; text?: string }[] }) =>
  JSON.parse(result.content[0]!.text!);

const ids = (deck: FakeDeck) => deck.slides.map((slide) => slide.id);
const titles = (deck: FakeDeck) =>
  deck.slides.map((slide) => slide.shapes[0]!.textFrame!.textRange.text);

describe("each structured write is one host call with one mutation sync", () => {
  // Before batching every guarded write took 2 host calls and 10-14 syncs, 2-3 of which changed the deck.
  // The check before a write still costs up to 3 read-only syncs (slide list → shape types → placeholder text).
  it.each([
    [
      "update_shapes (3 shapes)",
      "update_shapes",
      {
        slides: [
          {
            slideId: "s2",
            updates: [
              { shapeId: "s2-title", text: "New title" },
              { shapeId: "s2-body", text: "New body", font: { bold: true } },
              { shapeId: "s2-note", left: 200, top: 300 },
            ],
          },
        ],
      },
      5,
    ],
    [
      "add_shape",
      "add_shape",
      { slideId: "s2", kind: "textbox", left: 10, top: 10, width: 200, height: 40, text: "Hi" },
      5,
    ],
    ["add_slide", "add_slide", { layoutId: "layout-content", position: 2 }, 3],
    [
      "edit_table",
      "edit_table",
      {
        slideId: "s2",
        shapeId: "s2-table",
        insertRows: { index: 2, count: 1 },
        cells: [
          { row: 0, column: 0, text: "Q1" },
          { row: 2, column: 1, text: "Total", font: { bold: true } },
        ],
      },
      5,
    ],
    ["apply_layout", "apply_layout", { slideId: "s2", layoutId: "layout-title" }, 6],
  ])("%s", async (_label, name, args, syncs) => {
    const { measure, read } = setup();
    await read("s2");
    const { counts } = await measure(name, args);
    expect(counts).toEqual({ hostCalls: 1, syncs, writeSyncs: 1, selectionSyncs: 0 });
  });

  it("add_slide lands at its position and returns its placeholders", async () => {
    const { measure, deck } = setup();
    const { result } = await measure("add_slide", { layoutId: "layout-content", position: 2 });
    const { changed, verified, shapes } = receipt(result);
    expect(deck.slides[1]!.id).toBe(changed[0]);
    expect(verified.slides[0]).toMatchObject({ position: 2, layout: "Title and Content" });
    expect(shapes).toEqual([expect.objectContaining({ placeholder: "Title" })]);
  });

  it("delete_slides moves the selection off the on-screen slide in its own sync, then deletes", async () => {
    const { measure, read, deck } = setup();
    await read("s2");
    const { counts, result } = await measure("delete_slides", { slideIds: ["s2"] });
    expect(counts).toEqual({ hostCalls: 1, syncs: 5, writeSyncs: 1, selectionSyncs: 1 });
    expect(deck.slides.map((slide) => slide.id)).toEqual(["s1", "s3"]);
    expect(receipt(result).verified).toEqual({ slides: [], slideCount: 2 });
  });
});

describe("update_shapes across slides", () => {
  it("edits several slides in one host call and one mutation sync, with read-back per slide", async () => {
    const { measure, read, deck } = setup();
    await read("s1", "s3");
    const { counts, result } = await measure("update_shapes", {
      slides: [
        { slideId: "s1", updates: [{ shapeId: "s1-title", text: "One" }] },
        { slideId: "s3", updates: [{ shapeId: "s3-title", text: "Three" }] },
      ],
    });
    expect(counts).toMatchObject({ hostCalls: 1, writeSyncs: 1 });
    expect([text(deck, "s1", "s1-title"), text(deck, "s3", "s3-title")]).toEqual(["One", "Three"]);
    const { verified, warnings } = receipt(result);
    expect(warnings).toBeUndefined();
    expect(verified.s1["s1-title"].text).toBe("One");
    expect(verified.s3["s3-title"].text).toBe("Three");
  });

  it("applies the rest when one update can't, and names the ones that didn't", async () => {
    const { measure, read, deck } = setup();
    await read("s2");
    const { counts, result } = await measure("update_shapes", {
      slides: [
        {
          slideId: "s2",
          updates: [
            { shapeId: "s2-title", text: "Kept" },
            { shapeId: "s2-line", text: "Lines have no text" },
            { shapeId: "missing", left: 5 },
          ],
        },
      ],
    });
    expect(counts.writeSyncs).toBe(1);
    expect(text(deck, "s2", "s2-title")).toBe("Kept");
    expect(receipt(result).warnings).toEqual([
      "Shape s2-line on slide s2: A Line has no text, so text and font fields don't apply.",
      "Shape missing not found on slide s2. Call get_slide for current shape IDs.",
    ]);
  });

  it("keeps shape IDs apart per slide (placeholders share IDs across slides)", async () => {
    const { measure, read, deck } = setup();
    for (const slide of deck.slides) slide.shapes[0]!.id = "2";
    await read("s1", "s3");
    const { result } = await measure("update_shapes", {
      slides: [
        { slideId: "s1", updates: [{ shapeId: "2", delete: true }] },
        { slideId: "s3", updates: [{ shapeId: "2", text: "Three" }] },
      ],
    });
    const { verified, warnings } = receipt(result);
    expect(warnings).toBeUndefined();
    expect(verified.s1["2"]).toBe("deleted");
    expect(verified.s3["2"].text).toBe("Three");
  });

  it("refuses the whole batch when any slide changed since the model read it", async () => {
    const { measure, read, deck } = setup();
    await read("s1", "s3");
    deck.slides[2]!.shapes[2]!.textFrame!.textRange.text = "edited by the user";
    await expect(
      measure("update_shapes", {
        slides: [
          { slideId: "s1", updates: [{ shapeId: "s1-title", text: "One" }] },
          { slideId: "s3", updates: [{ shapeId: "s3-title", text: "Three" }] },
        ],
      }),
    ).rejects.toThrow(/^Slide 3 changed since you last read it; call get_slide again\.$/);
    expect(text(deck, "s1", "s1-title")).toBe("Title s1");
  });

  it("reports a failed read-back as a warning, not an error, so the model doesn't repeat the write", async () => {
    const { measure, read, office, deck } = setup();
    await read("s2");
    office.faults.failReadsAfterWrite = true;
    const { result } = await measure("update_shapes", {
      slides: [{ slideId: "s2", updates: [{ shapeId: "s2-title", text: "Applied" }] }],
    });
    expect(text(deck, "s2", "s2-title")).toBe("Applied");
    expect(receipt(result).warnings[0]).toMatch(
      /^The write was applied, but reading it back failed/,
    );
  });
});

it("undo restores a slide from the snapshot the write exported before changing it", async () => {
  const { measure, read, module, env, deck } = setup();
  await module.getContextBlock(env, "chat");
  await read("s2");
  const edit = { slides: [{ slideId: "s2", updates: [{ shapeId: "s2-title", text: "Edited" }] }] };
  await measure("update_shapes", edit);
  await measure("update_shapes", edit);

  const report = await module.undo.undoLastTurn(env, "chat");
  expect(report.restored).toBe(1);
  expect(deck.slides).toHaveLength(3);
  expect(deck.slides[1]!.shapes[0]!.textFrame!.textRange.text).toBe("Title s2");
});

describe("writes that committed are never reported as failures", () => {
  it("add_shape returns the created shape when styling fails and the failed batch can't be read again", async () => {
    const { measure, read, office, deck } = setup();
    await read("s2");
    office.faults.failReadsAfterWrite = true;
    const { result } = await measure("add_shape", {
      slideId: "s2",
      kind: "line",
      text: "Lines have no text frame",
      left: 0,
      top: 0,
      width: 100,
      height: 0,
    });
    expect(deck.slides[1]!.shapes).toHaveLength(6);
    expect(receipt(result).warnings[0]).toMatch(/^Shape \d+ was created but styling failed/);
  });

  it("duplicate_slide keeps the copy for undo when moving it fails", async () => {
    const { measure, office, deck, module, env } = setup();
    office.faults.failReadsAfterWrite = true;
    const { result } = await measure("duplicate_slide", { slideId: "s2", position: 1 });
    const { changed, warnings } = receipt(result);
    expect(ids(deck)).toEqual(["s1", "s2", changed[0], "s3"]);
    expect(warnings[0]).toMatch(/moving it failed/);

    office.faults.failReadsAfterWrite = false;
    await module.undo.undoLastTurn(env, "chat");
    expect(ids(deck)).toEqual(["s1", "s2", "s3"]);
  });

  it("writes proceed without undo snapshots when PowerPointApi 1.8 is missing", async () => {
    const { measure, read, deck } = setup("1.7");
    await read("s2");
    const { result } = await measure("update_shapes", {
      slides: [{ slideId: "s2", updates: [{ shapeId: "s2-title", text: "Applied" }] }],
    });
    expect(text(deck, "s2", "s2-title")).toBe("Applied");
    expect(receipt(result).warnings).toEqual([expect.stringMatching(/undo/i)]);
  });
});

describe("undo", () => {
  it("puts a moved slide back where it was", async () => {
    const { measure, module, env, deck } = setup();
    await module.getContextBlock(env, "chat");
    await measure("move_slide", { slideId: "s1", position: 3 });
    expect(ids(deck)).toEqual(["s2", "s3", "s1"]);

    await module.undo.undoLastTurn(env, "chat");
    expect(titles(deck)).toEqual(["Title s1", "Title s2", "Title s3"]);
  });

  it("finishes on the next undo after stopping part-way, without leaving copies", async () => {
    const { measure, read, module, env, office, deck } = setup();
    await module.getContextBlock(env, "chat");
    await read("s1", "s3");
    await measure("update_shapes", {
      slides: [
        { slideId: "s1", updates: [{ shapeId: "s1-title", text: "Edited" }] },
        { slideId: "s3", updates: [{ shapeId: "s3-title", text: "Edited" }] },
      ],
    });

    office.faults.failReadsAfterWrite = true;
    const partial = await module.undo.undoLastTurn(env, "chat");
    expect(partial.warnings.join(" ")).toMatch(/stopped part-way/);
    expect(module.undo.canUndo("chat")).toBe(true);

    office.faults.failReadsAfterWrite = false;
    await module.undo.undoLastTurn(env, "chat");
    expect(titles(deck)).toEqual(["Title s1", "Title s2", "Title s3"]);
    expect(module.undo.canUndo("chat")).toBe(false);
  });
});

it("delete_slides refuses to delete every slide before changing anything", async () => {
  const { measure, deck } = setup();
  await expect(measure("delete_slides", { slideIds: ["s1", "s2", "s3"] })).rejects.toThrow(
    /every slide/,
  );
  expect(ids(deck)).toEqual(["s1", "s2", "s3"]);
});

it("replace_slide refuses a slide that changed after it was exported", async () => {
  const { deck } = setup();
  const states = (await powerpointOps.get_slide_states!({
    slideIds: ["s2"],
    exportSlideIds: ["s2"],
  })) as Record<string, { fingerprint: string; base64: string }>;
  deck.slides[1]!.shapes[0]!.textFrame!.textRange.text = "Edited by the user";
  await expect(
    powerpointOps.replace_slide!({
      slideId: "s2",
      base64: states.s2!.base64,
      expectedFingerprints: { s2: states.s2!.fingerprint },
    }),
  ).rejects.toThrow(/changed since you last read it/);
  expect(ids(deck)).toEqual(["s1", "s2", "s3"]);
});

describe("code mode helpers", () => {
  it("textShapes skips placeholders without a text frame", async () => {
    const { deck } = setup();
    const picture = fakeShape("s1-picture", { placeholder: "Picture" });
    delete picture.textFrame;
    deck.slides[0]!.shapes.push(picture);
    const found = await PowerPoint.run(async (context) => {
      const footnote = createFootnoteHelpers(context);
      return (await footnote.textShapes(footnote.slide(1))).map((shape) => shape.id);
    });
    expect(found).toEqual(["s1-title", "s1-body", "s1-note"]);
  });

  it("deleteSlides moves the selection off the on-screen slide first", async () => {
    const { deck } = setup();
    await PowerPoint.run((context) => createFootnoteHelpers(context).deleteSlides(["s2"]));
    expect(ids(deck)).toEqual(["s1", "s3"]);
  });
});

it("code mode: loading every slide first and syncing once at the end lands the edits in one batch", async () => {
  const { office, deck } = setup();
  await PowerPoint.run(async (context) => {
    const footnote = createFootnoteHelpers(context);
    const shapes = await footnote.textShapes([1, 2, 3].map((number) => footnote.slide(number)));
    for (const { shape, text: current } of shapes) footnote.setText(shape, current.toUpperCase());
    await context.sync();
  });
  expect(office.totals()).toMatchObject({ syncs: 3, writeSyncs: 1 });
  expect(text(deck, "s3", "s3-note")).toBe("NOTE");
});

it("tool-call labels name slides by position once the outline is known, by ID before that", async () => {
  const { tool, module, env } = setup();
  expect(tool("get_slide").describeCall!({ slideId: "s3" })).toBe("Read slide s3");
  await module.getContextBlock(env, "chat");
  expect(tool("get_slide").describeCall!({ slideId: "s3" })).toBe("Read slide 3");
  expect(
    tool("update_shapes").describeCall!({
      slides: [
        { slideId: "s1", updates: [{}] },
        { slideId: "s3", updates: [{}, {}] },
      ],
    }),
  ).toBe("Update 3 shapes on slides 1, 3");
  expect(tool("insert_chart").describeCall!({ type: "bar", slideId: "s2" })).toBe(
    "Insert bar chart on slide 2",
  );
});
