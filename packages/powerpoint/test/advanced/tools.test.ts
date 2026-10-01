import type { OfficeHost, ProcessedAttachment, ToolEnv } from "@footnote/core/contracts";
import PptxGenJS from "pptxgenjs";
import { describe, expect, it } from "vitest";
import type { AdvancedToolDeps } from "../../src/advanced/deps.ts";
import { readNotes } from "../../src/advanced/notes.ts";
import { listSlideParts, loadPptx } from "../../src/advanced/ooxml.ts";
import { createAdvancedTools } from "../../src/advanced/tools.ts";

async function makePptx(slideCount: number): Promise<string> {
  const pptx = new PptxGenJS();
  for (let i = 1; i <= slideCount; i++)
    pptx.addSlide().addText(`Slide ${i}`, { x: 1, y: 1, w: 4, h: 1 });
  return (await pptx.write({ outputType: "base64" })) as string;
}

/** In-memory deck: slide ID → single-slide PPTX. Records every op call. */
function fakeHost(slides: Map<string, string>) {
  const calls: { op: string; args: any }[] = [];
  let nextId = 900;
  const call = async (op: string, args: any): Promise<unknown> => {
    calls.push({ op, args });
    switch (op) {
      case "get_deck":
        return {
          slideWidth: 960,
          slideHeight: 540,
          slides: [...slides.keys()].map((id, index) => ({ id, index })),
        };
      case "export_slide":
        return slides.get(args.slideId);
      case "replace_slide": {
        const id = `${nextId++}#`;
        slides.delete(args.slideId);
        slides.set(id, args.base64);
        return { slideId: id, shapes: [] };
      }
      case "insert_slides":
        return {
          changed: ["new#1"],
          createdSlideIds: ["new#1"],
          verified: {},
          warnings: [],
          fingerprints: {},
        };
      default:
        throw new Error(`unexpected op ${op}`);
    }
  };
  return { host: { call } as unknown as OfficeHost, calls };
}

function setup(slides: Map<string, string>, attachments: ProcessedAttachment[] = []) {
  const { host, calls } = fakeHost(slides);
  const hooks: string[] = [];
  const deps: AdvancedToolDeps = {
    beforeWrite: async (ids) => void hooks.push(`before ${ids.join(",")}`),
    afterWrite: async ({ createdSlideIds = [] }) =>
      void hooks.push(`after ${createdSlideIds.join(",")}`),
  };
  const env: ToolEnv = {
    host,
    attachments: { list: () => attachments, get: (id) => attachments.find((a) => a.id === id) },
    settings: () => {
      throw new Error("not used");
    },
  };
  const tools = Object.fromEntries(
    createAdvancedTools(env, deps, () => undefined).map((t) => [t.name, t]),
  );
  return { tools, calls, hooks };
}

const resultJson = (result: { content: { type: string; text?: string }[] }) =>
  JSON.parse(result.content[0]!.text!);

describe("advanced tools", () => {
  it("set_notes checkpoints, replaces the slide, and reports the new ID with read-back", async () => {
    const slides = new Map([["256#1", await makePptx(1)]]);
    const { tools, hooks } = setup(slides);

    const result = await tools.set_notes!.execute("call", {
      notes: [{ slideId: "256#1", text: "Line one\nLine two" }],
    });

    const receipt = resultJson(result);
    expect(receipt.replacedSlideIds).toEqual({ "256#1": "900#" });
    expect(receipt.verified.notesMatch).toEqual({ "900#": true });
    expect(hooks).toEqual(["before 256#1", "after 900#"]);
    expect(await readNotes(slides.get("900#")!)).toBe("Line one\nLine two");
  });

  it("insert_slides_from_file inserts the chosen slides after the slide at position - 1", async () => {
    const deck = new Map([
      ["a#", await makePptx(1)],
      ["b#", await makePptx(1)],
    ]);
    const attachment = {
      id: "att1",
      name: "source.pptx",
      kind: "pptx",
      base64: await makePptx(3),
    } as ProcessedAttachment;
    const { tools, calls, hooks } = setup(deck, [attachment]);

    await tools.insert_slides_from_file!.execute("call", {
      attachmentId: "att1",
      slides: [2],
      position: 2,
    });

    const insert = calls.find((c) => c.op === "insert_slides")!.args;
    expect(insert.targetSlideId).toBe("a#");
    expect(insert.formatting).toBe("UseDestinationTheme");
    expect(await listSlideParts(await loadPptx(insert.base64))).toHaveLength(1);
    expect(hooks).toEqual(["after new#1"]);
  });
});
