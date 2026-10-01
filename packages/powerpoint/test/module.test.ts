import type { FootnoteTool, OfficeHost, ToolEnv } from "@footnote/core/contracts";
import { describe, expect, it } from "vitest";
import { createPowerPointModule } from "../src/index.ts";
import type { DeckState, SlideDetail, SlideState, WriteReceipt } from "../src/ops/types.ts";

/** Deck of slide ID → fingerprint, in order. Edit `fingerprints` to simulate the user changing a slide. */
function fakeEnv(slideIds: string[]) {
  const fingerprints = new Map(slideIds.map((id) => [id, "v1"]));
  const calls: { op: string; args: any }[] = [];
  let restores = 0;
  const index = (id: string) => [...fingerprints.keys()].indexOf(id);

  const ops: Record<string, (args: any) => unknown> = {
    get_slide: ({ slideId }): SlideDetail => ({
      id: slideId,
      index: index(slideId),
      layout: "Title Only",
      fingerprint: fingerprints.get(slideId)!,
      shapes: [],
    }),
    get_slide_states: ({
      slideIds: ids,
      exportSlideIds = [],
    }: {
      slideIds: string[];
      exportSlideIds?: string[];
    }) => {
      const states: Record<string, SlideState> = {};
      for (const id of new Set([...ids, ...exportSlideIds])) {
        const fingerprint = fingerprints.get(id);
        if (!fingerprint) continue;
        states[id] = {
          index: index(id),
          fingerprint,
          ...(exportSlideIds.includes(id) && { base64: `pptx:${id}` }),
        };
      }
      return states;
    },
    update_shapes: ({ slideId }): WriteReceipt => {
      fingerprints.set(slideId, `${fingerprints.get(slideId)}+edit`);
      return {
        changed: ["2"],
        verified: {},
        warnings: [],
        fingerprints: { [slideId]: fingerprints.get(slideId)! },
      };
    },
    add_slide: (): WriteReceipt => {
      fingerprints.set("new", "n1");
      return {
        changed: ["new"],
        verified: {},
        warnings: [],
        fingerprints: { new: "n1" },
        createdSlideIds: ["new"],
      };
    },
    restore_slides: ({
      deleteSlideIds,
      inserts,
    }: {
      deleteSlideIds: string[];
      inserts: { slideId: string; base64: string }[];
    }) => {
      // Like PowerPoint: deleted slides go away, re-inserted copies get fresh IDs.
      for (const id of deleteSlideIds) fingerprints.delete(id);
      const idMap: Record<string, string> = {};
      for (const { slideId, base64 } of inserts) {
        restores += 1;
        idMap[slideId] = `${slideId}~r${restores}`;
        fingerprints.set(idMap[slideId], base64);
      }
      return { removedSlideIds: deleteSlideIds, restoredSlideIds: Object.values(idMap), idMap };
    },
    get_deck_state: (): DeckState => ({
      deck: {
        slides: [...fingerprints].map(([id, fingerprint], i) => ({
          id,
          index: i,
          layout: "Title Only",
          shapeCount: 1,
          fingerprint,
        })),
      },
      selection: { slideIds: [], shapeIds: [] },
      theme: { colors: { Accent1: "#0F766E" }, sampled: false, fonts: { heading: "Aptos" } },
    }),
  };

  const host: OfficeHost = {
    status: async () => ({ connected: true, host: "powerpoint", documentName: "Deck.pptx" }),
    onStatusChange: () => () => {},
    async call<T>(op: string, args?: unknown): Promise<T> {
      calls.push({ op, args });
      const handler = ops[op];
      if (!handler) throw new Error(`fake host has no op ${op}`);
      return handler(args) as T;
    },
    runCode: async () => ({ ok: true, logs: [] }),
  };
  const env: ToolEnv = {
    host,
    attachments: { list: () => [], get: () => undefined },
    settings: () => ({}) as any,
  };
  return { env, fingerprints, calls };
}

function tool(tools: FootnoteTool[], name: string): FootnoteTool {
  const found = tools.find((t) => t.name === name);
  if (!found) throw new Error(`missing tool ${name}`);
  return found;
}

const edit = { slideId: "s2", updates: [{ shapeId: "2", text: "Hi" }] };

describe("powerpointModule", () => {
  it("rejects a write to a slide that changed since the model read it", async () => {
    const module = createPowerPointModule();
    const { env, fingerprints, calls } = fakeEnv(["s1", "s2"]);
    module.undo.beginTurn("chat", "t1");
    const tools = module.createTools(env);

    await tool(tools, "get_slide").execute("c1", { slideId: "s2" });
    fingerprints.set("s2", "user-edit");

    await expect(tool(tools, "update_shapes").execute("c2", edit)).rejects.toThrow(
      "Slide 2 changed since you last read it; call get_slide again.",
    );
    expect(calls.some((call) => call.op === "update_shapes")).toBe(false);
  });

  it("snapshots a slide once per turn and undo restores it", async () => {
    const module = createPowerPointModule();
    const { env, calls } = fakeEnv(["s1", "s2"]);
    module.undo.beginTurn("chat", "t1");
    const tools = module.createTools(env);
    await module.getContextBlock(env, "chat");

    await tool(tools, "get_slide").execute("c1", { slideId: "s2" });
    await tool(tools, "update_shapes").execute("c2", edit);
    await tool(tools, "update_shapes").execute("c3", edit); // own receipt keeps the read fresh

    const exports = calls
      .filter((call) => call.op === "get_slide_states")
      .map((call) => call.args.exportSlideIds);
    expect(exports).toEqual([["s2"], []]);
    expect(module.undo.canUndo("chat")).toBe(true);

    const report = await module.undo.undoLastTurn(env, "chat");
    expect(calls.at(-1)).toEqual({
      op: "restore_slides",
      args: { deleteSlideIds: ["s2"], inserts: [{ slideId: "s2", index: 1, base64: "pptx:s2" }] },
    });
    expect(report.restored).toBe(1);
    expect(module.undo.canUndo("chat")).toBe(false);
  });

  it("undoes several turns on the same slide without leaving duplicate copies", async () => {
    const module = createPowerPointModule();
    const { env, fingerprints } = fakeEnv(["s1", "s2"]);
    module.undo.beginTurn("chat", "t1");
    const tools = module.createTools(env);
    await module.getContextBlock(env, "chat");
    await tool(tools, "get_slide").execute("c1", { slideId: "s2" });
    await tool(tools, "update_shapes").execute("c2", edit);
    module.undo.beginTurn("chat", "t2");
    await tool(tools, "update_shapes").execute("c3", edit);

    const report = await module.undo.undoTurns(env, "chat", ["t1", "t2", "t-from-before-reload"]);

    // t2's restore gives s2 a new ID; t1's undo must replace that copy, not add a second one.
    expect([...fingerprints.keys()]).toHaveLength(2);
    expect([...fingerprints.values()]).toContain("pptx:s2");
    expect(report.restored).toBe(2);
    expect(report.warnings.join(" ")).toContain("1 earlier turn couldn't be undone");
    expect(module.undo.canUndo("chat")).toBe(false);
  });

  it("undo removes created slides and warns about edits made after the turn", async () => {
    const module = createPowerPointModule();
    const { env, fingerprints, calls } = fakeEnv(["s1"]);
    module.undo.beginTurn("chat", "t1");
    const tools = module.createTools(env);
    await module.getContextBlock(env, "chat");

    await tool(tools, "add_slide").execute("c1", {});
    fingerprints.set("new", "user-edit");
    module.undo.beginTurn("chat", "t2");

    const report = await module.undo.undoLastTurn(env, "chat");
    expect(calls.at(-1)).toEqual({
      op: "restore_slides",
      args: { deleteSlideIds: ["new"], inserts: [] },
    });
    expect(report.removed).toBe(1);
    expect(report.warnings).toEqual([
      "Slide 2 was edited after that turn; undo discarded those later edits.",
    ]);
  });

  it("deck state shows the theme once and reports slides modified since the previous block", async () => {
    const module = createPowerPointModule();
    const { env, fingerprints } = fakeEnv(["s1", "s2"]);

    const first = await module.getContextBlock(env, "chat");
    expect(first).toContain("Theme:");
    expect(first).not.toContain("Changes since");

    fingerprints.set("s1", "user-edit");
    const second = await module.getContextBlock(env, "chat");
    expect(second).not.toContain("Theme:");
    expect(second).toContain("Changes since last update: modified 1 (s1)");
  });
});
