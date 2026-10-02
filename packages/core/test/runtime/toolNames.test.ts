import {
  createFauxCore,
  fauxAssistantMessage,
  fauxToolCall,
  normalizeContext,
} from "@earendil-works/pi-ai";
import { describe, expect, it } from "vitest";
import { repairToolName, withToolNameRepair } from "../../src/runtime/toolNames.ts";

const known = ["execute_office_js", "get_slide", "add_slide"];

describe("repairToolName", () => {
  it("maps a prefixed name to the registered tool it ends with", () => {
    expect(repairToolName("someone_execute_office_js", known)).toBe("execute_office_js");
    expect(repairToolName("functions.get_slide", known)).toBe("get_slide");
  });

  it("leaves known, unknown and ambiguous names alone", () => {
    expect(repairToolName("add_slide", known)).toBe("add_slide");
    expect(repairToolName("make_coffee", known)).toBe("make_coffee");
    expect(repairToolName("x_slide", ["get_slide", "x_slide_y"])).toBe("x_slide");
  });
});

describe("withToolNameRepair", () => {
  it("repairs names in streamed tool calls and in the final message", async () => {
    const faux = createFauxCore({ api: "faux", provider: "faux", models: [{ id: "faux-model" }] });
    faux.setResponses([fauxAssistantMessage(fauxToolCall("someone_execute_office_js", {}))]);
    const streamFn = withToolNameRepair(faux.streamSimple, () => known);

    const stream = await streamFn(faux.getModel(), normalizeContext({ messages: [] }));
    const streamedNames: string[] = [];
    for await (const event of stream) {
      if (event.type === "toolcall_end") streamedNames.push(event.toolCall.name);
    }
    const message = await stream.result();

    expect(streamedNames).toEqual(["execute_office_js"]);
    expect(message.content).toEqual([expect.objectContaining({ name: "execute_office_js" })]);
  });
});
