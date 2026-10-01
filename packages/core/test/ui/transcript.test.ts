import type { ToolResultMessage } from "@earendil-works/pi-ai";
import { describe, expect, test } from "vitest";
import type { ApprovalRequest } from "../../src/contracts.ts";
import {
  clampThinkingLevel,
  formatTokens,
  splitUserContent,
  toolCallStatus,
} from "../../src/ui/chat/transcript.ts";

const result = (isError: boolean): ToolResultMessage => ({
  role: "toolResult",
  toolCallId: "c1",
  toolName: "get_slide",
  content: [],
  isError,
  timestamp: 0,
});

const approval: ApprovalRequest = {
  id: "a1",
  toolCallId: "c1",
  toolName: "update_shapes",
  access: "write",
  summary: "",
  args: {},
};

describe("toolCallStatus", () => {
  test("a result decides the status", () => {
    expect(toolCallStatus("c1", result(false), [approval], true)).toBe("done");
    expect(toolCallStatus("c1", result(true), [], false)).toBe("error");
  });

  test("without a result: approval, then running while streaming, else not run", () => {
    expect(toolCallStatus("c1", undefined, [approval], true)).toBe("awaiting-approval");
    expect(toolCallStatus("c2", undefined, [approval], true)).toBe("running");
    expect(toolCallStatus("c2", undefined, [], false)).toBe("not-run");
  });
});

describe("splitUserContent", () => {
  test("separates typed text, attachment excerpts, images and context blocks", () => {
    const image = { type: "image", data: "AAAA", mimeType: "image/png" } as const;
    expect(
      splitUserContent([
        { type: "text", text: "Fix slide 2" },
        {
          type: "text",
          text: '<attachment id="a" name="report.pdf" kind="pdf">page 1…</attachment>',
        },
        {
          type: "text",
          text: `<attachment id="b" name=${JSON.stringify('say "hi".txt')}>hi</attachment>`,
        },
        { type: "text", text: "<deck_state>…</deck_state>" },
        image,
      ]),
    ).toEqual({
      text: "Fix slide 2",
      attachmentNames: ["report.pdf", 'say "hi".txt'],
      images: [image],
    });
  });

  test("passes plain string content through", () => {
    expect(splitUserContent("hello")).toEqual({ text: "hello", attachmentNames: [], images: [] });
  });
});

describe("clampThinkingLevel", () => {
  test("keeps a supported level, else picks the nearest, lower on ties", () => {
    expect(clampThinkingLevel("high", ["low", "high"])).toBe("high");
    expect(clampThinkingLevel("max", ["off", "low", "medium", "high"])).toBe("high");
    expect(clampThinkingLevel("off", ["low", "medium"])).toBe("low");
    expect(clampThinkingLevel("medium", ["low", "high"])).toBe("low");
    expect(clampThinkingLevel("medium", [])).toBeUndefined();
  });
});

test("formatTokens", () => {
  expect([formatTokens(950), formatTokens(48_200), formatTokens(1_000_000)]).toEqual([
    "950",
    "48k",
    "1M",
  ]);
});
