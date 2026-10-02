import type { AssistantMessage, ToolResultMessage, UserMessage } from "@earendil-works/pi-ai";
import { describe, expect, test } from "vitest";
import type { ApprovalRequest } from "../../src/contracts.ts";
import {
  clampThinkingLevel,
  formatTokens,
  liveAssistantMessage,
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

  test("without a result: approval, then running in the live message, else not run", () => {
    expect(toolCallStatus("c1", undefined, [approval], true)).toBe("awaiting-approval");
    expect(toolCallStatus("c2", undefined, [approval], true)).toBe("running");
    // c2 waits behind c1's approval: it hasn't started yet.
    expect(toolCallStatus("c2", undefined, [approval], true, "c1")).toBe("queued");
    expect(toolCallStatus("c1", undefined, [], true, "c1")).toBe("running");
    expect(toolCallStatus("c2", undefined, [], false)).toBe("not-run");
  });
});

describe("liveAssistantMessage", () => {
  const user = (text: string): UserMessage => ({ role: "user", content: text, timestamp: 0 });
  const assistant = {
    role: "assistant",
    content: [],
    stopReason: "toolUse",
  } as unknown as AssistantMessage;

  test("is the newest assistant message of the running run", () => {
    expect(liveAssistantMessage({ messages: [user("a"), assistant], isStreaming: true })).toBe(
      assistant,
    );
    const streaming = { ...assistant };
    expect(
      liveAssistantMessage({
        messages: [user("a"), assistant],
        streamingMessage: streaming,
        isStreaming: true,
      }),
    ).toBe(streaming);
  });

  test("leaves a stopped run's unfinished calls out of later runs", () => {
    expect(
      liveAssistantMessage({ messages: [user("a"), assistant, user("b")], isStreaming: true }),
    ).toBeUndefined();
    expect(
      liveAssistantMessage({ messages: [user("a"), assistant], isStreaming: false }),
    ).toBeUndefined();
  });
});

describe("splitUserContent", () => {
  test("separates typed text, attachment excerpts and images", () => {
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
        image,
      ]),
    ).toEqual({
      text: "Fix slide 2",
      attachmentNames: ["report.pdf", 'say "hi".txt'],
      images: [image],
    });
  });

  test("shows attachment-like or context-like typed text instead of hiding it or throwing", () => {
    const typed = '<attachment name="C:\\presentations\\deck.pptx">';
    expect(splitUserContent([{ type: "text", text: typed }]).text).toBe(typed);
    expect(splitUserContent([{ type: "text", text: "<context> means…" }]).text).toBe(
      "<context> means…",
    );
    expect(
      splitUserContent([
        { type: "text", text: "Fix it" },
        { type: "text", text: typed },
      ]),
    ).toEqual({ text: `Fix it\n\n${typed}`, attachmentNames: [], images: [] });
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
