import type { ImageContent, TextContent, ToolResultMessage } from "@earendil-works/pi-ai";
import { describe, expect, it } from "vitest";
import {
  compactionThresholdTokens,
  compactMessages,
  createCompactionBreaker,
  findCutIndex,
} from "../../src/compaction/compaction.ts";
import type { AgentMessage } from "../../src/contracts.ts";
import { toLlmMessages } from "../../src/context/messages.ts";
import { capToolResultText } from "../../src/context/toolResultBudget.ts";
import { needsApproval } from "../../src/permissions/permissionGate.ts";

const image: ImageContent = { type: "image", data: "AAAA", mimeType: "image/png" };

function toolResult(
  content: (TextContent | ImageContent)[],
  toolCallId = "call",
): ToolResultMessage {
  return {
    role: "toolResult",
    toolCallId,
    toolName: "render_slide",
    content,
    isError: false,
    timestamp: 0,
  };
}

function assistantCall(id: string): AgentMessage {
  return {
    role: "assistant",
    content: [{ type: "toolCall", id, name: "render_slide", arguments: {} }],
    api: "faux",
    provider: "faux",
    model: "m",
    usage: {
      input: 0,
      output: 0,
      cacheRead: 0,
      cacheWrite: 0,
      totalTokens: 0,
      cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 },
    },
    stopReason: "toolUse",
    timestamp: 0,
  };
}

describe("permissions", () => {
  it.each([
    ["read", "ask", false, false],
    ["web", "ask", false, false],
    ["write", "ask", false, true],
    ["write", "ask", true, false],
    ["code", "ask", true, true],
    ["code", "full", false, false],
    ["write", "full", false, false],
  ] as const)(
    "%s in %s mode (writes allowed: %s) asks: %s",
    (access, mode, writesAllowed, asks) => {
      expect(needsApproval(access, mode, writesAllowed)).toBe(asks);
    },
  );
});

describe("tool result budget", () => {
  it("caps text with a truncation note and keeps images", () => {
    const capped = capToolResultText([{ type: "text", text: "a".repeat(30) }, image], 20);
    expect(capped).toEqual([
      { type: "text", text: "a".repeat(20) },
      image,
      { type: "text", text: expect.stringContaining("10 of 30 characters omitted") },
    ]);
  });

  it("leaves results within budget alone", () => {
    expect(capToolResultText([{ type: "text", text: "short" }], 20)).toBeUndefined();
  });
});

describe("messages sent to the model", () => {
  it("keeps every image from the current turn and only the newest earlier ones", () => {
    const renders = (count: number) =>
      Array.from({ length: count }, (_, n) =>
        toolResult([{ type: "text", text: `render ${n}` }, image]),
      );
    const userMessage = {
      role: "user",
      content: [{ type: "text", text: "next" }],
      timestamp: 1,
    } as AgentMessage;
    const sent = toLlmMessages([...renders(4), userMessage, ...renders(4)], 3);
    const imageCounts = sent.map((message) =>
      typeof message.content === "string" || message.role === "system"
        ? 0
        : message.content.filter((block) => block.type === "image").length,
    );
    expect(imageCounts).toEqual([0, 1, 1, 1, 0, 1, 1, 1, 1]);
    expect(JSON.stringify(sent[0])).toContain("Earlier image removed");
  });

  it("drops all images for models without image input", () => {
    const userMessage = {
      role: "user",
      content: [{ type: "text", text: "go" }],
      timestamp: 1,
    } as AgentMessage;
    const sent = toLlmMessages([userMessage, toolResult([image])], 0);
    expect(JSON.stringify(sent)).not.toContain('"type":"image"');
  });

  it("sends document state and compaction summaries as user text", () => {
    const sent = toLlmMessages(
      [
        { role: "compactionSummary", summary: "Earlier work", timestamp: 1 },
        { role: "documentContext", text: "<deck_state/>", timestamp: 2 },
      ],
      3,
    );
    expect(sent.map((message) => message.role)).toEqual(["user", "user"]);
    expect(JSON.stringify(sent)).toContain("Earlier work");
    expect(JSON.stringify(sent)).toContain("<deck_state/>");
  });
});

describe("compaction", () => {
  it("triggers at window minus capped output reserve minus 13k", () => {
    expect(compactionThresholdTokens(200_000, 64_000)).toBe(167_000);
    expect(compactionThresholdTokens(40_000, 4_000)).toBe(23_000);
  });

  it("never cuts between a tool call and its result", () => {
    const big = "x".repeat(4_000);
    const messages: AgentMessage[] = [
      { role: "system", content: "prompt", timestamp: 0 },
      { role: "user", content: "start", timestamp: 0 },
      assistantCall("a"),
      toolResult([{ type: "text", text: big }], "a"),
      assistantCall("b"),
      toolResult([{ type: "text", text: big }], "b"),
    ];
    for (const keep of [1, 500, 1_000, 1_500, 2_500]) {
      const cut = findCutIndex(messages, keep);
      if (cut !== undefined) expect(messages[cut]?.role).not.toBe("toolResult");
    }
    expect(findCutIndex(messages, 1_000)).toBe(4);
  });

  it("replaces the head with a summary and keeps the system message and the tail", async () => {
    const messages: AgentMessage[] = [
      { role: "system", content: "prompt", timestamp: 0 },
      { role: "user", content: "old request", timestamp: 0 },
      { role: "user", content: "x".repeat(400), timestamp: 0 },
    ];
    const compacted = await compactMessages(
      messages,
      50,
      async (transcript) => `summary of: ${transcript}`,
    );
    expect(compacted?.map((message) => message.role)).toEqual([
      "system",
      "compactionSummary",
      "user",
    ]);
    expect(compacted?.[1]).toMatchObject({ summary: "summary of: User: old request" });
  });

  it("stops after three compactions within three turns", () => {
    const breaker = createCompactionBreaker();
    for (const turn of [1, 2, 3]) {
      expect(breaker.tripped(turn)).toBe(false);
      breaker.record(turn);
    }
    expect(breaker.tripped(3)).toBe(true);
    expect(breaker.tripped(4)).toBe(false);
  });
});

describe("toLlmMessages with a broken image", () => {
  it("replaces image blocks whose data isn't a string so the provider accepts the request", () => {
    const broken = {
      role: "toolResult",
      toolCallId: "t1",
      toolName: "render_slide",
      content: [
        { type: "text", text: "Slide 1:" },
        { type: "image", data: null, mimeType: "image/png" },
      ],
      isError: false,
      timestamp: 1,
    } as unknown as AgentMessage;
    const [converted] = toLlmMessages([broken], 3);
    const content = converted!.content as { type: string; text?: string }[];
    expect(content.map((block) => block.type)).toEqual(["text", "text"]);
    expect(content[1]!.text).toContain("Image unavailable");
  });
});
