import {
  fauxAssistantMessage,
  fauxToolCall,
  getSystemMessageText,
  type Message,
  type TranscriptContext,
} from "@earendil-works/pi-ai";
import { describe, expect, it } from "vitest";
import type { ChatSession, ToolAccess } from "../../src/contracts.ts";
import { newChatRecord, sessionHarness } from "./fakes.ts";

function lastText(message: Message | undefined): string {
  if (!message || message.role === "system" || message.role === "assistant") return "";
  if (typeof message.content === "string") return message.content;
  return message.content.flatMap((block) => (block.type === "text" ? [block.text] : [])).join("");
}

function systemText(context: TranscriptContext | undefined): string {
  const first = context?.messages[0];
  return first?.role === "system" ? getSystemMessageText(first) : "";
}

/** Answers every approval request the moment it appears. */
function autoAnswer(
  session: ChatSession,
  decide: (access: ToolAccess) => Parameters<ChatSession["resolveApproval"]>[1],
) {
  const seen: ToolAccess[] = [];
  session.subscribe((state) => {
    for (const request of state.pendingApprovals) {
      seen.push(request.access);
      session.resolveApproval(request.id, decide(request.access));
    }
  });
  return seen;
}

describe("chat session", () => {
  it("sends the document state after the user's message and titles the chat from it", async () => {
    const harness = sessionHarness();
    const requests: TranscriptContext[] = [];
    harness.faux.setResponses([
      (context) => {
        requests.push(context);
        return fauxAssistantMessage("Done.");
      },
    ]);
    const session = harness.open(newChatRecord());

    await session.send("Tidy slide 2\nplease");

    const sent = requests[0]?.messages ?? [];
    expect(lastText(sent.at(-2))).toContain("Tidy slide 2");
    expect(lastText(sent.at(-1))).toBe("<deck_state>v0</deck_state>");
    expect(systemText(requests[0])).toContain("Host rules.");
    const state = session.getState();
    expect(state.messages.map((message) => message.role)).toEqual([
      "user",
      "documentContext",
      "assistant",
    ]);
    expect(state.title).toBe("Tidy slide 2");
    expect(harness.turns).toHaveLength(1);
  });

  it("asks before writes until the user allows writes for the chat, and refreshes state only after writes", async () => {
    const harness = sessionHarness();
    harness.faux.setResponses([
      fauxAssistantMessage(fauxToolCall("get_deck", {})),
      fauxAssistantMessage(fauxToolCall("add_slide", {})),
      fauxAssistantMessage(fauxToolCall("add_slide", {})),
      fauxAssistantMessage("Added two slides."),
    ]);
    const session = harness.open(newChatRecord());
    const asked = autoAnswer(session, () => ({ allow: true, scope: "chat" }));

    await session.send("Add two slides");

    expect(asked).toEqual(["write"]);
    const contexts = session
      .getState()
      .messages.flatMap((message) => (message.role === "documentContext" ? [message.text] : []));
    expect(contexts).toEqual([
      "<deck_state>v0</deck_state>",
      "<deck_state>v1</deck_state>",
      "<deck_state>v2</deck_state>",
    ]);
  });

  it("returns a denial with the user's comment and tells the model not to retry unchanged", async () => {
    const harness = sessionHarness();
    let toolResultSeen = "";
    harness.faux.setResponses([
      fauxAssistantMessage(fauxToolCall("add_slide", {})),
      (context) => {
        toolResultSeen = lastText(
          context.messages.findLast((message) => message.role === "toolResult"),
        );
        return fauxAssistantMessage("Understood.");
      },
    ]);
    const session = harness.open(newChatRecord());
    autoAnswer(session, () => ({ allow: false, comment: "Keep the deck at 5 slides" }));

    await session.send("Add a slide");

    expect(toolResultSeen).toContain("Keep the deck at 5 slides");
    expect(toolResultSeen).toContain("Do not retry the same call unchanged");
  });

  it("asks for code every time even after writes are allowed for the chat", async () => {
    const harness = sessionHarness();
    harness.faux.setResponses([
      fauxAssistantMessage(fauxToolCall("run_code", { code: "return 1" })),
      fauxAssistantMessage(fauxToolCall("run_code", { code: "return 2" })),
      fauxAssistantMessage("Ran it."),
    ]);
    const session = harness.open(newChatRecord({ writesAllowed: true }));
    const asked = autoAnswer(session, () => ({ allow: true, scope: "chat" }));

    await session.send("Run code twice");

    expect(asked).toEqual(["code", "code"]);
  });

  it("cancels a pending approval when the user stops the response", async () => {
    const harness = sessionHarness();
    harness.faux.setResponses([
      fauxAssistantMessage(fauxToolCall("add_slide", {})),
      // What a provider returns when called with an aborted signal.
      fauxAssistantMessage([], { stopReason: "aborted", errorMessage: "Request was aborted" }),
    ]);
    const session = harness.open(newChatRecord());
    session.subscribe((state) => {
      if (state.pendingApprovals.length > 0) session.abort();
    });

    await session.send("Add a slide");

    const state = session.getState();
    expect(state.pendingApprovals).toEqual([]);
    expect(state.isStreaming).toBe(false);
    expect(state.error).toBeUndefined();
    expect(harness.contextCalls).toEqual([0]);
    expect(
      state.messages.some((message) => message.role === "toolResult" && !message.isError),
    ).toBe(false);
  });

  it("persists the chat and reopens it with its transcript and chat-wide write permission", async () => {
    const harness = sessionHarness();
    harness.faux.setResponses([
      fauxAssistantMessage(fauxToolCall("add_slide", {})),
      fauxAssistantMessage("Done."),
    ]);
    const session = harness.open(newChatRecord());
    autoAnswer(session, () => ({ allow: true, scope: "chat" }));
    await session.send("Add a slide");

    expect(await harness.chatStore.list("doc-1")).toEqual([
      expect.objectContaining({ id: "chat-1", title: "Add a slide" }),
    ]);
    const record = await harness.chatStore.get("chat-1");
    harness.faux.setResponses([
      fauxAssistantMessage(fauxToolCall("add_slide", {})),
      fauxAssistantMessage("Again."),
    ]);
    const reopened = harness.open(record!);
    const asked = autoAnswer(reopened, () => ({ allow: true }));
    await reopened.send("One more");

    expect(asked).toEqual([]);
    expect(reopened.getState().messages.filter((message) => message.role === "user")).toHaveLength(
      2,
    );
  });

  it("compacts mid-run when the transcript passes the model's budget, then re-injects the document state", async () => {
    // Window 40k, output reserve 4k: compaction above 23k tokens. Each read returns ~4k tokens.
    const harness = sessionHarness({
      readResultChars: 16_000,
      contextWindow: 40_000,
      maxTokens: 4_000,
    });
    const isSummaryRequest = (context: TranscriptContext) =>
      systemText(context).startsWith("You summarize");
    let reads = 0;
    const summaryTranscripts: string[] = [];
    harness.faux.setResponses(
      Array.from({ length: 12 }, () => (context: TranscriptContext) => {
        if (isSummaryRequest(context)) {
          summaryTranscripts.push(lastText(context.messages.at(-1)));
          return fauxAssistantMessage("SUMMARY: user wants the deck read.");
        }
        reads += 1;
        return reads <= 7
          ? fauxAssistantMessage(fauxToolCall("get_deck", {}))
          : fauxAssistantMessage("Read it all.");
      }),
    );
    const session = harness.open(newChatRecord());

    await session.send("Read the deck seven times");

    const messages = session.getState().messages;
    const summaryIndex = messages.findIndex((message) => message.role === "compactionSummary");
    expect(summaryIndex).toBeGreaterThanOrEqual(0);
    expect(summaryTranscripts[0]).toContain("User: Read the deck seven times");
    expect(messages.slice(summaryIndex).some((message) => message.role === "documentContext")).toBe(
      true,
    );
    expect(session.getState().error).toBeUndefined();
    expect(messages.at(-1)).toMatchObject({ role: "assistant" });
    expect(session.getState().contextUsage!.tokens).toBeLessThan(23_000);
  });
});

async function untilIdle(session: ChatSession): Promise<void> {
  for (let i = 0; i < 200; i += 1) {
    const state = session.getState();
    if (!state.isStreaming && state.queuedMessages.length === 0) return;
    await new Promise((resolve) => setTimeout(resolve, 5));
  }
  throw new Error("session never went idle");
}

describe("queued messages", () => {
  it("sends a message queued during a run as the next turn once the run finishes", async () => {
    const harness = sessionHarness();
    const requests: TranscriptContext[] = [];
    let session!: ChatSession;
    harness.faux.setResponses([
      (context) => {
        requests.push(context);
        session.queue("Then shorten the title");
        expect(session.getState().queuedMessages.map((message) => message.text)).toEqual([
          "Then shorten the title",
        ]);
        return fauxAssistantMessage("First done.");
      },
      (context) => {
        requests.push(context);
        return fauxAssistantMessage("Second done.");
      },
    ]);
    session = harness.open(newChatRecord());

    await session.send("Tidy slide 2");
    await untilIdle(session);

    expect(requests).toHaveLength(2);
    expect(JSON.stringify(requests[1]!.messages.at(-1))).toContain("Then shorten the title");
    expect(harness.turns).toHaveLength(2);
  });

  it("steers the running task with a queued message instead of waiting for the run to end", async () => {
    const harness = sessionHarness();
    const requests: TranscriptContext[] = [];
    let session!: ChatSession;
    harness.faux.setResponses([
      (context) => {
        requests.push(context);
        session.queue("Use teal instead");
        session.steerQueued(session.getState().queuedMessages[0]!.id);
        return fauxAssistantMessage("Working on it.");
      },
      (context) => {
        requests.push(context);
        return fauxAssistantMessage("Switched to teal.");
      },
    ]);
    session = harness.open(newChatRecord());

    await session.send("Recolor the deck");
    await untilIdle(session);

    expect(requests).toHaveLength(2);
    expect(JSON.stringify(requests[1]!.messages)).toContain("Use teal instead");
    // Steering continues the same run: one user turn, not a second one.
    expect(harness.turns).toHaveLength(1);
  });

  it("keeps the queue when the user stops the run", async () => {
    const harness = sessionHarness();
    harness.faux.setResponses([
      fauxAssistantMessage(fauxToolCall("add_slide", {})),
      fauxAssistantMessage([], { stopReason: "aborted", errorMessage: "Request was aborted" }),
    ]);
    const session = harness.open(newChatRecord());
    session.subscribe((state) => {
      if (state.pendingApprovals.length > 0 && state.queuedMessages.length === 0) {
        session.queue("Later");
        session.abort();
      }
    });

    await session.send("Add a slide");
    await new Promise((resolve) => setTimeout(resolve, 20));

    const state = session.getState();
    expect(state.isStreaming).toBe(false);
    expect(state.queuedMessages.map((message) => message.text)).toEqual(["Later"]);
  });
});
