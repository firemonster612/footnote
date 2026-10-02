import {
  fauxAssistantMessage,
  fauxToolCall,
  getSystemMessageText,
  type Message,
  type TranscriptContext,
} from "@earendil-works/pi-ai";
import { describe, expect, it } from "vitest";
import type { ChatSession, SkillDefinition, ToolAccess, ToolEnv } from "../../src/contracts.ts";
import { fakeHost, newChatRecord, sessionHarness } from "./fakes.ts";

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

describe("revert", () => {
  it("undoes the request's turn and every later one, drops them from the chat, and returns the text", async () => {
    const harness = sessionHarness();
    harness.faux.setResponses([
      fauxAssistantMessage("One."),
      fauxAssistantMessage("Two."),
      fauxAssistantMessage("Three."),
    ]);
    const session = harness.open(newChatRecord());
    await session.send("First request");
    await session.send("Second request");
    await session.send("Third request");
    const requests = session.getState().messages.filter((message) => message.role === "user");

    const { text } = await session.revertTo(requests[1]!.timestamp);

    expect(text).toBe("Second request");
    expect(harness.undoneTurnIds).toEqual([harness.turns.slice(1)]);
    const remaining = session.getState().messages;
    expect(remaining.filter((message) => message.role === "user")).toHaveLength(1);
    expect(remaining.at(-1)?.role).toBe("assistant");

    // The next send starts fresh from there, and a later revert doesn't touch the reverted turns again.
    await session.send("Replacement");
    const replacement = session.getState().messages.findLast((message) => message.role === "user")!;
    await session.revertTo(replacement.timestamp);
    expect(harness.undoneTurnIds.at(-1)).toEqual([harness.turns.at(-1)]);
  });
});

const userTexts = (session: ChatSession) =>
  session
    .getState()
    .messages.flatMap((message) => (message.role === "user" ? [lastText(message)] : []));

/** A promise the test settles later. */
function gate<T = void>() {
  let open!: (value: T) => void;
  const opened = new Promise<T>((resolve) => (open = resolve));
  return { open, opened };
}

describe("operations that overlap", () => {
  it("queues a message sent while a revert restores the document, then sends it", async () => {
    const harness = sessionHarness();
    harness.faux.setResponses([
      fauxAssistantMessage("One."),
      fauxAssistantMessage("Two."),
      fauxAssistantMessage("Meanwhile done."),
    ]);
    const session = harness.open(newChatRecord());
    await session.send("First");
    await session.send("Second");
    const undo = gate();
    harness.hostModule.undo.undoTurns = async (_env, _chatId, turnIds) => {
      await undo.opened;
      return { restored: turnIds.length, removed: 0, warnings: [] };
    };

    const reverting = session.revertTo(session.getState().revertibleRequests[1]!);
    void session.send("Meanwhile");
    expect(session.getState().queuedMessages.map((message) => message.text)).toEqual(["Meanwhile"]);
    undo.open();
    await reverting;
    await untilIdle(session);

    expect(userTexts(session)).toEqual(["First", "Meanwhile"]);
  });

  it("puts a request stopped before its run started back at the front of the queue", async () => {
    const harness = sessionHarness();
    const model = gate<Awaited<ReturnType<typeof harness.client.resolveModel>>>();
    const resolveModel = harness.client.resolveModel;
    harness.client.resolveModel = () => model.opened;
    const session = harness.open(newChatRecord());

    const sending = session.send("Tidy slide 2");
    session.queue("Later");
    session.abort();
    model.open(await resolveModel("faux-model"));
    await sending;
    await new Promise((resolve) => setTimeout(resolve, 20));

    const state = session.getState();
    expect(state.isStreaming).toBe(false);
    expect(state.queuedMessages.map((message) => message.text)).toEqual(["Tidy slide 2", "Later"]);
    expect(state.messages).toEqual([]);
    expect(state.revertibleRequests).toEqual([]);
    expect(harness.faux.state.callCount).toBe(0);
  });

  it("moves a steer the model never received back to the queue when the user stops", async () => {
    const harness = sessionHarness();
    const requests: TranscriptContext[] = [];
    let session!: ChatSession;
    harness.faux.setResponses([
      () => {
        session.steer("Use teal");
        session.abort();
        return fauxAssistantMessage([], {
          stopReason: "aborted",
          errorMessage: "Request was aborted",
        });
      },
      (context) => {
        requests.push(context);
        return fauxAssistantMessage("Done.");
      },
    ]);
    session = harness.open(newChatRecord());

    await session.send("Recolor the deck");
    expect(session.getState().queuedMessages.map((message) => message.text)).toEqual(["Use teal"]);
    session.removeQueued(session.getState().queuedMessages[0]!.id);
    await session.send("Never mind");

    expect(JSON.stringify(requests[0]!.messages)).not.toContain("Use teal");
  });

  it("stops saving once disposed, so a deleted chat stays deleted", async () => {
    const harness = sessionHarness();
    const reply = gate();
    harness.faux.setResponses([
      async () => {
        await reply.opened;
        return fauxAssistantMessage("Done.");
      },
    ]);
    const session = harness.open(newChatRecord());

    const sending = session.send("Add a slide");
    await new Promise((resolve) => setTimeout(resolve, 0));
    session.dispose();
    await harness.chatStore.delete("chat-1");
    reply.open();
    await sending;
    await new Promise((resolve) => setTimeout(resolve, 0));

    expect(await harness.chatStore.get("chat-1")).toBeUndefined();
  });
});

describe("saving", () => {
  it("saves the chat as soon as a run starts, so a reopened panel lists it", async () => {
    const harness = sessionHarness();
    let listed: unknown[] = [];
    harness.faux.setResponses([
      async () => {
        await new Promise((resolve) => setTimeout(resolve, 0));
        listed = await harness.chatStore.list("doc-1");
        return fauxAssistantMessage("Done.");
      },
    ]);
    const session = harness.open(newChatRecord());

    await session.send("Add a slide");

    expect(listed).toEqual([expect.objectContaining({ id: "chat-1", title: "Add a slide" })]);
  });

  it("stores the emptied chat when the first request is reverted", async () => {
    const harness = sessionHarness();
    harness.faux.setResponses([fauxAssistantMessage("Done.")]);
    const session = harness.open(newChatRecord());
    await session.send("First");

    await session.revertTo(session.getState().revertibleRequests[0]!);

    expect((await harness.chatStore.get("chat-1"))?.messages).toEqual([]);
  });
});

describe("tools", () => {
  it("blocks document tools when the chat's tab now shows another document", async () => {
    const harness = sessionHarness({
      officeHost: {
        ...fakeHost,
        status: async () => ({ connected: true, host: "powerpoint", documentId: "doc-2" }),
      },
    });
    harness.faux.setResponses([
      fauxAssistantMessage(fauxToolCall("add_slide", {})),
      fauxAssistantMessage("Stopped."),
    ]);
    const session = harness.open(newChatRecord({ writesAllowed: true }));

    await session.send("Add a slide");

    const result = session.getState().messages.find((message) => message.role === "toolResult");
    expect(result).toMatchObject({ isError: true });
    expect(JSON.stringify(result)).toContain("different document");
  });

  it("gives host tools an env bound to the chat", () => {
    const harness = sessionHarness();
    const createTools = harness.hostModule.createTools;
    let env: ToolEnv | undefined;
    harness.hostModule.createTools = (toolEnv) => {
      env = toolEnv;
      return createTools(toolEnv);
    };

    harness.open(newChatRecord());

    expect(env?.chatId).toBe("chat-1");
  });

  it("labels tool calls with the tool's label when it can't describe the call", async () => {
    const harness = sessionHarness();
    harness.faux.setResponses([
      fauxAssistantMessage(fauxToolCall("get_deck", {}, { id: "call-1" })),
      fauxAssistantMessage("Read it."),
    ]);
    const session = harness.open(newChatRecord());

    await session.send("Read the deck");

    expect(session.getState().toolCallLabels).toEqual({ "call-1": "get_deck" });
  });

  it("lists skills added after the chat opened", async () => {
    const skills: SkillDefinition[] = [];
    const harness = sessionHarness({ skills: () => skills });
    const requests: TranscriptContext[] = [];
    harness.faux.setResponses(
      Array.from({ length: 2 }, () => (context: TranscriptContext) => {
        requests.push(context);
        return fauxAssistantMessage("Done.");
      }),
    );
    const session = harness.open(newChatRecord());
    await session.send("First");

    skills.push({
      name: "pitch-deck",
      description: "Build pitch decks.",
      body: "",
      source: "custom",
    });
    await session.send("Second");

    expect(systemText(requests[0])).not.toContain("pitch-deck");
    expect(systemText(requests[1])).toContain("pitch-deck");
  });
});
