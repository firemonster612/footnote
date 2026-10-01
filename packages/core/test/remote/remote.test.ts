import { fauxAssistantMessage, fauxToolCall } from "@earendil-works/pi-ai";
import { describe, expect, it } from "vitest";
import type { ChatSession, FootnoteApp } from "../../src/contracts.ts";
import { connectFootnoteApp } from "../../src/remote/client.ts";
import type { RemotePort } from "../../src/remote/protocol.ts";
import { serveFootnoteApp } from "../../src/remote/server.ts";
import { loadSettings } from "../../src/settings/settings.ts";
import { fakeHost, memoryStore, newChatRecord, sessionHarness } from "../runtime/fakes.ts";

/** Two connected ends that deliver asynchronously and JSON-copy every message, like a chrome.runtime port. */
function portPair() {
  const end = () => ({
    messageListeners: [] as ((message: any) => void)[],
    disconnectListeners: [] as (() => void)[],
  });
  const a = end();
  const b = end();
  const port = (self: typeof a, other: typeof a): RemotePort<any, any> => ({
    postMessage(message) {
      const copy: unknown = JSON.parse(JSON.stringify(message));
      setTimeout(() => other.messageListeners.forEach((listener) => listener(copy)));
    },
    onMessage: { addListener: (listener) => void self.messageListeners.push(listener) },
    onDisconnect: { addListener: (listener) => void self.disconnectListeners.push(listener) },
  });
  return { server: port(a, b), client: port(b, a) };
}

async function servedApp() {
  const harness = sessionHarness();
  const settings = await loadSettings(memoryStore());
  const sessions = new Map<string, ChatSession>();
  const app: FootnoteApp = {
    host: fakeHost,
    hostModule: harness.hostModule,
    settings,
    models: { list: async () => [] },
    chats: {
      list: async () => [],
      async create() {
        const session = harness.open(newChatRecord());
        sessions.set(session.getState().id, session);
        return session;
      },
      open: async (chatId) => sessions.get(chatId)!,
      delete: async () => {},
    },
    skills: { list: () => [] },
  };
  const ports = portPair();
  void serveFootnoteApp(app, ports.server);
  const remote = await connectFootnoteApp(ports.client, harness.hostModule);
  return { harness, app, sessions, remote };
}

describe("remote app", () => {
  it("mirrors a run's state, including the streamed reply and final transcript", async () => {
    const { harness, sessions, remote } = await servedApp();
    harness.faux.setResponses([fauxAssistantMessage("Hello from the engine.")]);
    const session = await remote.chats.create("doc-1");
    const seen: boolean[] = [];
    session.subscribe((state) => seen.push(state.isStreaming));

    await session.send("Hi");

    const served = sessions.get(session.getState().id)!.getState();
    expect(session.getState()).toEqual(JSON.parse(JSON.stringify(served)));
    expect(session.getState().messages.map((message) => message.role)).toEqual([
      "user",
      "documentContext",
      "assistant",
    ]);
    expect(seen).toContain(true);
    expect(seen.at(-1)).toBe(false);
  });

  it("shows approvals raised in the engine and resolves them from the view", async () => {
    const { harness, remote } = await servedApp();
    harness.faux.setResponses([
      fauxAssistantMessage(fauxToolCall("add_slide", {})),
      fauxAssistantMessage("Added it."),
    ]);
    const session = await remote.chats.create("doc-1");
    const asked: string[] = [];
    session.subscribe((state) => {
      for (const request of state.pendingApprovals) {
        if (asked.includes(request.id)) continue;
        asked.push(request.id);
        session.resolveApproval(request.id, { allow: true });
      }
    });

    await session.send("Add a slide");

    expect(asked).toHaveLength(1);
    const messages = session.getState().messages;
    expect(messages.find((message) => message.role === "toolResult")).toMatchObject({
      toolName: "add_slide",
      isError: false,
    });
    expect(session.getState().pendingApprovals).toEqual([]);
  });

  it("sends staged files' contents to the engine", async () => {
    const { sessions, remote } = await servedApp();
    const session = await remote.chats.create("doc-1");

    const meta = await session.stageAttachment(
      new File(["slide notes"], "notes.txt", { type: "text/plain" }),
    );

    expect(meta).toMatchObject({ name: "notes.txt", kind: "text", size: 11 });
    const served = sessions.get(session.getState().id)!.getState();
    expect(served.stagedAttachments.map((attachment) => attachment.name)).toEqual(["notes.txt"]);
  });

  it("round-trips calls, including the engine's errors and transcript truncation", async () => {
    const { harness, app, remote } = await servedApp();
    harness.faux.setResponses([fauxAssistantMessage("One."), fauxAssistantMessage("Two.")]);
    const session = await remote.chats.create("doc-1");
    await session.send("First");
    await session.send("Second");

    await remote.settings.update({ firecrawlApiKey: "fc-key" });
    const [, second] = session.getState().revertibleRequests;
    const { text } = await session.revertTo(second!);

    expect(app.settings.get().firecrawlApiKey).toBe("fc-key");
    expect(remote.settings.get().firecrawlApiKey).toBe("fc-key");
    expect(text).toBe("Second");
    expect(session.getState().messages.map((message) => message.role)).toEqual([
      "user",
      "documentContext",
      "assistant",
    ]);
    await expect(session.revertTo(12345)).rejects.toThrow("no longer in this chat");
  });
});
