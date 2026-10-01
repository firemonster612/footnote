// In-memory FootnoteApp for the standalone UI dev page and UI tests. Not part of the package exports.

import type {
  AssistantMessage,
  ImageContent,
  ToolCall,
  ToolResultMessage,
  UserMessage,
} from "@earendil-works/pi-ai";
import { Type } from "typebox";
import type {
  AgentMessage,
  ApprovalDecision,
  AttachmentMeta,
  ChatSession,
  ChatSessionState,
  ChatSummary,
  FootnoteApp,
  FootnoteTool,
  HostModule,
  ModelInfo,
  OfficeHost,
  OfficeHostStatus,
  Settings,
  ToolAccess,
} from "../../contracts.ts";

export interface FakeAppOptions {
  /** Start with no endpoint configured. */
  firstRun?: boolean;
  /** Drop the host connection shortly after start. */
  disconnect?: boolean;
}

const documentId = "doc-q3-review";
const disconnectedStatus: OfficeHostStatus = {
  connected: false,
  reason: "Open the Claude add-in in PowerPoint to connect",
};

export function createFakeApp(options: FakeAppOptions = {}): FootnoteApp {
  let settings: Settings = {
    endpoint: options.firstRun
      ? { baseUrl: "", apiKey: "" }
      : { baseUrl: "http://localhost:8317", apiKey: "sk-demo" },
    defaultPermissionMode: "ask",
    modelId: "claude-opus-5-5",
    thinkingLevel: "high",
    apiOverrides: {},
    customSkills: [],
  };
  const settingsListeners = new Set<(settings: Settings) => void>();
  const sessions = new Map<string, FakeChatSession>();
  const demo = new FakeChatSession("chat-demo", demoState());
  sessions.set(demo.id, demo);
  const summaries: ChatSummary[] = [
    { id: demo.id, documentId, title: demo.getState().title, updatedAt: Date.now() - 60_000 },
    {
      id: "chat-old",
      documentId,
      title: "Rebuild the agenda slide",
      updatedAt: Date.now() - 3 * 86_400_000,
    },
  ];

  return {
    host: createFakeHost(options.disconnect === true),
    hostModule: fakeHostModule,
    settings: {
      get: () => settings,
      update: async (patch) => {
        settings = { ...settings, ...patch };
        for (const listener of settingsListeners) listener(settings);
      },
      subscribe: (listener) => {
        settingsListeners.add(listener);
        return () => settingsListeners.delete(listener);
      },
    },
    models: {
      list: async () => {
        await delay(300);
        if (!settings.endpoint.baseUrl) throw new Error("Set an endpoint URL first.");
        return fakeModels.map((model) => ({
          ...model,
          api: settings.apiOverrides[model.id] ?? model.api,
        }));
      },
    },
    chats: {
      list: async (id) => summaries.filter((chat) => chat.documentId === id),
      create: async (id) => {
        const chatId = `chat-${Date.now()}`;
        const session = new FakeChatSession(chatId, emptyState(chatId, settings));
        sessions.set(session.id, session);
        summaries.push({ id: chatId, documentId: id, title: "New chat", updatedAt: Date.now() });
        return session;
      },
      open: async (chatId) => {
        const session =
          sessions.get(chatId) ?? new FakeChatSession(chatId, emptyState(chatId, settings));
        sessions.set(chatId, session);
        return session;
      },
      delete: async (chatId) => {
        sessions.delete(chatId);
        summaries.splice(0, summaries.length, ...summaries.filter((chat) => chat.id !== chatId));
      },
    },
    skills: { list: () => [] },
  };
}

function createFakeHost(disconnect: boolean): OfficeHost {
  let status: OfficeHostStatus = {
    connected: true,
    host: "powerpoint",
    documentId,
    documentName: "Q3 Review.pptx",
  };
  const listeners = new Set<(status: OfficeHostStatus) => void>();
  if (disconnect) {
    setTimeout(() => {
      status = disconnectedStatus;
      for (const listener of listeners) listener(status);
    }, 500);
  }
  return {
    status: async () => status,
    onStatusChange: (listener) => {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    call: async () => {
      throw new Error("The fake host runs no ops.");
    },
    runCode: async () => ({
      ok: false,
      logs: [],
      error: { message: "The fake host runs no code." },
    }),
  };
}

const fakeModels: ModelInfo[] = [
  {
    id: "claude-opus-5-5",
    ownedBy: "anthropic",
    api: "anthropic-messages",
    thinkingLevels: ["off", "low", "medium", "high", "xhigh", "max"],
    contextWindow: 1_000_000,
    supportsImages: true,
  },
  {
    id: "gpt-6.1-sol",
    ownedBy: "openai",
    api: "openai-responses",
    thinkingLevels: ["low", "medium", "high", "xhigh", "max"],
    contextWindow: 400_000,
    supportsImages: true,
  },
  {
    id: "qwen3-coder",
    ownedBy: "alibaba",
    api: "openai-completions",
    thinkingLevels: ["off", "low", "medium", "high"],
    contextWindow: 128_000,
    supportsImages: false,
  },
];

function fakeTool(
  name: string,
  label: string,
  access: ToolAccess,
  describeCall: (args: any) => string,
): FootnoteTool {
  return {
    name,
    label,
    access,
    describeCall,
    description: label,
    parameters: Type.Object({}),
    execute: async () => {
      throw new Error("Fake tools don't run.");
    },
  };
}

const fakeHostModule: HostModule = {
  kind: "powerpoint",
  ops: {},
  systemPrompt: "",
  createTools: () => [
    fakeTool("get_slide", "Read slide", "read", (args) => `Read slide ${args.slideId}`),
    fakeTool("render_slide", "Render slide", "read", (args) => `Render slide ${args.slideId}`),
    fakeTool(
      "update_shapes",
      "Update shapes",
      "write",
      (args) => `Update ${args.updates.length} shapes on slide ${args.slideId}`,
    ),
    fakeTool("execute_office_js", "Run Office.js", "code", (args) => args.explanation),
  ],
  getContextBlock: async () => "",
  undo: {
    beginTurn: () => {},
    canUndo: () => false,
    undoLastTurn: async () => ({ restored: 0, removed: 0, warnings: [] }),
  },
  skills: [],
};

class FakeChatSession implements ChatSession {
  private listeners = new Set<(state: ChatSessionState) => void>();
  private timers: ReturnType<typeof setTimeout>[] = [];

  constructor(
    readonly id: string,
    private state: ChatSessionState,
  ) {}

  getState = () => this.state;

  subscribe(listener: (state: ChatSessionState) => void) {
    this.listeners.add(listener);
    return () => {
      this.listeners.delete(listener);
    };
  }

  async send(text: string) {
    if (this.state.isStreaming) return this.queue(text);
    const attachments = this.state.stagedAttachments.map(
      (attachment) =>
        ({
          type: "text",
          text: `<attachment id="${attachment.id}" name="${attachment.name}">…</attachment>`,
        }) as const,
    );
    this.update({
      messages: [...this.state.messages, userMessage([{ type: "text", text }, ...attachments])],
      title: this.state.messages.length === 0 ? text.slice(0, 60) : this.state.title,
      stagedAttachments: [],
      isStreaming: true,
      error: undefined,
    });
    this.streamReply();
  }

  steer(text: string) {
    this.update({ messages: [...this.state.messages, userMessage([{ type: "text", text }])] });
  }

  queue(text: string) {
    if (!this.state.isStreaming) return void this.send(text);
    this.update({
      queuedMessages: [
        ...this.state.queuedMessages,
        { id: crypto.randomUUID(), text, attachments: this.state.stagedAttachments },
      ],
      stagedAttachments: [],
    });
  }

  removeQueued(id: string) {
    this.update({
      queuedMessages: this.state.queuedMessages.filter((message) => message.id !== id),
    });
  }

  steerQueued(id: string) {
    const message = this.state.queuedMessages.find((candidate) => candidate.id === id);
    if (!message) return;
    this.removeQueued(id);
    if (this.state.isStreaming) this.steer(message.text);
    else void this.send(message.text);
  }

  /** Mirrors the runtime: a finished run sends the oldest queued message. */
  private finishRun(patch: Partial<ChatSessionState>) {
    this.update({ ...patch, isStreaming: false });
    const [next] = this.state.queuedMessages;
    if (!next) return;
    this.update({ queuedMessages: this.state.queuedMessages.slice(1) });
    void this.send(next.text);
  }

  abort() {
    for (const timer of this.timers) clearTimeout(timer);
    this.timers = [];
    const partial = this.state.streamingMessage;
    const messages =
      partial?.role === "assistant"
        ? [...this.state.messages, { ...partial, stopReason: "aborted" as const }]
        : this.state.messages;
    this.update({
      messages,
      streamingMessage: undefined,
      isStreaming: false,
      pendingApprovals: [],
    });
  }

  resolveApproval(approvalId: string, decision: ApprovalDecision) {
    const request = this.state.pendingApprovals.find((approval) => approval.id === approvalId);
    if (!request) return;
    const text = decision.allow
      ? "Updated 3 shapes. Verified bounds and text."
      : `The user denied this call.${decision.comment ? ` Comment: ${decision.comment}` : ""} Don't retry it unchanged.`;
    const pendingApprovals = this.state.pendingApprovals.filter((approval) => approval !== request);
    this.update({
      messages: [
        ...this.state.messages,
        toolResult(request.toolCallId, request.toolName, [{ type: "text", text }], !decision.allow),
      ],
      pendingApprovals,
    });
    if (pendingApprovals.length === 0) this.streamReply();
  }

  setPermissionMode(permissionMode: ChatSessionState["permissionMode"]) {
    this.update({ permissionMode });
  }

  setModel(modelId: string) {
    this.update({ modelId });
  }

  setThinkingLevel(thinkingLevel: ChatSessionState["thinkingLevel"]) {
    this.update({ thinkingLevel });
  }

  async stageAttachment(file: File): Promise<AttachmentMeta> {
    if (file.size > 50 * 1024 * 1024) throw new Error("Files over 50 MB can't be attached.");
    const meta: AttachmentMeta = {
      id: `att-${Date.now()}`,
      name: file.name,
      mimeType: file.type,
      kind: file.type.startsWith("image/") ? "image" : "text",
      size: file.size,
      summary: `${Math.max(1, Math.round(file.size / 1024))} KB`,
    };
    this.update({ stagedAttachments: [...this.state.stagedAttachments, meta] });
    return meta;
  }

  unstageAttachment(id: string) {
    this.update({
      stagedAttachments: this.state.stagedAttachments.filter((attachment) => attachment.id !== id),
    });
  }

  async undoLastTurn() {
    await delay(400);
    this.update({ canUndo: false });
    return {
      restored: 2,
      removed: 1,
      warnings: ["Slide 4 changed after the turn; those edits were replaced."],
    };
  }

  private update(patch: Partial<ChatSessionState>) {
    this.state = { ...this.state, ...patch };
    for (const listener of this.listeners) listener(this.state);
  }

  /** Streams thinking, then text word by word, then a render tool call, then finishes. */
  private streamReply() {
    const words =
      "Done. The title now uses the theme heading font at **32 pt**, and the body text fits inside its placeholder.".split(
        " ",
      );
    const call: ToolCall = {
      type: "toolCall",
      id: `call-${Date.now()}`,
      name: "render_slide",
      arguments: { slideId: "258" },
    };
    const steps: (() => void)[] = [
      () =>
        this.update({
          isStreaming: true,
          streamingMessage: assistantMessage([
            { type: "thinking", thinking: "Checking the result against the render." },
          ]),
        }),
      ...words.map(
        (_, index) => () =>
          this.update({
            streamingMessage: assistantMessage([
              { type: "thinking", thinking: "Checking the result against the render." },
              { type: "text", text: words.slice(0, index + 1).join(" ") },
            ]),
          }),
      ),
      () => {
        const message = this.state.streamingMessage;
        if (message?.role !== "assistant") return;
        this.update({
          messages: [...this.state.messages, { ...message, content: [...message.content, call] }],
          streamingMessage: undefined,
        });
      },
      () =>
        this.finishRun({
          messages: [
            ...this.state.messages,
            toolResult(call.id, call.name, [slideImage("Q3 revenue", "#0F766E")]),
          ],
          canUndo: true,
          contextUsage: {
            tokens: (this.state.contextUsage?.tokens ?? 0) + 3_100,
            window: this.state.contextUsage?.window ?? 200_000,
          },
        }),
    ];
    this.timers = steps.map((step, index) => setTimeout(step, 200 + index * 90));
  }
}

function emptyState(id: string, settings: Settings): ChatSessionState {
  return {
    id,
    title: "New chat",
    messages: [],
    isStreaming: false,
    pendingApprovals: [],
    permissionMode: settings.defaultPermissionMode,
    modelId: settings.modelId,
    thinkingLevel: settings.thinkingLevel,
    stagedAttachments: [],
    queuedMessages: [],
    canUndo: false,
  };
}

/** Every message, tool and approval state the UI renders. */
function demoState(): ChatSessionState {
  const getSlide: ToolCall = {
    type: "toolCall",
    id: "call-1",
    name: "get_slide",
    arguments: { slideId: "257" },
  };
  const render: ToolCall = {
    type: "toolCall",
    id: "call-2",
    name: "render_slide",
    arguments: { slideId: "257" },
  };
  const staleWrite: ToolCall = {
    type: "toolCall",
    id: "call-3",
    name: "update_shapes",
    arguments: { slideId: "257", updates: [{ shapeId: "4", text: "Revenue grew 18%" }] },
  };
  const write: ToolCall = {
    type: "toolCall",
    id: "call-4",
    name: "update_shapes",
    arguments: {
      slideId: "257",
      updates: [
        { shapeId: "4", text: "Revenue grew 18%" },
        { shapeId: "5", fontSize: 18 },
        { shapeId: "7", delete: true },
      ],
    },
  };
  const code = `const slide = footnote.slide("257");\nconst shapes = slide.shapes;\nshapes.load("items/name,items/left,items/width");\nawait context.sync();\nfor (const shape of shapes.items) shape.left = footnote.pt(48);\nawait context.sync();\nreturn shapes.items.length;`;
  const runCode: ToolCall = {
    type: "toolCall",
    id: "call-5",
    name: "execute_office_js",
    arguments: { explanation: "Align every shape on slide 3 to the left margin", code },
  };
  const pendingRender: ToolCall = {
    type: "toolCall",
    id: "call-6",
    name: "render_slide",
    arguments: { slideId: "258" },
  };

  const messages: AgentMessage[] = [
    {
      role: "compactionSummary",
      summary:
        "The user is reworking the Q3 review deck. Slides 1–2 are done; slide 3 (ID 257) still needs the revenue numbers.",
      timestamp: Date.now(),
    },
    userMessage([
      { type: "text", text: "Tighten slide 3 and add the revenue numbers from this file." },
      {
        type: "text",
        text: '<attachment id="att-1" name="Q3 revenue.xlsx">| Region | Q3 |…</attachment>',
      },
      slideImage("Whiteboard sketch", "#64748b"),
    ]),
    { role: "documentContext", text: "<deck_state>slides: 12</deck_state>", timestamp: Date.now() },
    assistantMessage([
      {
        type: "thinking",
        thinking:
          "Slide 3 is the revenue summary. Read its shape tree and render it before changing anything.",
      },
      { type: "text", text: "I'll read slide 3 first." },
      getSlide,
      render,
    ]),
    toolResult(getSlide.id, getSlide.name, [
      {
        type: "text",
        text: '{"slideId":"257","shapes":[{"id":"4","name":"Title 1","text":"Revenue"},{"id":"5","name":"Content 2"}]}',
      },
    ]),
    toolResult(render.id, render.name, [slideImage("Revenue", "#0F766E")]),
    assistantMessage([{ type: "text", text: "Applying the layout changes." }], {
      stopReason: "error",
      errorMessage: "Provider returned 529: overloaded. Try again.",
    }),
    userMessage([{ type: "text", text: "Try again, and keep the title font." }]),
    assistantMessage([
      {
        type: "text",
        text: [
          "### Plan",
          "",
          "1. Replace the title with the **headline number**.",
          "2. Shrink the body to `18 pt` and remove the empty placeholder.",
          "   - Keep the theme font",
          "   - Keep the margins",
          "",
          "| Region | Q2 | Q3 |",
          "| --- | ---: | ---: |",
          "| EMEA | 4.1 | 4.9 |",
          "| APAC | 3.2 | 3.6 |",
          "",
          "> Source: *Q3 revenue.xlsx*, sheet 1. See [the style guide](https://example.com/style).",
          "",
          "```js",
          'footnote.slide("257")',
          "```",
        ].join("\n"),
      },
      staleWrite,
      write,
      runCode,
      pendingRender,
    ]),
    toolResult(
      staleWrite.id,
      staleWrite.name,
      [{ type: "text", text: "Slide 3 changed since you last read it; call get_slide again." }],
      true,
    ),
  ];

  return {
    id: "chat-demo",
    title: "Tighten slide 3",
    messages,
    isStreaming: true,
    pendingApprovals: [
      {
        id: "approval-1",
        toolCallId: write.id,
        toolName: write.name,
        access: "write",
        summary: "Update 3 shapes on slide 257",
        args: write.arguments,
      },
      {
        id: "approval-2",
        toolCallId: runCode.id,
        toolName: runCode.name,
        access: "code",
        summary: "Align every shape on slide 3 to the left margin",
        code,
        args: runCode.arguments,
      },
    ],
    permissionMode: "ask",
    modelId: "claude-opus-5-5",
    thinkingLevel: "high",
    queuedMessages: [
      { id: "queued-1", text: "Then give slide 4 the same treatment", attachments: [] },
      {
        id: "queued-2",
        text: "Also check the speaker notes on every slide for typos and tighten anything longer than three sentences",
        attachments: [],
      },
    ],
    stagedAttachments: [
      {
        id: "att-2",
        name: "brand-guidelines.pdf",
        mimeType: "application/pdf",
        kind: "pdf",
        size: 812_000,
        summary: "12 pages",
      },
    ],
    canUndo: true,
    contextUsage: { tokens: 48_200, window: 200_000 },
  };
}

function userMessage(content: UserMessage["content"]): UserMessage {
  return { role: "user", content, timestamp: Date.now() };
}

function assistantMessage(
  content: AssistantMessage["content"],
  extra: Partial<AssistantMessage> = {},
): AssistantMessage {
  return {
    role: "assistant",
    content,
    api: "anthropic-messages",
    provider: "anthropic",
    model: "claude-opus-5-5",
    usage: {
      input: 0,
      output: 0,
      cacheRead: 0,
      cacheWrite: 0,
      totalTokens: 0,
      cost: { input: 0, output: 0, cacheRead: 0, cacheWrite: 0, total: 0 },
    },
    stopReason: "stop",
    timestamp: Date.now(),
    ...extra,
  };
}

function toolResult(
  toolCallId: string,
  toolName: string,
  content: ToolResultMessage["content"],
  isError = false,
): ToolResultMessage {
  return { role: "toolResult", toolCallId, toolName, content, isError, timestamp: Date.now() };
}

/** A 16:9 slide-like SVG standing in for a render. */
function slideImage(title: string, color: string): ImageContent {
  const svg = `<svg xmlns="http://www.w3.org/2000/svg" width="640" height="360"><rect width="640" height="360" fill="#fff"/><rect width="640" height="8" fill="${color}"/><text x="48" y="96" font-family="Segoe UI, sans-serif" font-size="40" font-weight="600" fill="#111">${title}</text><rect x="48" y="140" width="380" height="14" rx="4" fill="#d4d4d4"/><rect x="48" y="170" width="300" height="14" rx="4" fill="#d4d4d4"/><rect x="460" y="140" width="132" height="170" fill="${color}" opacity=".8"/></svg>`;
  return { type: "image", mimeType: "image/svg+xml", data: btoa(svg) };
}

function delay(ms: number) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}
