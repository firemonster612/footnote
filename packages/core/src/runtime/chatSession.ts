import {
  Agent,
  type AgentContext,
  type AgentLoopTurnUpdate,
  type PrepareNextTurnContext,
} from "@earendil-works/pi-agent-core";
import {
  clampThinkingLevel,
  normalizeContext,
  type Api,
  type ImageContent,
  type Model,
  type TextContent,
  type UserMessage,
} from "@earendil-works/pi-ai";
import {
  attachmentToContent,
  createAttachmentStore,
  createAttachmentTool,
  processAttachment,
} from "../attachments/index.ts";
import {
  compactionThresholdTokens,
  compactMessages,
  createCompactionBreaker,
  estimateRequestTokens,
  keepRecentTokensFor,
  summaryInstructions,
} from "../compaction/compaction.ts";
import {
  documentContextMessage,
  recentImageMessages,
  toLlmMessages,
  type DocumentContextMessage,
} from "../context/messages.ts";
import { buildSystemPrompt } from "../context/systemPrompt.ts";
import { capToolResultText } from "../context/toolResultBudget.ts";
import type {
  AgentMessage,
  ChatSession,
  ChatSessionState,
  FootnoteTool,
  HostModule,
  OfficeHost,
  SkillDefinition,
  ToolEnv,
} from "../contracts.ts";
import { createPermissionGate } from "../permissions/permissionGate.ts";
import type { ProviderClient } from "../providers/providerClient.ts";
import type { SettingsHandle } from "../settings/settings.ts";
import type { ChatRecord, ChatStore } from "../storage/chatStore.ts";
import { createSkillTool, formatSkillListing } from "../skills/index.ts";
import { createWebTools } from "../web/index.ts";

export interface ChatSessionDeps {
  record: ChatRecord;
  host: OfficeHost;
  hostModule: HostModule;
  settings: SettingsHandle;
  provider: ProviderClient;
  chatStore: ChatStore;
  skills: () => SkillDefinition[];
}

const summaryMaxTokens = 8_000;
const titleMaxChars = 60;

function errorText(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function titleFrom(text: string): string | undefined {
  const firstLine = text.trim().split("\n")[0]?.trim();
  if (!firstLine) return undefined;
  return firstLine.length <= titleMaxChars
    ? firstLine
    : `${firstLine.slice(0, titleMaxChars - 1)}…`;
}

export function createChatSession({
  record,
  host,
  hostModule,
  settings,
  provider,
  chatStore,
  skills,
}: ChatSessionDeps): ChatSession {
  const chatId = record.id;
  const attachments = createAttachmentStore();
  const env: ToolEnv = { host, attachments, settings: settings.get };
  // Host tools may keep per-chat state (e.g. last-read fingerprints), so they live as long as the session.
  const sessionTools: FootnoteTool[] = [
    ...hostModule.createTools(env),
    createSkillTool(skills),
    createAttachmentTool(env),
  ];
  const currentTools = () => [...sessionTools, ...createWebTools(env)];
  let tools = currentTools();

  let title = record.title;
  let modelId = record.modelId;
  let thinkingLevel = record.thinkingLevel;
  let model: Model<Api> | undefined;
  let error: string | undefined;
  let stagedIds: string[] = [];
  let contextTokens: number | undefined;
  let turnCount = 0;
  // Set between send() starting and the agent run starting, so a second send steers instead of racing.
  let startingRun: AbortController | undefined;
  const compactionBreaker = createCompactionBreaker();

  const permissions = createPermissionGate(
    { mode: record.permissionMode, writesAllowed: record.writesAllowed },
    () => {
      notify();
      void persist();
    },
  );

  const toLlm = (messages: AgentMessage[]) =>
    toLlmMessages(messages, model?.input.includes("image") ? recentImageMessages : 0);

  const agent = new Agent({
    initialState: {
      systemPrompt: buildSystemPrompt(hostModule.systemPrompt, formatSkillListing(skills())),
      tools,
      messages: record.messages,
      thinkingLevel,
    },
    streamFn: provider.streamFn,
    convertToLlm: toLlm,
    beforeToolCall: async ({ toolCall, args }, signal) => {
      const tool = tools.find((candidate) => candidate.name === toolCall.name);
      return tool ? permissions.check(tool, toolCall, args, signal) : undefined;
    },
    afterToolCall: async ({ result }) => {
      const content = capToolResultText(result.content);
      return content && { content };
    },
    prepareNextTurnWithContext: prepareNextTurn,
    // Office edits depend on each other's order (add a slide, then fill it), so tool calls run one at a time.
    toolExecution: "sequential",
    sessionId: chatId,
  });

  agent.subscribe((event) => {
    if (event.type === "turn_end") turnCount += 1;
    if (event.type !== "message_update") updateContextTokens();
    notify();
    // Not awaited: Pi awaits listeners, and saving shouldn't hold up the run. Saves are queued in order.
    if (event.type === "turn_end" || event.type === "agent_end") void persist();
  });

  const listeners = new Set<(state: ChatSessionState) => void>();
  let snapshot = buildState();

  function buildState(): ChatSessionState {
    const streamingMessage = agent.state.streamingMessage;
    const lastMessage = agent.state.messages.at(-1);
    // A run the user stopped isn't an error worth showing.
    const stoppedByUser = lastMessage?.role === "assistant" && lastMessage.stopReason === "aborted";
    const pendingError = error ?? (stoppedByUser ? undefined : agent.state.errorMessage);
    return {
      id: chatId,
      title,
      messages: agent.state.messages.filter((message) => message.role !== "system"),
      ...(streamingMessage && { streamingMessage }),
      isStreaming: agent.state.isStreaming || startingRun !== undefined,
      pendingApprovals: permissions.pendingApprovals(),
      permissionMode: permissions.state().mode,
      ...(modelId !== undefined && { modelId }),
      thinkingLevel,
      stagedAttachments: attachments.list().filter((meta) => stagedIds.includes(meta.id)),
      canUndo: hostModule.undo.canUndo(chatId),
      ...(model &&
        contextTokens !== undefined && {
          contextUsage: { tokens: contextTokens, window: model.contextWindow },
        }),
      ...(pendingError !== undefined && { error: pendingError }),
    };
  }

  function notify() {
    snapshot = buildState();
    for (const listener of listeners) listener(snapshot);
  }

  function updateContextTokens() {
    contextTokens = estimateRequestTokens(toLlm(agent.state.messages));
  }

  async function persist() {
    const messages = agent.state.messages.filter((message) => message.role !== "system");
    if (messages.length === 0) return;
    const { mode, writesAllowed } = permissions.state();
    try {
      await chatStore.put({
        id: chatId,
        documentId: record.documentId,
        title,
        updatedAt: Date.now(),
        messages,
        permissionMode: mode,
        writesAllowed,
        ...(modelId !== undefined && { modelId }),
        thinkingLevel,
      });
    } catch (cause) {
      error = `Couldn't save this chat: ${errorText(cause)}`;
      notify();
    }
  }

  /** The host's current state block, unless it repeats the newest one already in the transcript. */
  async function freshDocumentContext(
    messages: AgentMessage[],
  ): Promise<DocumentContextMessage | undefined> {
    let text: string;
    try {
      text = await hostModule.getContextBlock(env, chatId);
    } catch (cause) {
      text = `<document_state_unavailable>${errorText(cause)}</document_state_unavailable>`;
    }
    if (!text.trim()) return undefined;
    const newest = messages.findLast((message) => message.role === "documentContext");
    return newest?.text === text ? undefined : documentContextMessage(text);
  }

  async function summarize(transcript: string, signal?: AbortSignal): Promise<string> {
    if (!model) throw new Error("No model selected.");
    const context = normalizeContext({
      systemPrompt: summaryInstructions,
      messages: [{ role: "user", content: transcript, timestamp: Date.now() }],
    });
    const response = await provider.streamFn(model, context, {
      maxTokens: summaryMaxTokens,
      ...(signal && { signal }),
    });
    const message = await response.result();
    if (message.stopReason === "error" || message.stopReason === "aborted") {
      throw new Error(message.errorMessage ?? `summary request ${message.stopReason}`);
    }
    return message.content
      .flatMap((block) => (block.type === "text" ? [block.text] : []))
      .join("\n");
  }

  /** Compacted messages when the transcript is over the model's budget, else undefined. */
  async function compactIfNeeded(
    messages: AgentMessage[],
    signal?: AbortSignal,
  ): Promise<AgentMessage[] | undefined> {
    if (!model) return undefined;
    const tokens = estimateRequestTokens(toLlm(messages));
    if (tokens <= compactionThresholdTokens(model.contextWindow, model.maxTokens)) return undefined;
    if (compactionBreaker.tripped(turnCount)) {
      error =
        "This chat is still too long after compacting it several times. Start a new chat to continue.";
      return undefined;
    }
    compactionBreaker.record(turnCount);
    try {
      return await compactMessages(
        messages,
        keepRecentTokensFor(model.contextWindow),
        summarize,
        signal,
      );
    } catch (cause) {
      error = `Couldn't compact the conversation: ${errorText(cause)}`;
      return undefined;
    }
  }

  async function prepareNextTurn(
    turn: PrepareNextTurnContext,
    signal?: AbortSignal,
  ): Promise<AgentLoopTurnUpdate | undefined> {
    if (signal?.aborted) return undefined;
    const compacted = await compactIfNeeded(turn.context.messages, signal);
    if (compacted) agent.state.messages = compacted;
    const messages = compacted ?? turn.context.messages;
    // Reads can't change the document, so only refresh its state after writes (or after compaction dropped it).
    const changedDocument = turn.toolResults.some((result) => {
      const access = tools.find((tool) => tool.name === result.toolName)?.access;
      return access === "write" || access === "code";
    });
    const documentContext =
      changedDocument || compacted ? await freshDocumentContext(messages) : undefined;
    if (!compacted && !documentContext) return undefined;
    const context: AgentContext | undefined = compacted && { ...turn.context, messages: compacted };
    return { ...(context && { context }), ...(documentContext && { messages: [documentContext] }) };
  }

  function takeUserMessage(text: string): UserMessage {
    const attached = stagedIds.flatMap((id): (TextContent | ImageContent)[] => {
      const attachment = attachments.get(id);
      return attachment ? attachmentToContent(attachment) : [];
    });
    stagedIds = [];
    return { role: "user", content: [{ type: "text", text }, ...attached], timestamp: Date.now() };
  }

  async function startRun(text: string) {
    const starting = new AbortController();
    startingRun = starting;
    error = undefined;
    notify();
    try {
      if (!modelId) throw new Error("Choose a model first.");
      model = await provider.resolveModel(modelId);
      thinkingLevel = clampThinkingLevel(model, thinkingLevel);
      tools = currentTools();
      agent.state.model = model;
      agent.state.thinkingLevel = thinkingLevel;
      agent.state.tools = tools;

      hostModule.undo.beginTurn(chatId, crypto.randomUUID());
      const userMessage = takeUserMessage(text);
      if (!agent.state.messages.some((message) => message.role === "user"))
        title = titleFrom(text) ?? title;
      const compacted = await compactIfNeeded(agent.state.messages, starting.signal);
      if (compacted) agent.state.messages = compacted;
      const documentContext = await freshDocumentContext(agent.state.messages);
      if (starting.signal.aborted) return;
      startingRun = undefined;
      await agent.prompt(documentContext ? [userMessage, documentContext] : [userMessage]);
    } catch (cause) {
      error = errorText(cause);
    } finally {
      startingRun = undefined;
      notify();
    }
  }

  const session: ChatSession = {
    getState: () => snapshot,
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    send(text) {
      if (agent.state.isStreaming || startingRun) {
        session.steer(text);
        return Promise.resolve();
      }
      return startRun(text);
    },
    steer(text) {
      if (!agent.state.isStreaming && !startingRun) {
        void session.send(text);
        return;
      }
      agent.steer(takeUserMessage(text));
      notify();
    },
    abort() {
      // Aborting the run's signal also cancels pending approvals (see PermissionGate.check).
      startingRun?.abort();
      agent.abort();
    },
    resolveApproval: (approvalId, decision) => permissions.resolveApproval(approvalId, decision),
    setPermissionMode: (mode) => permissions.setMode(mode),
    setModel(id) {
      modelId = id;
      notify();
      void persist();
    },
    setThinkingLevel(level) {
      thinkingLevel = model ? clampThinkingLevel(model, level) : level;
      agent.state.thinkingLevel = thinkingLevel;
      notify();
      void persist();
    },
    async stageAttachment(file) {
      const attachment = await processAttachment(file);
      attachments.add(attachment);
      stagedIds = [...stagedIds, attachment.id];
      notify();
      const { base64: _base64, text: _text, images: _images, ...meta } = attachment;
      return meta;
    },
    unstageAttachment(id) {
      stagedIds = stagedIds.filter((stagedId) => stagedId !== id);
      attachments.remove(id);
      notify();
    },
    async undoLastTurn() {
      if (agent.state.isStreaming) throw new Error("Stop the response before undoing.");
      const report = await hostModule.undo.undoLastTurn(env, chatId);
      agent.state.messages = [
        ...agent.state.messages,
        documentContextMessage(
          `<undo>The user undid the document changes from the previous turn (${report.restored} restored, ${report.removed} removed). Don't redo them unless asked.</undo>`,
        ),
      ];
      notify();
      await persist();
      return report;
    },
  };
  return session;
}
