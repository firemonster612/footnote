import { withToolNameRepair } from "./toolNames.ts";
import {
  Agent,
  type AgentContext,
  type AgentLoopTurnUpdate,
  type BeforeToolCallResult,
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

/** A session as the engine holds it. */
export interface LiveChatSession extends ChatSession {
  /** Stops the run and every later save, so deleting the chat afterwards sticks. */
  dispose(): void;
}

interface QueuedRequest {
  id: string;
  text: string;
  attachmentIds: string[];
}

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
}: ChatSessionDeps): LiveChatSession {
  const chatId = record.id;
  const attachments = createAttachmentStore();
  const env: ToolEnv = { chatId, host, attachments, settings: settings.get };
  // Host tools may keep per-chat state (e.g. last-read fingerprints), so they live as long as the session.
  const hostTools = hostModule.createTools(env);
  const sessionTools: FootnoteTool[] = [
    ...hostTools,
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
  let queued: QueuedRequest[] = [];
  // Steers handed to Pi during the current run; ones the model never received go back to the queue when it ends.
  let steers: { message: UserMessage; request: QueuedRequest }[] = [];
  // Labels that resolved a slide position never change; ones that fell back to a raw ID are retried when the
  // transcript changes (a new deck state may resolve them).
  const settledLabels = new Map<string, string>();
  const rawSlideId = /\d+#\d+/;
  let committedLabels: {
    count: number;
    last: AgentMessage | undefined;
    labels: Record<string, string>;
  } = { count: -1, last: undefined, labels: {} };
  // Chats saved before turn tracking: treat every request as a turn start. Their turn IDs are unknown to the
  // host, so reverting them reports that the slide changes can't be undone.
  let turnStarts =
    record.turnStarts ??
    record.messages.flatMap((message) =>
      message.role === "user"
        ? [{ messageTimestamp: message.timestamp, turnId: `untracked-${message.timestamp}` }]
        : [],
    );
  let contextTokens: number | undefined;
  let turnCount = 0;
  // Set between send() starting and the agent run starting, so a second send steers instead of racing.
  let startingRun: AbortController | undefined;
  // Set while a revert or undo restores the document. Requests queue behind it and start once it's done.
  let restoring = false;
  let queuedWhileRestoring = false;
  let disposed = false;
  // A chat is stored from its first request on; one that never had a request isn't worth a record.
  let stored = record.messages.length > 0;
  const compactionBreaker = createCompactionBreaker();

  const permissions = createPermissionGate(
    { mode: record.permissionMode, writesAllowed: record.writesAllowed },
    () => {
      notify();
      void persist();
    },
  );

  const toLlm = (messages: AgentMessage[]) =>
    toLlmMessages(
      messages,
      model?.input.includes("image") ? recentImageMessages : 0,
      turnStarts.at(-1)?.messageTimestamp,
    );

  const agent = new Agent({
    initialState: {
      systemPrompt: buildSystemPrompt(hostModule.systemPrompt, formatSkillListing(skills())),
      tools,
      messages: record.messages,
      thinkingLevel,
    },
    streamFn: withToolNameRepair(provider.streamFn, () => tools.map((tool) => tool.name)),
    convertToLlm: toLlm,
    beforeToolCall: async ({ toolCall, args }, signal) => {
      const tool = tools.find((candidate) => candidate.name === toolCall.name);
      if (!tool) return undefined;
      const decision = await permissions.check(tool, toolCall, args, signal);
      if (decision?.block || !hostTools.includes(tool)) return decision;
      return blockIfDocumentChanged();
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
    // Not awaited: Pi awaits listeners, and saving shouldn't hold up the run. Saves are queued in order. A request
    // is saved as it lands, so a panel reopened mid-run lists the chat.
    const requestLanded = event.type === "message_end" && event.message.role === "user";
    if (requestLanded || event.type === "turn_end" || event.type === "agent_end") void persist();
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
      stagedAttachments: attachmentMetas(stagedIds),
      revertibleRequests: turnStarts.map((start) => start.messageTimestamp),
      toolCallLabels: streamingMessage
        ? { ...transcriptLabels(), ...toolCallLabels([streamingMessage]) }
        : transcriptLabels(),
      queuedMessages: queued.map(({ id, text, attachmentIds }) => ({
        id,
        text,
        attachments: attachmentMetas(attachmentIds),
      })),
      canUndo: hostModule.undo.canUndo(chatId),
      ...(model &&
        contextTokens !== undefined && {
          contextUsage: { tokens: contextTokens, window: model.contextWindow },
        }),
      ...(pendingError !== undefined && { error: pendingError }),
    };
  }

  /** Labels for the transcript's tool calls, rebuilt only when the transcript changed (not per streaming delta). */
  function transcriptLabels(): Record<string, string> {
    const messages = agent.state.messages;
    if (committedLabels.count !== messages.length || committedLabels.last !== messages.at(-1)) {
      committedLabels = {
        count: messages.length,
        last: messages.at(-1),
        labels: toolCallLabels(messages),
      };
    }
    return committedLabels.labels;
  }

  function toolCallLabels(messages: AgentMessage[]): Record<string, string> {
    const labels: Record<string, string> = {};
    for (const message of messages) {
      if (message.role !== "assistant") continue;
      for (const part of message.content) {
        if (part.type !== "toolCall") continue;
        const settled = settledLabels.get(part.id);
        if (settled) {
          labels[part.id] = settled;
          continue;
        }
        const tool = tools.find((candidate) => candidate.name === part.name);
        if (!tool) continue;
        try {
          const label = tool.describeCall?.(part.arguments) ?? tool.label;
          labels[part.id] = label;
          if (!rawSlideId.test(label)) settledLabels.set(part.id, label);
        } catch {
          // Arguments of a call that is still streaming can be partial.
        }
      }
    }
    return labels;
  }

  function attachmentMetas(ids: string[]) {
    return attachments.list().filter((meta) => ids.includes(meta.id));
  }

  function notify() {
    snapshot = buildState();
    for (const listener of listeners) listener(snapshot);
  }

  function updateContextTokens() {
    contextTokens = estimateRequestTokens(toLlm(agent.state.messages));
  }

  async function persist() {
    if (disposed) return;
    const messages = agent.state.messages.filter((message) => message.role !== "system");
    if (!stored && messages.length === 0) return;
    stored = true;
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
        turnStarts,
      });
    } catch (cause) {
      error = `Couldn't save this chat: ${errorText(cause)}`;
      notify();
    }
  }

  /** Blocks document tools once the chat's tab shows another document. A disconnected host may come back to this one. */
  async function blockIfDocumentChanged(): Promise<BeforeToolCallResult | undefined> {
    const { connected, documentId } = await host.status();
    // An unsaved deck (add-in IDs like "unsaved-<uuid>") gets its real ID when saved; that's the same document.
    const savedSinceStart = record.documentId.startsWith("unsaved-");
    if (!connected || !documentId || documentId === record.documentId || savedSinceStart)
      return undefined;
    return {
      block: true,
      reason:
        "The tab this chat runs in now shows a different document, so document tools are blocked. Ask the user to switch back to the original document, or to start a new chat in this one.",
    };
  }

  /** The skill listing is part of the system prompt; skills added or removed since the last run update it. */
  function refreshSystemPrompt() {
    const [leading, ...rest] = agent.state.messages;
    const systemPrompt = buildSystemPrompt(hostModule.systemPrompt, formatSkillListing(skills()));
    if (leading?.role !== "system" || leading.content === systemPrompt) return;
    agent.state.messages = [{ ...leading, content: systemPrompt }, ...rest];
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

  function takeStagedIds(): string[] {
    const ids = stagedIds;
    stagedIds = [];
    return ids;
  }

  // Requests are identified by timestamp (revert, turn tracking), so two sent in the same millisecond must differ.
  let lastRequestTimestamp = Math.max(
    0,
    ...record.messages.map((message) => message.timestamp ?? 0),
  );
  function nextRequestTimestamp(): number {
    lastRequestTimestamp = Math.max(Date.now(), lastRequestTimestamp + 1);
    return lastRequestTimestamp;
  }

  function userMessage(text: string, attachmentIds: string[]): UserMessage {
    const attached = attachmentIds.flatMap((id): (TextContent | ImageContent)[] => {
      const attachment = attachments.get(id);
      return attachment ? attachmentToContent(attachment) : [];
    });
    return {
      role: "user",
      content: [{ type: "text", text }, ...attached],
      timestamp: nextRequestTimestamp(),
    };
  }

  function isRunning(): boolean {
    return agent.state.isStreaming || startingRun !== undefined;
  }

  function isBusy(): boolean {
    return isRunning() || restoring;
  }

  function assertIdle(action: string) {
    if (isRunning()) throw new Error(`Stop the response before ${action}.`);
    if (restoring) throw new Error(`Wait for the current undo to finish before ${action}.`);
  }

  /** A stopped or failed run leaves the queue for the user; only a run that ended normally hands over. */
  function endedNormally(): boolean {
    const lastMessage = agent.state.messages.at(-1);
    return (
      error === undefined &&
      !(
        lastMessage?.role === "assistant" &&
        (lastMessage.stopReason === "aborted" || lastMessage.stopReason === "error")
      )
    );
  }

  function startNextQueued() {
    const [next, ...rest] = queued;
    if (!next || isBusy() || disposed) return;
    queued = rest;
    void startRun(next.text, next.attachmentIds);
  }

  function steerNow(text: string, attachmentIds: string[]) {
    const message = userMessage(text, attachmentIds);
    steers = [...steers, { message, request: { id: crypto.randomUUID(), text, attachmentIds } }];
    agent.steer(message);
    notify();
  }

  /** Pi keeps undelivered steers across a stop; take them back so the next run doesn't send them unseen. */
  function takeUndeliveredSteers(): QueuedRequest[] {
    const delivered = new Set(
      agent.state.messages.flatMap((message) =>
        message.role === "user" ? [message.timestamp] : [],
      ),
    );
    const undelivered = steers.filter(({ message }) => !delivered.has(message.timestamp));
    steers = [];
    agent.clearSteeringQueue();
    return undelivered.map(({ request }) => request);
  }

  /** Runs a revert or undo with sends queued behind it; ones sent meanwhile start when it's done. */
  async function restore<T>(work: () => Promise<T>): Promise<T> {
    restoring = true;
    try {
      return await work();
    } finally {
      restoring = false;
      if (queuedWhileRestoring) {
        queuedWhileRestoring = false;
        startNextQueued();
      }
    }
  }

  async function startRun(text: string, attachmentIds: string[] = takeStagedIds()) {
    const starting = new AbortController();
    startingRun = starting;
    error = undefined;
    notify();
    // A request stopped before its run started goes back to the front of the queue instead of vanishing.
    let unstarted: QueuedRequest[] = [];
    try {
      if (!modelId) throw new Error("Choose a model first.");
      model = await provider.resolveModel(modelId);
      thinkingLevel = clampThinkingLevel(model, thinkingLevel);
      tools = currentTools();
      agent.state.model = model;
      agent.state.thinkingLevel = thinkingLevel;
      agent.state.tools = tools;
      refreshSystemPrompt();
      const compacted = await compactIfNeeded(agent.state.messages, starting.signal);
      if (compacted) agent.state.messages = compacted;
      const documentContext = await freshDocumentContext(agent.state.messages);
      if (starting.signal.aborted) {
        // Stopping isn't an error, even if it cut a compaction short.
        error = undefined;
        unstarted = [{ id: crypto.randomUUID(), text, attachmentIds }];
        return;
      }

      const turnId = crypto.randomUUID();
      hostModule.undo.beginTurn(chatId, turnId);
      const prompt = userMessage(text, attachmentIds);
      turnStarts = [...turnStarts, { messageTimestamp: prompt.timestamp, turnId }];
      if (!agent.state.messages.some((message) => message.role === "user"))
        title = titleFrom(text) ?? title;
      startingRun = undefined;
      await agent.prompt(documentContext ? [prompt, documentContext] : [prompt]);
    } catch (cause) {
      error = errorText(cause);
    } finally {
      startingRun = undefined;
      queued = [...unstarted, ...takeUndeliveredSteers(), ...queued];
      notify();
    }
    if (endedNormally()) startNextQueued();
  }

  const session: LiveChatSession = {
    getState: () => snapshot,
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    send(text) {
      if (isBusy()) {
        session.queue(text);
        return Promise.resolve();
      }
      return startRun(text);
    },
    steer(text) {
      if (isRunning()) steerNow(text, takeStagedIds());
      else session.queue(text);
    },
    queue(text) {
      if (!isBusy()) {
        void startRun(text);
        return;
      }
      if (restoring) queuedWhileRestoring = true;
      queued = [...queued, { id: crypto.randomUUID(), text, attachmentIds: takeStagedIds() }];
      notify();
    },
    removeQueued(id) {
      const removed = queued.find((message) => message.id === id);
      if (!removed) return;
      queued = queued.filter((message) => message !== removed);
      for (const attachmentId of removed.attachmentIds) attachments.remove(attachmentId);
      notify();
    },
    steerQueued(id) {
      const message = queued.find((candidate) => candidate.id === id);
      if (!message || restoring) return;
      queued = queued.filter((candidate) => candidate !== message);
      if (isRunning()) steerNow(message.text, message.attachmentIds);
      else void startRun(message.text, message.attachmentIds);
    },
    abort() {
      // Aborting the run's signal also cancels pending approvals (see PermissionGate.check).
      startingRun?.abort();
      agent.abort();
    },
    dispose() {
      disposed = true;
      session.abort();
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
    async revertTo(messageTimestamp) {
      assertIdle("reverting");
      return restore(() => revertTo(messageTimestamp));
    },
    async undoLastTurn() {
      assertIdle("undoing");
      return restore(undoLastTurn);
    },
  };

  async function revertTo(messageTimestamp: number) {
    const index = agent.state.messages.findIndex(
      (message) => message.role === "user" && message.timestamp === messageTimestamp,
    );
    const request = agent.state.messages[index];
    if (!request || request.role !== "user")
      throw new Error("That message is no longer in this chat.");
    if (!turnStarts.some((start) => start.messageTimestamp === messageTimestamp))
      throw new Error(
        "Only requests that started a turn can be reverted; this one steered a running task.",
      );
    const undoTurnIds = turnStarts
      .filter((start) => start.messageTimestamp >= messageTimestamp)
      .map((start) => start.turnId);
    // Undo first: if the document can't be restored, keep the chat as it is rather than half-reverting.
    const undo = await hostModule.undo.undoTurns(env, chatId, undoTurnIds);
    agent.state.messages = agent.state.messages.slice(0, index);
    turnStarts = turnStarts.filter((start) => start.messageTimestamp < messageTimestamp);
    error = undefined;
    notify();
    await persist();
    const text =
      typeof request.content === "string"
        ? request.content
        : (request.content.find((block) => block.type === "text")?.text ?? "");
    return { text, undo };
  }

  async function undoLastTurn() {
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
  }

  return session;
}
