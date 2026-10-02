// Shared contracts between core (agent + UI), host modules (PowerPoint, later Word/Excel), and shells.
// Changing a type here changes every package; keep it small and stable.

import type { AgentMessage, AgentTool, ThinkingLevel } from "@earendil-works/pi-agent-core";
import type { TSchema } from "typebox";

// ---------------------------------------------------------------------------
// Office host: RPC to the realm where Office.js runs.
// Add-in shell: same page. Extension shell: bridge script inside the add-in frame.
// ---------------------------------------------------------------------------

export type OfficeHostKind = "powerpoint" | "word" | "excel";

export interface OfficeHostStatus {
  connected: boolean;
  host?: OfficeHostKind;
  /** Stable per-document key for chat history (Office document URL when available). */
  documentId?: string;
  documentName?: string;
  /** Why we're not connected, phrased for the user (e.g. "Open the Claude add-in in PowerPoint"). */
  reason?: string;
  /** Requirement sets the host supports, e.g. { PowerPointApi: "1.10" }. */
  apiVersions?: Record<string, string>;
}

export interface OfficeCallOptions {
  signal?: AbortSignal;
  timeoutMs?: number;
}

export interface OfficeErrorInfo {
  message: string;
  /** OfficeExtension.Error code when available. */
  code?: string;
  debugInfo?: unknown;
}

export interface CodeRunResult {
  ok: boolean;
  /** JSON-safe return value of the model's code. */
  result?: unknown;
  logs: string[];
  error?: OfficeErrorInfo;
  /** True when the run timed out or the bridge dropped: edits may or may not have applied. */
  outcomeUnknown?: boolean;
  /** Slide IDs in deck order before and after the run, so code that adds slides can be undone. */
  slides?: { before: string[]; after: string[] };
}

export interface OfficeHost {
  status(): Promise<OfficeHostStatus>;
  onStatusChange(listener: (status: OfficeHostStatus) => void): () => void;
  /** Runs a registered op from the host module's OpRegistry inside the Office realm. Args/result are structured-clone safe. */
  call<T = unknown>(op: string, args?: unknown, options?: OfficeCallOptions): Promise<T>;
  /** Code mode: runs a model-written async function body with `context` (PowerPoint.RequestContext) and `footnote` (helpers) in scope. */
  runCode(code: string, options?: OfficeCallOptions): Promise<CodeRunResult>;
}

/** An op runs inside the Office realm. It owns its own `PowerPoint.run` / `Word.run` / `Excel.run`. */
export type OfficeOp = (args: any) => Promise<unknown>;
export type OpRegistry = Record<string, OfficeOp>;

/** Thrown by OfficeHost.call when the op failed inside Office. */
export class OfficeOpError extends Error {
  constructor(
    readonly op: string,
    readonly info: OfficeErrorInfo,
  ) {
    super(`${op} failed: ${describeOfficeError(info)}`);
    this.name = "OfficeOpError";
  }
}

/** "GeneralException" alone is useless; Office's debugInfo says which call failed. */
export function describeOfficeError(info: OfficeErrorInfo): string {
  const debug = (info.debugInfo ?? {}) as { errorLocation?: string; statement?: string };
  const details = [
    info.code && info.code !== info.message ? `code ${info.code}` : undefined,
    debug.errorLocation ? `at ${debug.errorLocation}` : undefined,
    debug.statement ? `statement: ${debug.statement.slice(0, 200)}` : undefined,
  ].filter(Boolean);
  return details.length > 0 ? `${info.message} (${details.join("; ")})` : info.message;
}

// ---------------------------------------------------------------------------
// Tools
// ---------------------------------------------------------------------------

/** Drives permissions. read: never asks. web: never asks. write: asks in Ask mode. code: always asks in Ask mode. */
export type ToolAccess = "read" | "write" | "code" | "web";

export type FootnoteTool<TParameters extends TSchema = TSchema, TDetails = any> = AgentTool<
  TParameters,
  TDetails
> & {
  access: ToolAccess;
  /** One-line, user-facing description of a specific call, shown in approval prompts and tool cards. */
  describeCall?: (args: any) => string;
};

/** Everything a tool may need at execution time. Built per chat by core. */
export interface ToolEnv {
  host: OfficeHost;
  attachments: AttachmentStore;
  settings: () => Settings;
}

// ---------------------------------------------------------------------------
// Host modules: one per Office app. Ops run in the Office realm; the rest runs with the agent.
// ---------------------------------------------------------------------------

export interface UndoReport {
  restored: number;
  removed: number;
  warnings: string[];
}

export interface HostModule {
  kind: OfficeHostKind;
  /** Bundled into the Office realm by each shell. */
  ops: OpRegistry;
  /** Host-specific section of the system prompt. Static, so it stays in the cached prefix. */
  systemPrompt: string;
  createTools(env: ToolEnv): FootnoteTool[];
  /** Per-turn state block (e.g. <deck_state>), appended after cached history. Includes changes since the previous call for this chat. */
  getContextBlock(env: ToolEnv, chatId: string): Promise<string>;
  /** Turn-scoped checkpoints. Core calls beginTurn before each user turn; write tools snapshot what they touch. */
  undo: {
    beginTurn(chatId: string, turnId: string): void;
    canUndo(chatId: string): boolean;
    undoLastTurn(env: ToolEnv, chatId: string): Promise<UndoReport>;
    /** Undoes these turns' document changes, newest first. Warns about turns whose checkpoints are gone (e.g. after a reload). */
    undoTurns(env: ToolEnv, chatId: string, turnIds: string[]): Promise<UndoReport>;
  };
  skills: SkillDefinition[];
}

// ---------------------------------------------------------------------------
// Attachments
// ---------------------------------------------------------------------------

export type AttachmentKind = "image" | "pdf" | "docx" | "spreadsheet" | "pptx" | "text";

export interface AttachmentMeta {
  id: string;
  name: string;
  mimeType: string;
  kind: AttachmentKind;
  size: number;
  /** e.g. "12 pages", "3 sheets", "1920×1080". */
  summary?: string;
}

export interface ProcessedAttachment extends AttachmentMeta {
  /** Raw bytes as base64 (images for insertion, pptx for slide import). */
  base64: string;
  /** Extracted text/markdown (pdf, docx, text, pptx), or CSV-like markdown tables (spreadsheet). */
  text?: string;
  /** Rendered pages or embedded images the model can look at. */
  images?: { mimeType: string; base64: string; label: string }[];
}

export interface AttachmentStore {
  list(): AttachmentMeta[];
  get(id: string): ProcessedAttachment | undefined;
}

// ---------------------------------------------------------------------------
// Skills
// ---------------------------------------------------------------------------

export interface SkillDefinition {
  name: string;
  /** When to use it. Shown in the skill listing; keep under ~200 characters. */
  description: string;
  /** Full instructions, returned by load_skill. Markdown. */
  body: string;
  /** Supporting references loadable by path via load_skill({ name, file }). */
  files?: Record<string, string>;
  source: "bundled" | "custom";
}

// ---------------------------------------------------------------------------
// Settings and storage
// ---------------------------------------------------------------------------

export type PermissionMode = "ask" | "full";
export type ApiFormat = "anthropic-messages" | "openai-responses" | "openai-completions";

export interface Settings {
  endpoint: { baseUrl: string; apiKey: string };
  firecrawlApiKey?: string;
  defaultPermissionMode: PermissionMode;
  modelId?: string;
  thinkingLevel: ThinkingLevel;
  /** Per-model API format override; otherwise inferred from /v1/models owned_by. */
  apiOverrides: Record<string, ApiFormat>;
  customSkills: SkillDefinition[];
}

/** chrome.storage.local in the extension, localStorage/IndexedDB in the add-in. */
export interface KeyValueStore {
  get<T>(key: string): Promise<T | undefined>;
  set<T>(key: string, value: T): Promise<void>;
  delete(key: string): Promise<void>;
}

// ---------------------------------------------------------------------------
// Chat session: implemented by the core runtime, consumed by the UI.
// ---------------------------------------------------------------------------

export interface ApprovalRequest {
  id: string;
  toolCallId: string;
  toolName: string;
  access: ToolAccess;
  /** From FootnoteTool.describeCall, or the tool label. */
  summary: string;
  /** Present for code tools. */
  code?: string;
  args: unknown;
}

export interface ApprovalDecision {
  allow: boolean;
  /** "chat" allows every write (not code) for the rest of this chat. */
  scope?: "once" | "chat";
  comment?: string;
}

export interface ModelInfo {
  id: string;
  ownedBy?: string;
  api: ApiFormat;
  /** Levels this model accepts, including "off" when thinking can be disabled. */
  thinkingLevels: ThinkingLevel[];
  contextWindow: number;
  supportsImages: boolean;
}

export interface ChatSummary {
  id: string;
  documentId: string;
  title: string;
  updatedAt: number;
}

export interface QueuedMessage {
  id: string;
  text: string;
  /** Attachments that were staged when the message was queued; they go with it. */
  attachments: AttachmentMeta[];
}

export interface ChatSessionState {
  id: string;
  title: string;
  messages: AgentMessage[];
  /** Text/thinking/tool-call deltas of the message currently streaming, if any. */
  streamingMessage?: AgentMessage;
  isStreaming: boolean;
  pendingApprovals: ApprovalRequest[];
  permissionMode: PermissionMode;
  modelId?: string;
  thinkingLevel: ThinkingLevel;
  /** Attachments staged for the next message. */
  stagedAttachments: AttachmentMeta[];
  /** Messages waiting for the current run to finish, oldest first. Sent one per run. */
  queuedMessages: QueuedMessage[];
  /** Timestamps of user messages that started a turn; only these can be reverted to (steered messages can't). */
  revertibleRequests: number[];
  /** Tool-card label per tool call ID, from the host's describeCall with the engine's own document state. */
  toolCallLabels: Record<string, string>;
  canUndo: boolean;
  contextUsage?: { tokens: number; window: number };
  error?: string;
}

export interface ChatSession {
  getState(): ChatSessionState;
  subscribe(listener: (state: ChatSessionState) => void): () => void;
  /** Starts a run, or queues the message when one is already running. */
  send(text: string): Promise<void>;
  /** Injects a message while the agent is working (delivered after the current tool batch). */
  steer(text: string): void;
  /** Queues a message (with the staged attachments) to send when the current run finishes. */
  queue(text: string): void;
  removeQueued(id: string): void;
  /** Steers the running task with a queued message now, or sends it right away when nothing is running. */
  steerQueued(id: string): void;
  abort(): void;
  resolveApproval(approvalId: string, decision: ApprovalDecision): void;
  setPermissionMode(mode: PermissionMode): void;
  setModel(modelId: string): void;
  setThinkingLevel(level: ThinkingLevel): void;
  stageAttachment(file: File): Promise<AttachmentMeta>;
  unstageAttachment(id: string): void;
  undoLastTurn(): Promise<UndoReport>;
  /**
   * Reverts the chat to just before the request sent at `messageTimestamp`: undoes the document changes of that
   * turn and every later one, then drops the request and everything after it. Returns the request's text.
   */
  revertTo(messageTimestamp: number): Promise<{ text: string; undo: UndoReport }>;
}

export interface FootnoteApp {
  host: OfficeHost;
  hostModule: HostModule;
  settings: {
    get(): Settings;
    update(patch: Partial<Settings>): Promise<void>;
    subscribe(listener: (settings: Settings) => void): () => void;
  };
  models: {
    /** Fetches /v1/models from the configured endpoint. Throws with a user-readable message on failure. */
    list(): Promise<ModelInfo[]>;
  };
  chats: {
    list(documentId: string): Promise<ChatSummary[]>;
    create(documentId: string): Promise<ChatSession>;
    open(chatId: string): Promise<ChatSession>;
    delete(chatId: string): Promise<void>;
  };
  skills: {
    list(): SkillDefinition[];
  };
}

export type { AgentMessage, ThinkingLevel };
