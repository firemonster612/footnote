// Wire format between a FootnoteApp running in one context (the engine) and views of it in others (a side panel).
// Every message is JSON: chrome.runtime ports serialize with JSON, not structured clone.

import type {
  AgentMessage,
  ChatSession,
  ChatSessionState,
  OfficeHostStatus,
  Settings,
  SkillDefinition,
} from "../contracts.ts";

/** The slice of chrome.runtime.Port a remote app needs, so a chrome port can be passed as is. */
export interface RemotePort<Incoming, Outgoing> {
  postMessage(message: Outgoing): void;
  onMessage: { addListener(listener: (message: Incoming) => void): void };
  onDisconnect: { addListener(listener: () => void): void };
}

export type AppMethod =
  | "settings.update"
  | "models.list"
  | "chats.list"
  | "chats.create"
  | "chats.open"
  | "chats.delete"
  | "host.call"
  | "host.runCode";

export type SessionMethod = Exclude<keyof ChatSession, "getState" | "subscribe">;

export type CallMessage =
  | { id: number; method: AppMethod; args: unknown[] }
  | { id: number; chatId: string; method: SessionMethod; args: unknown[] };

/** A `File` as JSON. */
export interface TransferredFile {
  name: string;
  type: string;
  base64: string;
}

/**
 * Session state without the full transcript: `messages` replaces the receiver's messages from `start` on, so
 * streaming deltas don't resend the whole chat.
 */
export type SessionUpdate = Omit<ChatSessionState, "messages"> & {
  messages: { start: number; items: AgentMessage[] };
};

export type ServerMessage =
  | { type: "hello"; settings: Settings; skills: SkillDefinition[]; status: OfficeHostStatus }
  | { type: "settings"; settings: Settings; skills: SkillDefinition[] }
  | { type: "status"; status: OfficeHostStatus }
  | { type: "session"; chatId: string; update: SessionUpdate }
  | { type: "result"; id: number; ok: true; value: unknown }
  | { type: "result"; id: number; ok: false; error: string };

export type ServerPort = RemotePort<CallMessage, ServerMessage>;
export type ClientPort = RemotePort<ServerMessage, CallMessage>;
