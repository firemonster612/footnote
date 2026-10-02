// A FootnoteApp that mirrors one served elsewhere (see serveFootnoteApp): state arrives as pushes, methods run as calls.

import { bytesToBase64 } from "../attachments/base64.ts";
import type {
  ChatSession,
  ChatSessionState,
  FootnoteApp,
  HostModule,
  OfficeHostStatus,
  Settings,
  SkillDefinition,
} from "../contracts.ts";
import type {
  AppMethod,
  CallMessage,
  ClientPort,
  ServerMessage,
  SessionMethod,
  SessionUpdate,
  TransferredFile,
} from "./protocol.ts";

interface RemoteChat {
  state: ChatSessionState;
  listeners: Set<(state: ChatSessionState) => void>;
  session?: ChatSession;
}

interface PendingCall {
  resolve(value: unknown): void;
  reject(error: Error): void;
}

type Hello = Extract<ServerMessage, { type: "hello" }>;

const disconnectedMessage = "Lost the connection to Footnote. Reopen the panel to reconnect.";
const viewHostCallMessage = "Office calls run in Footnote's engine, not in a view of it.";

/**
 * Resolves once the server has sent its first snapshot. `hostModule` is the view's own copy: the UI reads tool
 * labels from it, and a module can't cross the port.
 */
export async function connectFootnoteApp(
  port: ClientPort,
  hostModule: HostModule,
): Promise<FootnoteApp> {
  const pending = new Map<number, PendingCall>();
  const chats = new Map<string, RemoteChat>();
  const settingsListeners = new Set<(settings: Settings) => void>();
  const statusListeners = new Set<(status: OfficeHostStatus) => void>();
  let nextCallId = 0;
  let settings: Settings;
  let skills: SkillDefinition[];
  let status: OfficeHostStatus;
  let disconnected = false;
  let receiveHello: { resolve(hello: Hello): void; reject(error: Error): void } | undefined;
  const hello = new Promise<Hello>((resolve, reject) => (receiveHello = { resolve, reject }));

  port.onMessage.addListener((message) => {
    switch (message.type) {
      case "hello":
        receiveHello?.resolve(message);
        break;
      case "settings":
        ({ settings, skills } = message);
        for (const listener of settingsListeners) listener(settings);
        break;
      case "status":
        status = message.status;
        for (const listener of statusListeners) listener(status);
        break;
      case "session":
        applySessionUpdate(message.chatId, message.update);
        break;
      case "result": {
        const call = pending.get(message.id);
        pending.delete(message.id);
        if (message.ok) call?.resolve(message.value);
        else call?.reject(new Error(message.error));
        break;
      }
      default:
        message satisfies never;
    }
  });

  port.onDisconnect.addListener(() => {
    disconnected = true;
    const error = new Error(disconnectedMessage);
    receiveHello?.reject(error);
    for (const call of pending.values()) call.reject(error);
    pending.clear();
  });

  function applySessionUpdate(chatId: string, { messages, ...update }: SessionUpdate): void {
    const chat = chats.get(chatId);
    const previous = chat?.state.messages ?? [];
    const state = {
      ...update,
      messages: [...previous.slice(0, messages.start), ...messages.items],
    };
    if (!chat) {
      chats.set(chatId, { state, listeners: new Set() });
      return;
    }
    chat.state = state;
    for (const listener of chat.listeners) listener(state);
  }

  /** Resolves with what the server's handler returned; the caller names its type. */
  function request<T>(message: (id: number) => CallMessage): Promise<T> {
    if (disconnected) return Promise.reject(new Error(disconnectedMessage));
    const id = nextCallId++;
    const result = new Promise<T>((resolve, reject) =>
      // The server answers each call id with its handler's result, which has the type the caller asked for.
      pending.set(id, { resolve: resolve as (value: unknown) => void, reject }),
    );
    port.postMessage(message(id));
    return result;
  }

  const callApp = <T>(method: AppMethod, ...args: unknown[]) =>
    request<T>((id) => ({ id, method, args }));

  async function sessionFor(chatIdPromise: Promise<string>): Promise<ChatSession> {
    const chatId = await chatIdPromise;
    // The server pushes the session's full state before answering the call that opened it.
    const chat = chats.get(chatId)!;
    chat.session ??= remoteSession(chatId, chat);
    return chat.session;
  }

  function remoteSession(chatId: string, chat: RemoteChat): ChatSession {
    const run = <T = void>(method: SessionMethod, ...args: unknown[]) =>
      request<T>((id) => ({ id, chatId, method, args }));
    const fire = (method: SessionMethod, ...args: unknown[]) => {
      run(method, ...args).catch((error: unknown) =>
        console.error(`Footnote: ${method} failed`, error),
      );
    };
    return {
      getState: () => chat.state,
      subscribe(listener) {
        chat.listeners.add(listener);
        return () => chat.listeners.delete(listener);
      },
      send: (text) => run("send", text),
      steer: (text) => fire("steer", text),
      queue: (text) => fire("queue", text),
      removeQueued: (id) => fire("removeQueued", id),
      steerQueued: (id) => fire("steerQueued", id),
      abort: () => fire("abort"),
      resolveApproval: (approvalId, decision) => fire("resolveApproval", approvalId, decision),
      setPermissionMode: (mode) => fire("setPermissionMode", mode),
      setModel: (modelId) => fire("setModel", modelId),
      setThinkingLevel: (level) => fire("setThinkingLevel", level),
      stageAttachment: async (file) => run("stageAttachment", await transferFile(file)),
      unstageAttachment: (id) => fire("unstageAttachment", id),
      undoLastTurn: () => run("undoLastTurn"),
      revertTo: (messageTimestamp) => run("revertTo", messageTimestamp),
    };
  }

  ({ settings, skills, status } = await hello);

  return {
    host: {
      status: async () => status,
      onStatusChange(listener) {
        statusListeners.add(listener);
        return () => statusListeners.delete(listener);
      },
      // Tools run in the engine, behind its approvals; a view only shows their results.
      call: () => Promise.reject(new Error(viewHostCallMessage)),
      runCode: () => Promise.reject(new Error(viewHostCallMessage)),
    },
    hostModule,
    settings: {
      get: () => settings,
      update: (patch) => callApp("settings.update", patch),
      subscribe(listener) {
        settingsListeners.add(listener);
        return () => settingsListeners.delete(listener);
      },
    },
    models: { list: () => callApp("models.list") },
    chats: {
      list: (documentId) => callApp("chats.list", documentId),
      create: (documentId) => sessionFor(callApp("chats.create", documentId)),
      open: (chatId) => sessionFor(callApp("chats.open", chatId)),
      async delete(chatId) {
        await callApp("chats.delete", chatId);
        chats.delete(chatId);
      },
    },
    skills: { list: () => skills },
  };
}

async function transferFile(file: File): Promise<TransferredFile> {
  const base64 = bytesToBase64(new Uint8Array(await file.arrayBuffer()));
  return { name: file.name, type: file.type, base64 };
}
