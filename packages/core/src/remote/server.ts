// Serves a FootnoteApp over a port: pushes settings, host status and session state, and runs the view's calls.

import { base64ToBytes } from "../attachments/base64.ts";
import type { AgentMessage, ChatSession, FootnoteApp } from "../contracts.ts";
import type {
  AppMethod,
  CallMessage,
  ServerMessage,
  ServerPort,
  SessionMethod,
  TransferredFile,
} from "./protocol.ts";

interface WatchedChat {
  chatId: string;
  session: ChatSession;
  /** The messages the view has, by identity; updates send only what changed after them. */
  sent: AgentMessage[];
  lastPushAt: number;
  /** A push waiting out the throttle. */
  pendingPush?: ReturnType<typeof setTimeout>;
  unsubscribe(): void;
}

/** Streaming changes state on every delta; pushing at most this often keeps the port from flooding. The first
 * change after a quiet spell goes out at once, so the view sees a run start without delay. */
const PUSH_INTERVAL_MS = 50;

/** Serves `app` to the view on the other end of `port` until the port disconnects. */
export async function serveFootnoteApp(app: FootnoteApp, port: ServerPort): Promise<void> {
  const watched = new Map<string, WatchedChat>();
  let connected = true;

  function post(message: ServerMessage): void {
    if (!connected) return;
    try {
      port.postMessage(message);
    } catch {
      // The view closed between our check and the post; its disconnect handler cleans up.
    }
  }

  const subscriptions: (() => void)[] = [];
  port.onDisconnect.addListener(() => {
    connected = false;
    for (const unsubscribe of subscriptions) unsubscribe();
    for (const chatId of watched.keys()) unwatch(chatId);
  });

  function push(chat: WatchedChat): void {
    clearTimeout(chat.pendingPush);
    chat.pendingPush = undefined;
    chat.lastPushAt = Date.now();
    const { messages, ...state } = chat.session.getState();
    const changed = messages.findIndex((message, index) => message !== chat.sent[index]);
    const start = changed === -1 ? messages.length : changed;
    chat.sent = messages;
    post({
      type: "session",
      chatId: chat.chatId,
      update: { ...state, messages: { start, items: messages.slice(start) } },
    });
  }

  /** Sends throttled state now, so a call's result never reaches the view before the state it produced. */
  function flushPushes(): void {
    for (const chat of watched.values()) if (chat.pendingPush !== undefined) push(chat);
  }

  /** Starts pushing a session's state to the view, beginning with all of it. */
  function watch(session: ChatSession): string {
    const chatId = session.getState().id;
    if (watched.has(chatId)) return chatId;
    const chat: WatchedChat = { chatId, session, sent: [], lastPushAt: 0, unsubscribe: () => {} };
    chat.unsubscribe = session.subscribe(() => {
      if (chat.pendingPush !== undefined) return;
      const wait = chat.lastPushAt + PUSH_INTERVAL_MS - Date.now();
      if (wait <= 0) push(chat);
      else chat.pendingPush = setTimeout(() => push(chat), wait);
    });
    watched.set(chatId, chat);
    push(chat);
    return chatId;
  }

  function unwatch(chatId: string): void {
    const chat = watched.get(chatId);
    if (!chat) return;
    chat.unsubscribe();
    clearTimeout(chat.pendingPush);
    watched.delete(chatId);
  }

  const appCalls: Record<AppMethod, (...args: any[]) => unknown> = {
    "settings.update": (patch) => app.settings.update(patch),
    "models.list": () => app.models.list(),
    "chats.list": (documentId: string) => app.chats.list(documentId),
    "chats.create": async (documentId: string) => watch(await app.chats.create(documentId)),
    "chats.open": async (chatId: string) => watch(await app.chats.open(chatId)),
    "chats.delete": (chatId: string) => {
      unwatch(chatId);
      return app.chats.delete(chatId);
    },
  };

  function callSession(chatId: string, method: SessionMethod, args: unknown[]): unknown {
    const session = watched.get(chatId)?.session;
    if (!session) throw new Error("This chat isn't open here. Reopen it from the chat list.");
    if (method === "stageAttachment") {
      // The client sends files as TransferredFile (see connectFootnoteApp).
      return session.stageAttachment(receiveFile(args[0] as TransferredFile));
    }
    return Reflect.apply(session[method], session, args);
  }

  async function run(call: CallMessage): Promise<unknown> {
    return "chatId" in call
      ? callSession(call.chatId, call.method, call.args)
      : Reflect.apply(appCalls[call.method], undefined, call.args);
  }

  port.onMessage.addListener((call) => {
    run(call).then(
      (value) => {
        flushPushes();
        post({ type: "result", id: call.id, ok: true, value });
      },
      (error: unknown) => {
        flushPushes();
        post({
          type: "result",
          id: call.id,
          ok: false,
          error: error instanceof Error ? error.message : String(error),
        });
      },
    );
  });

  // Changes after this point are pushed, so the view never applies an older snapshot over a newer one.
  const status = await app.host.status();
  if (!connected) return;
  post({ type: "hello", settings: app.settings.get(), skills: app.skills.list(), status });
  subscriptions.push(
    app.settings.subscribe((settings) =>
      post({ type: "settings", settings, skills: app.skills.list() }),
    ),
    app.host.onStatusChange((status) => post({ type: "status", status })),
  );
}

function receiveFile({ name, type, base64 }: TransferredFile): File {
  return new File([base64ToBytes(base64)], name, { type });
}
