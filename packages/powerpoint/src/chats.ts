import type { ToolEnv } from "@footnote/core/contracts";
import type { SlideOutline } from "./ops/types.ts";

const MAX_UNDO_TURNS = 20;

/** Checkpoints of one user turn: slides as they were before the turn first wrote them, and slides it created. */
export interface Turn {
  id: string;
  snapshots: Map<string, { index: number; base64: string }>;
  created: Set<string>;
  /** Fingerprints after the turn's writes, to warn when undo discards later edits. */
  after: Map<string, string>;
}

export interface ChatState {
  /** Slide ID → fingerprint the model last saw (get_slide or its own write receipts). */
  observed: Map<string, string>;
  /** Outline sent in the previous <deck_state>, for change detection. */
  lastOutline?: SlideOutline[];
  lastThemeKey?: string;
  turns: Turn[];
  /** Every turn begun in this panel session, so a revert can tell "made no changes" from "checkpoints lost on reload". */
  seenTurnIds: Set<string>;
}

export const turnHasChanges = (turn: Turn): boolean =>
  turn.snapshots.size > 0 || turn.created.size > 0;

const newTurn = (id: string): Turn => ({
  id,
  snapshots: new Map(),
  created: new Set(),
  after: new Map(),
});

export interface ChatRegistry {
  get(chatId: string): ChatState;
  bind(env: ToolEnv, chatId: string): void;
  forEnv(env: ToolEnv): ChatState;
  beginTurn(chatId: string, turnId: string): void;
}

// ToolEnv carries no chat ID, so tools find their chat through the env object getContextBlock was called with,
// falling back to the chat whose turn began most recently.
export function createChatRegistry(): ChatRegistry {
  const chats = new Map<string, ChatState>();
  const envChats = new WeakMap<ToolEnv, string>();
  let activeChatId = "default";

  function get(chatId: string): ChatState {
    let chat = chats.get(chatId);
    if (!chat) {
      chat = { observed: new Map(), turns: [], seenTurnIds: new Set() };
      chats.set(chatId, chat);
    }
    return chat;
  }

  return {
    get,
    bind(env: ToolEnv, chatId: string): void {
      envChats.set(env, chatId);
    },
    forEnv(env: ToolEnv): ChatState {
      return get(envChats.get(env) ?? activeChatId);
    },
    beginTurn(chatId: string, turnId: string): void {
      activeChatId = chatId;
      const chat = get(chatId);
      chat.seenTurnIds.add(turnId);
      chat.turns = [
        ...chat.turns.filter(turnHasChanges).slice(-(MAX_UNDO_TURNS - 1)),
        newTurn(turnId),
      ];
    },
  };
}

/** The turn writes belong to; starts an implicit one when core never called beginTurn. */
export function currentTurn(chat: ChatState): Turn {
  const turn = chat.turns.at(-1) ?? newTurn("implicit");
  if (chat.turns.length === 0) chat.turns.push(turn);
  return turn;
}
