import type {
  AgentMessage,
  ChatSummary,
  KeyValueStore,
  PermissionMode,
  ThinkingLevel,
} from "../contracts.ts";

/** Everything needed to reopen a chat. Messages exclude system messages; the session rebuilds those. */
export interface ChatRecord extends ChatSummary {
  messages: AgentMessage[];
  permissionMode: PermissionMode;
  /** The user chose "Allow writes for this chat". */
  writesAllowed: boolean;
  modelId?: string;
  thinkingLevel: ThinkingLevel;
  /** Requests that started a turn (by message timestamp) and that turn's undo ID. Absent in chats saved before revert existed. */
  turnStarts?: TurnStart[];
}

export interface TurnStart {
  messageTimestamp: number;
  turnId: string;
}

export interface ChatStore {
  /** Most recently updated first. */
  list(documentId: string): Promise<ChatSummary[]>;
  get(chatId: string): Promise<ChatRecord | undefined>;
  put(record: ChatRecord): Promise<void>;
  delete(chatId: string): Promise<void>;
}

const chatKey = (chatId: string) => `chat:${chatId}`;
const documentIndexKey = (documentId: string) => `chats:${documentId}`;

/** Chats keyed by ID, plus a per-document index of summaries. */
export function createChatStore(kv: KeyValueStore): ChatStore {
  // Index updates are read-modify-write; serialize them so concurrent saves don't drop entries.
  let pending: Promise<unknown> = Promise.resolve();
  const serialized = <T>(task: () => Promise<T>): Promise<T> => {
    const run = pending.then(task);
    pending = run.catch(() => undefined);
    return run;
  };

  const readIndex = async (documentId: string) =>
    (await kv.get<ChatSummary[]>(documentIndexKey(documentId))) ?? [];

  return {
    list: async (documentId) =>
      (await readIndex(documentId)).toSorted((a, b) => b.updatedAt - a.updatedAt),
    get: (chatId) => kv.get<ChatRecord>(chatKey(chatId)),
    put: (record) =>
      serialized(async () => {
        await kv.set(chatKey(record.id), record);
        const summary: ChatSummary = {
          id: record.id,
          documentId: record.documentId,
          title: record.title,
          updatedAt: record.updatedAt,
        };
        const others = (await readIndex(record.documentId)).filter(
          (entry) => entry.id !== record.id,
        );
        await kv.set(documentIndexKey(record.documentId), [...others, summary]);
      }),
    delete: (chatId) =>
      serialized(async () => {
        const record = await kv.get<ChatRecord>(chatKey(chatId));
        if (!record) return;
        await kv.delete(chatKey(chatId));
        const remaining = (await readIndex(record.documentId)).filter(
          (entry) => entry.id !== chatId,
        );
        await kv.set(documentIndexKey(record.documentId), remaining);
      }),
  };
}
