import type {
  ChatSession,
  FootnoteApp,
  HostModule,
  KeyValueStore,
  OfficeHost,
  SkillDefinition,
} from "../contracts.ts";
import { createProviderClient } from "../providers/providerClient.ts";
import { loadSettings } from "../settings/settings.ts";
import { createChatStore, type ChatRecord } from "../storage/chatStore.ts";
import { openIndexedDbStore } from "../storage/keyValueStores.ts";
import { createChatSession } from "./chatSession.ts";

export interface FootnoteAppOptions {
  host: OfficeHost;
  hostModule: HostModule;
  /** Settings live here; chats too when IndexedDB is unavailable. */
  store: KeyValueStore;
}

export async function createFootnoteApp({
  host,
  hostModule,
  store,
}: FootnoteAppOptions): Promise<FootnoteApp> {
  const settings = await loadSettings(store);
  const provider = createProviderClient(settings.get);
  const chatStore = createChatStore((await openIndexedDbStore("footnote")) ?? store);
  const skills = (): SkillDefinition[] => [...hostModule.skills, ...settings.get().customSkills];
  // One live session per chat, so reopening a chat mid-response returns the running session.
  const openSessions = new Map<string, ChatSession>();

  const startSession = (record: ChatRecord) => {
    const session = createChatSession({
      record,
      host,
      hostModule,
      settings,
      provider,
      chatStore,
      skills,
    });
    openSessions.set(record.id, session);
    return session;
  };

  return {
    host,
    hostModule,
    settings,
    models: { list: () => provider.listModels() },
    chats: {
      list: (documentId) => chatStore.list(documentId),
      async create(documentId) {
        const { defaultPermissionMode, modelId, thinkingLevel } = settings.get();
        return startSession({
          id: crypto.randomUUID(),
          documentId,
          title: "New chat",
          updatedAt: Date.now(),
          messages: [],
          permissionMode: defaultPermissionMode,
          writesAllowed: false,
          ...(modelId !== undefined && { modelId }),
          thinkingLevel,
        });
      },
      async open(chatId) {
        const open = openSessions.get(chatId);
        if (open) return open;
        const record = await chatStore.get(chatId);
        if (!record) throw new Error("This chat no longer exists.");
        return startSession(record);
      },
      async delete(chatId) {
        openSessions.get(chatId)?.abort();
        openSessions.delete(chatId);
        await chatStore.delete(chatId);
      },
    },
    skills: { list: skills },
  };
}
