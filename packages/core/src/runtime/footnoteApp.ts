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
import { createChatSession, type LiveChatSession } from "./chatSession.ts";

export interface FootnoteEngineOptions {
  hostModule: HostModule;
  /** Settings live here; chats too when IndexedDB is unavailable. */
  store: KeyValueStore;
}

export interface FootnoteAppOptions extends FootnoteEngineOptions {
  host: OfficeHost;
}

/** Settings, chats and running sessions shared by every Office document the shell is attached to. */
export interface FootnoteEngine {
  /**
   * An app whose chats talk to `host`. A chat opened through it moves to `host` unless it's running, so a running
   * chat keeps the document it started on.
   */
  appFor(host: OfficeHost): FootnoteApp;
}

interface OpenChat {
  session: LiveChatSession;
  binding: { host: OfficeHost };
}

export async function createFootnoteApp({
  host,
  ...options
}: FootnoteAppOptions): Promise<FootnoteApp> {
  return (await createFootnoteEngine(options)).appFor(host);
}

export async function createFootnoteEngine({
  hostModule,
  store,
}: FootnoteEngineOptions): Promise<FootnoteEngine> {
  const settings = await loadSettings(store);
  const provider = createProviderClient(settings.get);
  const chatStore = createChatStore((await openIndexedDbStore("footnote")) ?? store);
  const skills = (): SkillDefinition[] => [...hostModule.skills, ...settings.get().customSkills];
  // One live session per chat, so reopening a chat mid-response returns the running session.
  const openChats = new Map<string, OpenChat>();

  function startSession(record: ChatRecord, host: OfficeHost): ChatSession {
    const binding = { host };
    const session = createChatSession({
      record,
      host: boundHost(binding),
      hostModule,
      settings,
      provider,
      chatStore,
      skills,
    });
    openChats.set(record.id, { session, binding });
    return session;
  }

  /** The chat's live session, moved to `host` unless it's running. */
  function openSession(chatId: string, host: OfficeHost): ChatSession | undefined {
    const open = openChats.get(chatId);
    if (open && !open.session.getState().isStreaming) open.binding.host = host;
    return open?.session;
  }

  return {
    appFor: (host) => ({
      host,
      hostModule,
      settings,
      models: { list: () => provider.listModels() },
      chats: {
        list: (documentId) => chatStore.list(documentId),
        async create(documentId) {
          const { defaultPermissionMode, modelId, thinkingLevel } = settings.get();
          return startSession(
            {
              id: crypto.randomUUID(),
              documentId,
              title: "New chat",
              updatedAt: Date.now(),
              messages: [],
              permissionMode: defaultPermissionMode,
              writesAllowed: false,
              ...(modelId !== undefined && { modelId }),
              thinkingLevel,
            },
            host,
          );
        },
        async open(chatId) {
          const open = openSession(chatId, host);
          if (open) return open;
          const record = await chatStore.get(chatId);
          if (!record) throw new Error("This chat no longer exists.");
          // Another open of this chat may have finished while this one read storage.
          return openSession(chatId, host) ?? startSession(record, host);
        },
        async delete(chatId) {
          // Disposed first: a running chat would otherwise save itself again after the delete.
          openChats.get(chatId)?.session.dispose();
          openChats.delete(chatId);
          await chatStore.delete(chatId);
        },
      },
      skills: { list: skills },
    }),
  };
}

/** A host that forwards to whichever host the chat is currently bound to. */
function boundHost(binding: { host: OfficeHost }): OfficeHost {
  return {
    status: () => binding.host.status(),
    onStatusChange: (listener) => binding.host.onStatusChange(listener),
    call: (op, args, options) => binding.host.call(op, args, options),
    runCode: (code, options) => binding.host.runCode(code, options),
  };
}
