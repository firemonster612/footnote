import { createFauxCore, type Api, type Model } from "@earendil-works/pi-ai";
import { Type } from "typebox";
import type { FootnoteTool, HostModule, KeyValueStore, OfficeHost } from "../../src/contracts.ts";
import type { ProviderClient } from "../../src/providers/providerClient.ts";
import { defaultSettings, type SettingsHandle } from "../../src/settings/settings.ts";
import { createChatStore, type ChatRecord } from "../../src/storage/chatStore.ts";
import { createChatSession } from "../../src/runtime/chatSession.ts";

/** Clones on the way in and out, like IndexedDB and chrome.storage. */
export function memoryStore(): KeyValueStore {
  const values = new Map<string, unknown>();
  return {
    get: async <T>(key: string) => structuredClone(values.get(key)) as T | undefined,
    set: async (key, value) => void values.set(key, structuredClone(value)),
    delete: async (key) => void values.delete(key),
  };
}

export const fakeHost: OfficeHost = {
  status: async () => ({ connected: true, host: "powerpoint", documentId: "doc-1" }),
  onStatusChange: () => () => {},
  call: async () => {
    throw new Error("No ops in these tests.");
  },
  runCode: async () => ({ ok: true, logs: [] }),
};

function textTool(name: string, access: FootnoteTool["access"], run: () => string): FootnoteTool {
  return {
    name,
    label: name,
    description: `${name} tool`,
    parameters: Type.Object({}),
    access,
    execute: async () => ({ content: [{ type: "text", text: run() }], details: undefined }),
  };
}

/** A host module whose document version bumps on every add_slide call. */
export function fakeHostModule(options: { readResultChars?: number } = {}) {
  let version = 0;
  const contextCalls: number[] = [];
  const turns: string[] = [];
  const hostModule: HostModule = {
    kind: "powerpoint",
    ops: {},
    systemPrompt: "Host rules.",
    createTools: () => [
      textTool("get_deck", "read", () => "x".repeat(options.readResultChars ?? 10)),
      textTool("add_slide", "write", () => `added slide ${++version}`),
      textTool("run_code", "code", () => "ran"),
    ],
    getContextBlock: async () => {
      contextCalls.push(version);
      return `<deck_state>v${version}</deck_state>`;
    },
    undo: {
      beginTurn: (_chatId, turnId) => void turns.push(turnId),
      canUndo: () => turns.length > 0,
      undoLastTurn: async () => ({ restored: 1, removed: 0, warnings: [] }),
    },
    skills: [],
  };
  return { hostModule, contextCalls, turns };
}

export function fakeSettings(): SettingsHandle {
  const settings = {
    ...defaultSettings,
    endpoint: { baseUrl: "http://proxy.test", apiKey: "k" },
    modelId: "faux-model",
  };
  return { get: () => settings, update: async () => {}, subscribe: () => () => {} };
}

export function fauxProvider(model: { contextWindow?: number; maxTokens?: number } = {}) {
  const faux = createFauxCore({
    api: "faux",
    provider: "faux",
    models: [{ id: "faux-model", reasoning: true, input: ["text", "image"], ...model }],
  });
  const client: ProviderClient = {
    listModels: async () => [],
    resolveModel: async () => faux.getModel() as Model<Api>,
    streamFn: faux.streamSimple,
  };
  return { faux, client };
}

export function newChatRecord(overrides: Partial<ChatRecord> = {}): ChatRecord {
  return {
    id: "chat-1",
    documentId: "doc-1",
    title: "New chat",
    updatedAt: 0,
    messages: [],
    permissionMode: "ask",
    writesAllowed: false,
    modelId: "faux-model",
    thinkingLevel: "medium",
    ...overrides,
  };
}

export function sessionHarness(
  options: { readResultChars?: number; contextWindow?: number; maxTokens?: number } = {},
) {
  const host = fakeHostModule(options);
  const provider = fauxProvider(options);
  const chatStore = createChatStore(memoryStore());
  const open = (record: ChatRecord) =>
    createChatSession({
      record,
      host: fakeHost,
      hostModule: host.hostModule,
      settings: fakeSettings(),
      provider: provider.client,
      chatStore,
      skills: () => [],
    });
  return { ...host, ...provider, chatStore, open };
}
