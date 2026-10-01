import { describe, expect, it } from "vitest";
import { defaultSettings, loadSettings } from "../../src/settings/settings.ts";
import { createChatStore } from "../../src/storage/chatStore.ts";
import { memoryStore, newChatRecord } from "./fakes.ts";

describe("settings", () => {
  it("starts from defaults and persists updates", async () => {
    const store = memoryStore();
    const settings = await loadSettings(store);
    expect(settings.get()).toEqual(defaultSettings);

    await settings.update({ modelId: "gpt-6-sol" });

    expect((await loadSettings(store)).get()).toEqual({ ...defaultSettings, modelId: "gpt-6-sol" });
  });
});

describe("chat store", () => {
  it("indexes chats per document, newest first, and forgets deleted ones", async () => {
    const chats = createChatStore(memoryStore());
    await Promise.all([
      chats.put(newChatRecord({ id: "a", updatedAt: 1 })),
      chats.put(newChatRecord({ id: "b", updatedAt: 2 })),
      chats.put(newChatRecord({ id: "c", documentId: "doc-2", updatedAt: 3 })),
    ]);

    expect((await chats.list("doc-1")).map((chat) => chat.id)).toEqual(["b", "a"]);

    await chats.delete("b");
    expect((await chats.list("doc-1")).map((chat) => chat.id)).toEqual(["a"]);
    expect(await chats.get("b")).toBeUndefined();
  });
});
