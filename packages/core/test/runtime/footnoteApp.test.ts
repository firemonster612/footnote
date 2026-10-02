import { describe, expect, it } from "vitest";
import { createFootnoteEngine } from "../../src/runtime/footnoteApp.ts";
import { createChatStore } from "../../src/storage/chatStore.ts";
import { fakeHost, fakeHostModule, memoryStore, newChatRecord } from "./fakes.ts";

describe("footnote engine", () => {
  it("opens one session for a chat opened twice at once", async () => {
    const store = memoryStore();
    await createChatStore(store).put(newChatRecord());
    const engine = await createFootnoteEngine({ hostModule: fakeHostModule().hostModule, store });
    const app = engine.appFor(fakeHost);

    const [first, second] = await Promise.all([app.chats.open("chat-1"), app.chats.open("chat-1")]);

    expect(first).toBe(second);
  });
});
