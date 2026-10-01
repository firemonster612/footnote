// The engine: an offscreen document that runs the agent, so a task keeps going while the side panel is closed.
// Each side panel is a view of one tab (see sidepanel/main.tsx); each tab gets an OfficeHost bound to it.

import {
  createFootnoteEngine,
  serveFootnoteApp,
  type ChatSession,
  type FootnoteApp,
  type FootnoteEngine,
  type KeyValueStore,
} from "@footnote/core";
import { powerpointModule } from "@footnote/powerpoint";
import { createExtensionHost } from "../../lib/extension-host.ts";
import { VIEW_PORT_PREFIX } from "../../lib/protocol.ts";
import { askWorker } from "../../lib/worker.ts";

/** How long the engine stays open with no panel and no running chat. Undo checkpoints live here, so not instantly. */
const IDLE_CLOSE_MS = 10 * 60_000;

interface TabActivity {
  views: number;
  sessions: Set<ChatSession>;
  /** What we last set; undefined once Chrome may have cleared it. */
  badge: string | undefined;
  app?: FootnoteApp;
}

/** Offscreen documents can't reach chrome.storage, so settings go through the service worker. */
const workerStore: KeyValueStore = {
  // Values under our keys are only written by set<T> below.
  get: async <T>(key: string) => (await askWorker("storageGet", key)) as T | undefined,
  set: (key, value) => askWorker("storageSet", key, value),
  delete: (key) => askWorker("storageDelete", key),
};

const engine = createFootnoteEngine({ hostModule: powerpointModule, store: workerStore });
const tabs = new Map<number, TabActivity>();
let idleTimer: ReturnType<typeof setTimeout> | undefined;

function tabActivity(tabId: number): TabActivity {
  const existing = tabs.get(tabId);
  if (existing) return existing;
  const created: TabActivity = { views: 0, sessions: new Set(), badge: "" };
  tabs.set(tabId, created);
  return created;
}

/** The tab's app, with every chat it opens tracked for the toolbar badge and the idle timer. */
function appForTab(ready: FootnoteEngine, tabId: number): FootnoteApp {
  const tab = tabActivity(tabId);
  if (tab.app) return tab.app;
  const host = createExtensionHost(powerpointModule.kind, tabId);
  // Chrome clears a tab's badge when the tab navigates, which reaches us as a host status change.
  host.onStatusChange(() => {
    tab.badge = undefined;
    updateActivity(tabId);
  });
  const app = ready.appFor(host);
  const track = (session: ChatSession) => {
    if (!tab.sessions.has(session)) {
      tab.sessions.add(session);
      session.subscribe(() => updateActivity(tabId));
    }
    return session;
  };
  tab.app = {
    ...app,
    chats: {
      ...app.chats,
      create: async (documentId) => track(await app.chats.create(documentId)),
      open: async (chatId) => track(await app.chats.open(chatId)),
    },
  };
  return tab.app;
}

/** Badges the toolbar icon while the tab's panel is closed, and closes the engine once nothing needs it. */
function updateActivity(tabId: number): void {
  const tab = tabActivity(tabId);
  const states = [...tab.sessions].map((session) => session.getState());
  const badge =
    tab.views > 0
      ? ""
      : states.some((state) => state.pendingApprovals.length > 0)
        ? "!"
        : states.some((state) => state.isStreaming)
          ? "…"
          : "";
  if (badge !== tab.badge) {
    tab.badge = badge;
    askWorker("setBadge", tabId, badge).catch(logError("couldn't badge the toolbar icon"));
  }
  scheduleIdleClose();
}

function scheduleIdleClose(): void {
  const busy = [...tabs.values()].some(
    (candidate) =>
      candidate.views > 0 ||
      [...candidate.sessions].some((session) => session.getState().isStreaming),
  );
  clearTimeout(idleTimer);
  idleTimer = busy
    ? undefined
    : setTimeout(
        () => askWorker("stopEngine").catch(logError("couldn't close the idle engine")),
        IDLE_CLOSE_MS,
      );
}

scheduleIdleClose();

// Registered before the engine finishes loading, so a panel that connects right after the document opens is heard.
chrome.runtime.onConnect.addListener((port) => {
  if (!port.name.startsWith(VIEW_PORT_PREFIX)) return;
  const tabId = Number(port.name.slice(VIEW_PORT_PREFIX.length));
  const tab = tabActivity(tabId);
  let connected = true;
  tab.views += 1;
  updateActivity(tabId);
  port.onDisconnect.addListener(() => {
    connected = false;
    tab.views -= 1;
    updateActivity(tabId);
  });
  engine
    .then((ready) => {
      if (connected) return serveFootnoteApp(appForTab(ready, tabId), port);
    })
    .catch((error: unknown) => {
      logError("couldn't serve the side panel")(error);
      port.disconnect();
    });
});

function logError(what: string): (error: unknown) => void {
  return (error) => console.error(`Footnote engine: ${what}`, error);
}
