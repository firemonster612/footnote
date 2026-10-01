// Service worker: owns the side panel's per-tab behavior and the engine's offscreen document, and runs the chrome.*
// calls the engine can't make itself (see lib/worker.ts). It keeps no state that matters, so Chrome can stop it any time.

import type { CodeRunResult } from "@footnote/core";
import { defineBackground } from "wxt/utils/define-background";
import { BRIDGE_ATTRIBUTE, CONNECT_BRIDGE_MESSAGE } from "../lib/protocol.ts";
import { serveWorker } from "../lib/worker.ts";

const ENGINE_URL = "engine.html";
const PANEL_PATH = "sidepanel.html";

export default defineBackground(() => {
  // The panel is enabled per tab (the one it was opened on), so it hides when the user switches to another tab and
  // comes back with that tab. The toolbar button enables it for the current tab before opening it.
  chrome.sidePanel
    .setPanelBehavior({ openPanelOnActionClick: false })
    .then(() => chrome.sidePanel.setOptions({ enabled: false }))
    .catch((error: unknown) => console.error("Footnote: couldn't set up the side panel", error));

  chrome.action.onClicked.addListener((tab) => {
    if (tab.id === undefined) return;
    // Not awaited: open() must run while Chrome still counts the click as a user gesture.
    void chrome.sidePanel.setOptions({
      tabId: tab.id,
      path: `${PANEL_PATH}?tab=${tab.id}`,
      enabled: true,
    });
    chrome.sidePanel.open({ tabId: tab.id }).catch((error: unknown) => {
      console.error("Footnote: couldn't open the side panel", error);
    });
  });

  let engineStarting: Promise<void> | undefined;

  serveWorker({
    startEngine() {
      engineStarting ??= startEngine().finally(() => (engineStarting = undefined));
      return engineStarting;
    },
    async stopEngine() {
      if (await chrome.offscreen.hasDocument()) await chrome.offscreen.closeDocument();
    },
    async findBridges(tabId) {
      const tab = await chrome.tabs.get(tabId).catch(() => undefined);
      const injections = await chrome.scripting
        .executeScript({
          target: { tabId, allFrames: true },
          func: (attribute: string) => document.documentElement.getAttribute(attribute),
          args: [BRIDGE_ATTRIBUTE],
        })
        .catch(() => []);
      const frames = injections.flatMap(({ frameId, result }) =>
        typeof result === "string" ? [{ frameId, hostName: result }] : [],
      );
      return { ...(tab?.url && { tabUrl: tab.url }), frames };
    },
    async connectBridge(tabId, frameId) {
      await chrome.tabs.sendMessage(tabId, { type: CONNECT_BRIDGE_MESSAGE }, { frameId });
    },
    userScriptsAvailable: async () => chrome.userScripts?.execute !== undefined,
    executeUserScript: (tabId, frameId, code) =>
      chrome.userScripts.execute<CodeRunResult>({
        target: { tabId, frameIds: [frameId] },
        world: "MAIN",
        injectImmediately: true,
        js: [{ code }],
      }),
    async setBadge(tabId, text) {
      // The tab may have closed since the engine decided to badge it.
      await chrome.action.setBadgeText({ tabId, text }).catch(() => {});
    },
    storageGet: async (key) => (await chrome.storage.local.get(key))[key],
    storageSet: (key, value) => chrome.storage.local.set({ [key]: value }),
    storageDelete: (key) => chrome.storage.local.remove(key),
  });
});

async function startEngine(): Promise<void> {
  if (await chrome.offscreen.hasDocument()) return;
  await chrome.offscreen.createDocument({
    url: ENGINE_URL,
    reasons: [chrome.offscreen.Reason.WORKERS],
    justification: "Runs Footnote's agent so a task keeps going while the side panel is closed.",
  });
}
