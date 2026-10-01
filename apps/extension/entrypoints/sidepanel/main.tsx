// The side panel is a view: the agent runs in the engine (an offscreen document) and keeps going when the panel closes.

import { connectFootnoteApp } from "@footnote/core";
import { FootnoteRoot } from "@footnote/core/ui";
import { powerpointModule } from "@footnote/powerpoint";
import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { VIEW_PORT_PREFIX } from "../../lib/protocol.ts";
import { askWorker } from "../../lib/worker.ts";
import "./style.css";

const RECONNECT_DELAY_MS = 1_000;

// The service worker opens the panel with ?tab=<id>; fall back to the active tab if Chrome opened it some other way.
const tabId =
  Number(new URLSearchParams(location.search).get("tab")) ||
  (await chrome.tabs.query({ active: true, currentWindow: true }))[0]?.id;
if (tabId === undefined) throw new Error("Footnote: the side panel has no tab to attach to.");

await askWorker("startEngine");
const port = chrome.runtime.connect({ name: `${VIEW_PORT_PREFIX}${tabId}` });
// The engine closed or restarted (e.g. the extension updated): start over against the new one.
port.onDisconnect.addListener(() => setTimeout(() => location.reload(), RECONNECT_DELAY_MS));
const app = await connectFootnoteApp(port, powerpointModule);

// index.html always contains #root.
createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <FootnoteRoot app={app} />
  </StrictMode>,
);
