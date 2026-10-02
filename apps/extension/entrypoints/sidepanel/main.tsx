// The side panel is a view: the agent runs in the engine (an offscreen document) and keeps going when the panel closes.

import { connectFootnoteApp } from "@footnote/core";
import { FootnoteRoot } from "@footnote/core/ui";
import { powerpointModule } from "@footnote/powerpoint";
import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { VIEW_PORT_PREFIX } from "../../lib/protocol.ts";
import { askWorker } from "../../lib/worker.ts";
import "./style.css";

const FIRST_RETRY_MS = 1_000;
const MAX_RETRY_MS = 60_000;
// Survives the reloads that retry, so an engine that keeps failing is retried less and less often.
const FAILED_ATTEMPTS_KEY = "footnote:failed-connects";

// index.html always contains #root.
const root = createRoot(document.getElementById("root")!);
// A failed connect both rejects and disconnects; retry once.
let retrying = false;

/** Shows why the panel can't reach the engine and reloads after a backoff that doubles with each failure. */
function retryLater(error: unknown): void {
  if (retrying) return;
  retrying = true;
  const failedAttempts = Number(sessionStorage.getItem(FAILED_ATTEMPTS_KEY) ?? 0);
  sessionStorage.setItem(FAILED_ATTEMPTS_KEY, String(failedAttempts + 1));
  const delayMs = Math.min(FIRST_RETRY_MS * 2 ** failedAttempts, MAX_RETRY_MS);
  const reason = error instanceof Error ? error.message : String(error);
  root.render(
    <p className="p-4 text-sm text-muted-foreground">
      {`Footnote couldn't connect: ${reason} Retrying in ${Math.round(delayMs / 1000)} s.`}
    </p>,
  );
  setTimeout(() => location.reload(), delayMs);
}

async function start(): Promise<void> {
  // The service worker opens the panel with ?tab=<id>; fall back to the active tab if Chrome opened it some other way.
  const tabId =
    Number(new URLSearchParams(location.search).get("tab")) ||
    (await chrome.tabs.query({ active: true, currentWindow: true }))[0]?.id;
  if (tabId === undefined) throw new Error("the side panel has no tab to attach to.");

  await askWorker("startEngine");
  const port = chrome.runtime.connect({ name: `${VIEW_PORT_PREFIX}${tabId}` });
  // The engine closed or restarted (e.g. the extension updated): start over against the new one.
  port.onDisconnect.addListener(() =>
    retryLater(new Error("Lost the connection to Footnote's engine.")),
  );
  const app = await connectFootnoteApp(port, powerpointModule);
  sessionStorage.removeItem(FAILED_ATTEMPTS_KEY);
  root.render(
    <StrictMode>
      <FootnoteRoot app={app} />
    </StrictMode>,
  );
}

start().catch(retryLater);
