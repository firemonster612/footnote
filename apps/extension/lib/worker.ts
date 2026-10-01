// Calls into the service worker for the chrome.* APIs the engine can't reach: an offscreen document only gets
// chrome.runtime. Each call is a one-shot message, so it works across service worker restarts.

import type { CodeRunResult } from "@footnote/core";

/** Frames in a tab whose bridge is ready. */
export interface TabBridges {
  tabUrl?: string;
  frames: { frameId: number; hostName: string }[];
}

export interface WorkerApi {
  /** Creates the engine's offscreen document unless it's already open. */
  startEngine(): Promise<void>;
  stopEngine(): Promise<void>;
  /** Empty when the tab is gone or can't be scripted (chrome:// pages, the Web Store). */
  findBridges(tabId: number): Promise<TabBridges>;
  /** Asks the relay in this frame to open a port to the engine. */
  connectBridge(tabId: number, frameId: number): Promise<void>;
  userScriptsAvailable(): Promise<boolean>;
  executeUserScript(
    tabId: number,
    frameId: number,
    code: string,
  ): Promise<chrome.userScripts.InjectionResult<CodeRunResult>[]>;
  setBadge(tabId: number, text: string): Promise<void>;
  storageGet(key: string): Promise<unknown>;
  storageSet(key: string, value: unknown): Promise<void>;
  storageDelete(key: string): Promise<void>;
}

type WorkerMethod = keyof WorkerApi;
type WorkerReply = { ok: true; value?: unknown } | { ok: false; error: string };

const WORKER_CHANNEL = "footnote:worker";

export async function askWorker<M extends WorkerMethod>(
  method: M,
  ...args: Parameters<WorkerApi[M]>
): Promise<Awaited<ReturnType<WorkerApi[M]>>> {
  const reply: WorkerReply | undefined = await chrome.runtime.sendMessage({
    channel: WORKER_CHANNEL,
    method,
    args,
  });
  if (!reply) throw new Error(`The service worker didn't answer ${method}.`);
  if (!reply.ok) throw new Error(reply.error);
  // serveWorker replies with what the handler for `method` resolved to.
  return reply.value as Awaited<ReturnType<WorkerApi[M]>>;
}

export function serveWorker(api: WorkerApi): void {
  chrome.runtime.onMessage.addListener(
    (message: { channel?: unknown; method: WorkerMethod; args: unknown[] }, _sender, reply) => {
      if (message.channel !== WORKER_CHANNEL) return false;
      Promise.resolve(Reflect.apply(api[message.method], api, message.args)).then(
        (value: unknown) => reply({ ok: true, value } satisfies WorkerReply),
        (error: unknown) =>
          reply({
            ok: false,
            error: error instanceof Error ? error.message : String(error),
          } satisfies WorkerReply),
      );
      return true; // replies asynchronously
    },
  );
}
