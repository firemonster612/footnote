// OfficeHost for the side panel: finds the bridge frame in the active tab and talks to it through the relay.

import type {
  CodeRunResult,
  OfficeCallOptions,
  OfficeHost,
  OfficeHostKind,
  OfficeHostStatus,
} from "@footnote/core";
import {
  hostNames,
  statusFromInfo,
  unknownOutcome,
  unwrapOpResponse,
  withDeadline,
} from "@footnote/shell-kit/host";
import { codeBodySource, type OfficeInfo, type OpResponse } from "@footnote/shell-kit/realm";
import {
  BRIDGE_ATTRIBUTE,
  BRIDGE_READY_MESSAGE,
  INFO_OP,
  PORT_NAME,
  REALM_GLOBAL,
  type BridgeRequest,
  type BridgeResponse,
} from "./protocol.ts";

const INFO_TIMEOUT_MS = 5_000;
const RECONNECT_WAIT_MS = 8_000;
const RECONNECT_POLL_MS = 1_000;
const USER_SCRIPTS_OFF =
  'Running code needs Chrome\'s user scripts permission (Chrome 135 or later). Open chrome://extensions, choose Details on Footnote, turn on "Allow User Scripts", then try again.';

interface BridgeFrame {
  tabId: number;
  tabUrl?: string;
  frameId: number;
}

interface Connection extends BridgeFrame {
  port: chrome.runtime.Port;
  pending: Map<string, (response: OpResponse) => void>;
}

export function createExtensionHost(kind: OfficeHostKind): OfficeHost {
  const notFoundReason = `Open the Claude add-in in ${hostNames[kind]} to connect.`;
  const notFound: OfficeHostStatus = { connected: false, reason: notFoundReason };
  let status = notFound;
  let connection: Connection | undefined;
  let latestRefresh = 0;
  const listeners = new Set<(status: OfficeHostStatus) => void>();

  function setStatus(next: OfficeHostStatus): void {
    status = next;
    for (const listener of listeners) listener(next);
  }

  /** Finds a frame in the active tab where the bridge is ready, preferring one running the expected Office host. */
  async function findBridgeFrame(): Promise<BridgeFrame | undefined> {
    const [tab] = await chrome.tabs.query({ active: true, currentWindow: true });
    if (tab?.id === undefined) return undefined;
    const frames = await chrome.scripting
      .executeScript({
        target: { tabId: tab.id, allFrames: true },
        func: (attribute: string) => document.documentElement.getAttribute(attribute),
        args: [BRIDGE_ATTRIBUTE],
      })
      // Pages we can't script (chrome://, the Web Store) can't host the add-in either.
      .catch(() => []);
    const ready = frames.filter((frame) => typeof frame.result === "string");
    const frame = ready.find((candidate) => candidate.result === hostNames[kind]) ?? ready[0];
    return frame && { tabId: tab.id, frameId: frame.frameId, ...(tab.url && { tabUrl: tab.url }) };
  }

  function connect(frame: BridgeFrame): Connection {
    const port = chrome.tabs.connect(frame.tabId, { frameId: frame.frameId, name: PORT_NAME });
    const opened: Connection = { ...frame, port, pending: new Map() };
    let answered = false;
    port.onMessage.addListener((response: BridgeResponse) => {
      answered = true;
      opened.pending.get(response.id)?.(response);
    });
    port.onDisconnect.addListener(() => {
      close(opened);
      // A port that never answered has no relay behind it (e.g. the extension was reloaded but the tab wasn't);
      // reconnecting would just loop. The failed request reports it instead.
      if (answered) void refresh();
    });
    return opened;
  }

  function close(closing: Connection): void {
    if (connection === closing) connection = undefined;
    closing.port.disconnect();
    for (const settle of closing.pending.values()) {
      settle({
        ok: false,
        error: {
          message:
            "Lost the connection to the add-in frame. Reload the PowerPoint tab if it doesn't reconnect.",
        },
      });
    }
  }

  async function request<T>(
    target: Connection,
    op: string,
    args: unknown,
    options?: OfficeCallOptions,
  ): Promise<T> {
    const id = crypto.randomUUID();
    const response = new Promise<OpResponse>((resolve) => target.pending.set(id, resolve));
    try {
      target.port.postMessage({ id, op, args } satisfies BridgeRequest);
      return unwrapOpResponse<T>(op, await withDeadline(response, `Op "${op}"`, options));
    } finally {
      target.pending.delete(id);
    }
  }

  /** Re-discovers the bridge frame; only the latest of overlapping refreshes applies its result. */
  async function refresh(): Promise<void> {
    const run = ++latestRefresh;
    const isLatest = () => run === latestRefresh;
    try {
      const frame = await findBridgeFrame();
      if (!isLatest()) return;
      if (!frame) {
        if (connection) close(connection);
        return setStatus(notFound);
      }
      if (connection?.tabId !== frame.tabId || connection.frameId !== frame.frameId) {
        if (connection) close(connection);
        connection = connect(frame);
      }
      const info = await request<OfficeInfo>(connection, INFO_OP, undefined, {
        timeoutMs: INFO_TIMEOUT_MS,
      });
      if (isLatest()) setStatus(statusFromInfo(info, kind, frame.tabUrl));
    } catch (error) {
      if (isLatest())
        setStatus({
          connected: false,
          reason: `Couldn't reach the add-in: ${errorMessage(error)}`,
        });
    }
  }

  /** The add-in frame drops its port for a moment when the pane reloads; wait for it instead of failing the tool call. */
  async function activeConnection(): Promise<Connection> {
    const deadline = Date.now() + RECONNECT_WAIT_MS;
    while (!connection || !status.connected) {
      if (Date.now() >= deadline) throw new Error(status.reason ?? notFoundReason);
      await refresh();
      if (connection && status.connected) break;
      await new Promise((resolve) => setTimeout(resolve, RECONNECT_POLL_MS));
    }
    return connection;
  }

  chrome.tabs.onActivated.addListener(() => void refresh());
  chrome.tabs.onUpdated.addListener((_tabId, change, tab) => {
    if (tab.active && change.status === "complete") void refresh();
  });
  chrome.runtime.onMessage.addListener((message: { type?: unknown }, sender) => {
    if (message.type === BRIDGE_READY_MESSAGE && sender.tab?.active) void refresh();
  });
  const firstRefresh = refresh();

  return {
    async status() {
      await firstRefresh;
      return status;
    },
    onStatusChange(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
    async call(op, args, options) {
      return request(await activeConnection(), op, args, options);
    },
    async runCode(code, options): Promise<CodeRunResult> {
      let target: Connection;
      try {
        target = await activeConnection();
      } catch (error) {
        return codeFailure(errorMessage(error));
      }
      if (!chrome.userScripts?.execute) return codeFailure(USER_SCRIPTS_OFF);
      try {
        const execution = chrome.userScripts.execute<CodeRunResult>({
          target: { tabId: target.tabId, frameIds: [target.frameId] },
          world: "MAIN",
          injectImmediately: true,
          js: [{ code: `globalThis.${REALM_GLOBAL}.runCode(${codeBodySource(code)})` }],
        });
        const [injection] = await withDeadline(execution, "The code run", options);
        if (!injection)
          return unknownOutcome(new Error("Chrome returned no result for the code run"));
        // An injection error means the code never started (syntax error, bridge missing), so the outcome is known.
        return injection.error === undefined ? injection.result : codeFailure(injection.error);
      } catch (error) {
        return unknownOutcome(error);
      }
    },
  };
}

function codeFailure(message: string): CodeRunResult {
  return { ok: false, logs: [], error: { message } };
}

function errorMessage(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}
