// OfficeHost for the engine, bound to one tab: finds the bridge frame there and talks to it through the relay.
// The tab stays fixed whichever tab is active, so a running chat keeps talking to the document it started on.

import type {
  CodeRunResult,
  OfficeCallOptions,
  OfficeHost,
  OfficeHostKind,
  OfficeHostStatus,
} from "@footnote/core";
import {
  hostNames,
  logOpTiming,
  statusFromInfo,
  unknownOutcome,
  unwrapOpResponse,
  withDeadline,
} from "@footnote/shell-kit/host";
import { codeBodySource, type OfficeInfo, type OpResponse } from "@footnote/shell-kit/realm";
import {
  BRIDGE_READY_MESSAGE,
  INFO_OP,
  PORT_NAME,
  REALM_GLOBAL,
  type BridgeRequest,
  type BridgeResponse,
} from "./protocol.ts";
import { askWorker } from "./worker.ts";

const INFO_TIMEOUT_MS = 5_000;
const CONNECT_TIMEOUT_MS = 5_000;
const RECONNECT_WAIT_MS = 8_000;
const RECONNECT_POLL_MS = 1_000;
const USER_SCRIPTS_OFF =
  'Running code needs Chrome\'s user scripts permission (Chrome 135 or later). Open chrome://extensions, choose Details on Footnote, turn on "Allow User Scripts", then try again.';

interface BridgeFrame {
  tabUrl?: string;
  frameId: number;
}

interface Connection {
  frameId: number;
  port: chrome.runtime.Port;
  pending: Map<string, (response: OpResponse) => void>;
}

export function createExtensionHost(kind: OfficeHostKind, tabId: number): OfficeHost {
  const notFoundReason = `Open the Claude add-in in ${hostNames[kind]} to connect.`;
  const notFound: OfficeHostStatus = { connected: false, reason: notFoundReason };
  let status = notFound;
  let connection: Connection | undefined;
  let latestRefresh = 0;
  const listeners = new Set<(status: OfficeHostStatus) => void>();
  const connectionListeners = new Set<(opened: Connection) => void>();

  function setStatus(next: OfficeHostStatus): void {
    status = next;
    for (const listener of listeners) listener(next);
  }

  /** Finds a frame in the tab where the bridge is ready, preferring one running the expected Office host. */
  async function findBridgeFrame(): Promise<BridgeFrame | undefined> {
    const { tabUrl, frames } = await askWorker("findBridges", tabId);
    const frame = frames.find((candidate) => candidate.hostName === hostNames[kind]) ?? frames[0];
    return frame && { frameId: frame.frameId, ...(tabUrl && { tabUrl }) };
  }

  /** The engine can't open ports into tabs, so it asks the frame's relay to open one (see adopt). */
  async function connect(frameId: number): Promise<Connection> {
    let onOpened: (opened: Connection) => void = () => {};
    const opened = new Promise<Connection>((resolve) => {
      onOpened = (candidate) => candidate.frameId === frameId && resolve(candidate);
      connectionListeners.add(onOpened);
    });
    try {
      await askWorker("connectBridge", tabId, frameId).catch((error: unknown) => {
        // No relay listening: the extension was reloaded or updated after the tab loaded.
        throw new Error(
          `the add-in frame didn't answer (${errorMessage(error)}). Reload the PowerPoint tab.`,
        );
      });
      return await withDeadline(opened, "Connecting to the add-in frame", {
        timeoutMs: CONNECT_TIMEOUT_MS,
      });
    } finally {
      connectionListeners.delete(onOpened);
    }
  }

  /** A relay port arrived: it replaces the current connection, whoever asked for it. */
  function adopt(port: chrome.runtime.Port, frameId: number): void {
    if (connection) close(connection);
    const opened: Connection = { frameId, port, pending: new Map() };
    connection = opened;
    let answered = false;
    port.onMessage.addListener((response: BridgeResponse) => {
      answered = true;
      opened.pending.get(response.id)?.(response);
    });
    port.onDisconnect.addListener(() => {
      close(opened);
      // A port that never answered has no bridge behind it; reconnecting would just loop. The failed request
      // reports it instead.
      if (answered) void refresh();
    });
    for (const listener of connectionListeners) listener(opened);
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
    const startedAt = performance.now();
    try {
      target.port.postMessage({ id, op, args } satisfies BridgeRequest);
      const settled = await withDeadline(response, `Op "${op}"`, options);
      logOpTiming(op, startedAt, settled);
      return unwrapOpResponse<T>(op, settled);
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
      const target =
        connection?.frameId === frame.frameId ? connection : await connect(frame.frameId);
      if (!isLatest()) return;
      const info = await request<OfficeInfo>(target, INFO_OP, undefined, {
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

  chrome.runtime.onConnect.addListener((port) => {
    const { tab, frameId } = port.sender ?? {};
    if (port.name === PORT_NAME && tab?.id === tabId && frameId !== undefined) adopt(port, frameId);
  });
  // Reloading the tab or reopening the add-in drops the port; the new bridge announces itself.
  chrome.runtime.onMessage.addListener((message: { type?: unknown }, sender) => {
    if (message.type === BRIDGE_READY_MESSAGE && sender.tab?.id === tabId) void refresh();
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
        if (!(await askWorker("userScriptsAvailable"))) return codeFailure(USER_SCRIPTS_OFF);
      } catch (error) {
        return codeFailure(errorMessage(error));
      }
      try {
        const execution = askWorker(
          "executeUserScript",
          tabId,
          target.frameId,
          `globalThis.${REALM_GLOBAL}.runCode(${codeBodySource(code)})`,
        );
        const startedAt = performance.now();
        const [injection] = await withDeadline(execution, "The code run", options);
        logOpTiming("execute_office_js", startedAt);
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
