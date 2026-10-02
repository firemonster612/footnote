import type { OpResponse } from "@footnote/shell-kit/realm";
import { afterEach, beforeEach, describe, expect, it, vi } from "vitest";
import { createExtensionHost } from "../lib/extension-host.ts";
import { INFO_OP, PORT_NAME, type BridgeRequest, type BridgeResponse } from "../lib/protocol.ts";
import type { WorkerApi } from "../lib/worker.ts";

const tabId = 7;
const frameId = 3;
const officeInfo = {
  hostName: "PowerPoint",
  documentUrl: "https://x.test/deck.pptx",
  apiVersions: {},
};

/** The add-in frame's end of a relay port. `answer` returns undefined to leave a request unanswered. */
function fakeBridge(answer: (request: BridgeRequest) => OpResponse | undefined) {
  const messageListeners: ((response: BridgeResponse) => void)[] = [];
  const disconnectListeners: (() => void)[] = [];
  const requests: BridgeRequest[] = [];
  const port = {
    name: PORT_NAME,
    sender: { tab: { id: tabId }, frameId },
    postMessage(request: BridgeRequest) {
      requests.push(request);
      const response =
        request.op === INFO_OP ? { ok: true as const, value: officeInfo } : answer(request);
      if (!response) return;
      setTimeout(() => {
        for (const listener of messageListeners) listener({ ...response, id: request.id });
      });
    },
    onMessage: {
      addListener: (listener: (response: BridgeResponse) => void) =>
        void messageListeners.push(listener),
    },
    onDisconnect: {
      addListener: (listener: () => void) => void disconnectListeners.push(listener),
    },
    disconnect() {},
    drop() {
      for (const listener of disconnectListeners) listener();
    },
  };
  return { port, requests };
}

/** Installs a `chrome` whose service worker serves `worker` and whose relay opens `bridge`'s port on request. */
function fakeChrome(bridge: ReturnType<typeof fakeBridge>, worker: Partial<WorkerApi> = {}) {
  const connectListeners: ((port: unknown) => void)[] = [];
  const api: Partial<WorkerApi> = {
    findBridges: async () => ({ frames: [{ frameId, hostName: "PowerPoint" }] }),
    connectBridge: async () => {
      for (const listener of connectListeners) listener(bridge.port);
    },
    userScriptsAvailable: async () => true,
    ...worker,
  };
  vi.stubGlobal("chrome", {
    runtime: {
      onConnect: {
        addListener: (listener: (port: unknown) => void) => void connectListeners.push(listener),
      },
      onMessage: { addListener: () => {} },
      async sendMessage({ method, args }: { method: keyof WorkerApi; args: unknown[] }) {
        try {
          const handler = api[method] as ((...args: unknown[]) => unknown) | undefined;
          if (!handler) throw new Error(`no fake for ${method}`);
          return { ok: true, value: await handler(...args) };
        } catch (error) {
          return { ok: false, error: (error as Error).message };
        }
      },
    },
  });
}

describe("extension host", () => {
  beforeEach(() => {
    vi.useFakeTimers();
    vi.stubGlobal("localStorage", { getItem: () => null });
  });
  afterEach(() => {
    vi.useRealTimers();
    vi.unstubAllGlobals();
  });

  it("waits for the add-in frame to come back, then makes the call", async () => {
    const bridge = fakeBridge(() => ({ ok: true, value: "deck" }));
    let frameReady = false;
    fakeChrome(bridge, {
      findBridges: async () => ({
        frames: frameReady ? [{ frameId, hostName: "PowerPoint" }] : [],
      }),
    });
    const host = createExtensionHost("powerpoint", tabId);

    const call = host.call("get_deck");
    await vi.advanceTimersByTimeAsync(500);
    frameReady = true;
    await vi.advanceTimersByTimeAsync(1_000);

    await expect(call).resolves.toBe("deck");
  });

  it("rejects with a TimeoutError when the add-in doesn't answer in time", async () => {
    const bridge = fakeBridge(() => undefined);
    fakeChrome(bridge);
    const host = createExtensionHost("powerpoint", tabId);

    const call = host.call("get_deck", undefined, { timeoutMs: 2_000 });
    const settled = expect(call).rejects.toMatchObject({ name: "TimeoutError" });
    await vi.advanceTimersByTimeAsync(2_000);

    await settled;
  });

  it("rejects with an OutcomeUnknownError when the port drops after the request was sent", async () => {
    const bridge = fakeBridge(() => undefined);
    fakeChrome(bridge);
    const host = createExtensionHost("powerpoint", tabId);

    const call = host.call("add_slide");
    const settled = expect(call).rejects.toMatchObject({ name: "OutcomeUnknownError" });
    await vi.advanceTimersByTimeAsync(0);
    expect(bridge.requests.map((request) => request.op)).toContain("add_slide");
    bridge.port.drop();

    await settled;
  });

  it("doesn't dispatch code whose run was already cancelled", async () => {
    const executeUserScript = vi.fn(async () => []);
    fakeChrome(
      fakeBridge(() => undefined),
      { executeUserScript },
    );
    const host = createExtensionHost("powerpoint", tabId);
    const controller = new AbortController();
    controller.abort();

    const result = await host.runCode("return 1", { signal: controller.signal });

    expect(result).toMatchObject({ ok: false });
    expect(result.outcomeUnknown).toBeUndefined();
    expect(executeUserScript).not.toHaveBeenCalled();
  });

  it("stops waiting for the add-in frame when the code run is cancelled", async () => {
    const executeUserScript = vi.fn(async () => []);
    fakeChrome(
      fakeBridge(() => undefined),
      {
        findBridges: async () => ({ frames: [] }),
        executeUserScript,
      },
    );
    const host = createExtensionHost("powerpoint", tabId);
    const controller = new AbortController();
    let settled = false;

    const run = host.runCode("return 1", { signal: controller.signal }).finally(() => {
      settled = true;
    });
    await vi.advanceTimersByTimeAsync(500);
    controller.abort();
    await vi.advanceTimersByTimeAsync(0);

    expect(settled).toBe(true);
    expect(await run).toMatchObject({ ok: false });
    expect(executeUserScript).not.toHaveBeenCalled();
  });
});
