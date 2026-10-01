// Messages between the side panel, the ISOLATED-world relay, and the MAIN-world bridge in the add-in frame.
//
// side panel ──chrome port (tab+frame)──▶ relay ──window.postMessage──▶ bridge
//            ◀───────────────────────────       ◀─────────────────────

import type { OpResponse } from "@footnote/shell-kit/realm";

/** Frames the bridge serves. Extra origins come from FOOTNOTE_ADDIN_MATCHES at build time (see wxt.config.ts). */
export const ADDIN_MATCHES = ["https://pivot.claude.ai/*"];

export const PORT_NAME = "footnote-bridge";

/** Pseudo-op the bridge answers with `OfficeInfo`. */
export const INFO_OP = "footnote.info";

/** Set on <html> by the bridge once Office is ready; value is `Office.context.host`. Readable from any world. */
export const BRIDGE_ATTRIBUTE = "data-footnote-bridge";

/** Relay → extension pages: the bridge in this frame became ready. */
export const BRIDGE_READY_MESSAGE = "footnote:bridge-ready";

/** MAIN-world global the bridge installs for `chrome.userScripts.execute` code runs. */
export const REALM_GLOBAL = "__footnoteRealm";

export interface BridgeRequest {
  id: string;
  op: string;
  args?: unknown;
}

export type BridgeResponse = OpResponse & { id: string };

const TO_BRIDGE = "footnote:to-bridge";
const TO_RELAY = "footnote:to-relay";

export function postToBridge(request: BridgeRequest): void {
  window.postMessage({ ...request, channel: TO_BRIDGE }, location.origin);
}

export function postToRelay(response: BridgeResponse): void {
  window.postMessage({ ...response, channel: TO_RELAY }, location.origin);
}

// The page shares this window and could post on our channels too; it already controls Office in its own frame,
// so trusting the channel tag gives it nothing new.
export function onBridgeRequest(listener: (request: BridgeRequest) => void): void {
  onChannel(TO_BRIDGE, (data) => listener(data as BridgeRequest));
}

export function onBridgeResponse(listener: (response: BridgeResponse) => void): void {
  onChannel(TO_RELAY, (data) => listener(data as BridgeResponse));
}

function onChannel(channel: string, listener: (data: object) => void): void {
  window.addEventListener("message", ({ source, data }: MessageEvent<unknown>) => {
    if (source !== window || typeof data !== "object" || data === null) return;
    if ("channel" in data && data.channel === channel) listener(data);
  });
}
