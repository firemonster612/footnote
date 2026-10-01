// ISOLATED world, same frames as the bridge: forwards side-panel port messages to the MAIN-world bridge and back.

import { defineContentScript } from "wxt/utils/define-content-script";
import {
  ADDIN_MATCHES,
  BRIDGE_ATTRIBUTE,
  BRIDGE_READY_MESSAGE,
  PORT_NAME,
  onBridgeResponse,
  postToBridge,
  type BridgeRequest,
} from "../lib/protocol.ts";

export default defineContentScript({
  matches: ADDIN_MATCHES,
  allFrames: true,
  main() {
    const portsByRequest = new Map<string, chrome.runtime.Port>();

    onBridgeResponse((response) => {
      const port = portsByRequest.get(response.id);
      portsByRequest.delete(response.id);
      try {
        port?.postMessage(response);
      } catch {
        // The side panel closed while the op ran; nobody is waiting for this response.
      }
    });

    chrome.runtime.onConnect.addListener((port) => {
      if (port.name !== PORT_NAME) return;
      port.onMessage.addListener((request: BridgeRequest) => {
        portsByRequest.set(request.id, port);
        postToBridge(request);
      });
    });

    whenBridgeReady(() => {
      // Rejects when no side panel is open to hear it; the panel discovers the frame itself when it opens.
      chrome.runtime.sendMessage({ type: BRIDGE_READY_MESSAGE }).catch(() => {});
    });
  },
});

function whenBridgeReady(callback: () => void): void {
  const root = document.documentElement;
  if (root.hasAttribute(BRIDGE_ATTRIBUTE)) return callback();
  const observer = new MutationObserver(() => {
    if (!root.hasAttribute(BRIDGE_ATTRIBUTE)) return;
    observer.disconnect();
    callback();
  });
  observer.observe(root, { attributes: true, attributeFilter: [BRIDGE_ATTRIBUTE] });
}
