// ISOLATED world, same frames as the bridge: forwards the engine's port messages to the MAIN-world bridge and back.

import { defineContentScript } from "wxt/utils/define-content-script";
import {
  ADDIN_MATCHES,
  BRIDGE_ATTRIBUTE,
  BRIDGE_READY_MESSAGE,
  CONNECT_BRIDGE_MESSAGE,
  PORT_NAME,
  onBridgeResponse,
  postToBridge,
  type BridgeRequest,
} from "../lib/protocol.ts";

export default defineContentScript({
  matches: ADDIN_MATCHES,
  allFrames: true,
  main() {
    let engine: chrome.runtime.Port | undefined;

    onBridgeResponse((response) => {
      try {
        engine?.postMessage(response);
      } catch {
        // The engine went away while the op ran; nobody is waiting for this response.
      }
    });

    // The engine can't open ports into tabs, so it asks (through the service worker) and we open one to it.
    // An open port is kept: the engine adopts the first one, and replacing it would drop requests in flight.
    chrome.runtime.onMessage.addListener((message: { type?: unknown }, _sender, reply) => {
      if (message.type !== CONNECT_BRIDGE_MESSAGE) return;
      if (!engine) {
        const opened = chrome.runtime.connect({ name: PORT_NAME });
        opened.onMessage.addListener((request: BridgeRequest) => postToBridge(request));
        opened.onDisconnect.addListener(() => {
          if (engine === opened) engine = undefined;
        });
        engine = opened;
      }
      reply(true);
    });

    whenBridgeReady(() => {
      // Rejects when the engine isn't running; it finds the frame itself when it starts.
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
