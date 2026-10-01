// MAIN world, inside the Claude add-in frame: serves PowerPoint ops to the relay and installs the code-run global.
// Only packaged code runs here, so the frame's CSP doesn't apply.

import { powerpointOps } from "@footnote/powerpoint/ops";
import { callOp, readOfficeInfo, runCodeBody } from "@footnote/shell-kit/realm";
import { defineContentScript } from "wxt/utils/define-content-script";
import {
  ADDIN_MATCHES,
  BRIDGE_ATTRIBUTE,
  INFO_OP,
  REALM_GLOBAL,
  onBridgeRequest,
  postToRelay,
} from "../lib/protocol.ts";

const OFFICE_POLL_MS = 250;
/** Frames that haven't loaded Office.js by then aren't add-in frames. */
const OFFICE_WAIT_MS = 60_000;

export default defineContentScript({
  matches: ADDIN_MATCHES,
  world: "MAIN",
  allFrames: true,
  async main() {
    if (!(await officeLoaded())) return;
    await Office.onReady();
    const { hostName } = readOfficeInfo();
    if (!hostName) return;

    const ops = { ...powerpointOps, [INFO_OP]: async () => readOfficeInfo() };
    onBridgeRequest(async ({ id, op, args }) =>
      postToRelay({ id, ...(await callOp(ops, op, args)) }),
    );
    Object.assign(globalThis, { [REALM_GLOBAL]: { runCode: runCodeBody } });
    document.documentElement.setAttribute(BRIDGE_ATTRIBUTE, hostName);
  },
});

async function officeLoaded(): Promise<boolean> {
  for (let waited = 0; waited < OFFICE_WAIT_MS; waited += OFFICE_POLL_MS) {
    if (typeof Office !== "undefined") return true;
    await new Promise((resolve) => setTimeout(resolve, OFFICE_POLL_MS));
  }
  return false;
}
