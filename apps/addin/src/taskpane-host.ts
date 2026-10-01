// OfficeHost for the task pane: Office.js runs in this page, so ops and code run in-process.

import type { CodeRunResult, HostModule, OfficeHost } from "@footnote/core";
import {
  statusFromInfo,
  unknownOutcome,
  unwrapOpResponse,
  withDeadline,
} from "@footnote/shell-kit/host";
import {
  callOp,
  codeBodySource,
  readOfficeInfo,
  runCodeBody,
  toErrorInfo,
  type CodeBody,
} from "@footnote/shell-kit/realm";

/** Chat history key for documents without a URL (never saved). They share one history. */
const UNSAVED_DOCUMENT_ID = "unsaved-document";

export function createTaskPaneHost(hostModule: HostModule): OfficeHost {
  return {
    status: async () => statusFromInfo(readOfficeInfo(), hostModule.kind, UNSAVED_DOCUMENT_ID),
    // The task pane lives and dies with its document, so there are no status changes to report.
    onStatusChange: () => () => {},
    call: async (op, args, options) =>
      unwrapOpResponse(
        op,
        await withDeadline(callOp(hostModule.ops, op, args), `Op "${op}"`, options),
      ),
    async runCode(code, options): Promise<CodeRunResult> {
      let body: CodeBody;
      try {
        body = new Function(`return ${codeBodySource(code)}`)();
      } catch (error) {
        return { ok: false, logs: [], error: toErrorInfo(error) };
      }
      try {
        return await withDeadline(runCodeBody(body), "The code run", options);
      } catch (error) {
        return unknownOutcome(error);
      }
    },
  };
}
