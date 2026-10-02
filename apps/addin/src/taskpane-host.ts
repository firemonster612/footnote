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
  codeLineOffsets,
  withSourceUrl,
  toErrorInfo,
  type CodeBody,
} from "@footnote/shell-kit/realm";

export function createTaskPaneHost(hostModule: HostModule): OfficeHost {
  // Chat history key while the document has no URL (never saved). The pane belongs to one document, so this
  // keeps its chats apart from other unsaved decks; once saved, the URL takes over.
  const unsavedDocumentId = `unsaved-${crypto.randomUUID()}`;
  return {
    status: async () => statusFromInfo(readOfficeInfo(), hostModule.kind, unsavedDocumentId),
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
        body = new Function(withSourceUrl(`return ${codeBodySource(code)}`))();
      } catch (error) {
        return { ok: false, logs: [], error: toErrorInfo(error) };
      }
      try {
        return await withDeadline(
          runCodeBody(body, { code, lineOffset: codeLineOffsets.newFunction }),
          "The code run",
          options,
        );
      } catch (error) {
        return unknownOutcome(error);
      }
    },
  };
}
