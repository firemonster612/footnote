import type { CodeRunResult, FootnoteTool, ToolEnv } from "@footnote/core/contracts";
import { Type } from "typebox";
import { codeHelpersReference } from "../advanced/index.ts";
import type { WriteGuard } from "../writeGuard.ts";
import { defineTool, SlideId } from "./schemas.ts";

const MAX_CODE_CHARS = 20_000;
const MAX_RESULT_CHARS = 8_000;
/** PowerPoint for the web applies large batches slowly; a 60s cap reported real, still-landing edits as failures. */
const CODE_TIMEOUT_MS = 300_000;

const clip = (text: string): string =>
  text.length > MAX_RESULT_CHARS
    ? `${text.slice(0, MAX_RESULT_CHARS)}\n… truncated (${text.length} chars); return less data.`
    : text;

/** Model-facing text for a code run. Exported for tests. */
export function formatCodeRun(run: CodeRunResult): string {
  const logs = run.logs.length > 0 ? `\nLogs:\n${run.logs.join("\n")}` : "";
  if (run.outcomeUnknown) {
    const reason = run.error ? ` (${run.error.message})` : "";
    return clip(
      `Outcome unknown${reason}: the code timed out or the connection dropped, so edits may or may not have applied. Do not rerun it: call get_deck or get_slide first, because the edits often landed.${logs}`,
    );
  }
  if (!run.ok) {
    const code = run.error?.code ? ` [${run.error.code}]` : "";
    return clip(`Error${code}: ${run.error?.message ?? "unknown error"}${logs}`);
  }
  const result = run.result === undefined ? "undefined" : JSON.stringify(run.result);
  return clip(`Result: ${result}${logs}`);
}

export function createCodeTool(env: ToolEnv, guard: WriteGuard): FootnoteTool {
  return defineTool({
    name: "execute_office_js",
    label: "Run Office.js",
    access: "code",
    executionMode: "sequential",
    description: `Runs an async function body inside PowerPoint with \`context\` (PowerPoint.RequestContext) and \`footnote\` (helpers) in scope. \
Use it for what the structured tools can't do (bulk edits across many slides, properties they don't expose, reading something unusual), not for edits a structured tool covers. \
Load before reading (\`.load("…")\` then \`await context.sync()\`), sync after writing, and \`return\` a small JSON-safe value; \`console.log\` output is returned too. \
Results over ${MAX_RESULT_CHARS} characters are cut. PowerPoint for the web applies big batches slowly: \`await context.sync()\` after each slide's edits rather than once at the end, and keep one run to a few slides. \
List the slides the code will change in slideIds so undo can restore them. If the outcome is unknown, re-read before retrying.

${codeHelpersReference}`,
    parameters: Type.Object({
      code: Type.String({
        maxLength: MAX_CODE_CHARS,
        description: "Body of an async function. Use await and return.",
      }),
      explanation: Type.String({ description: "One sentence for the user: what the code does." }),
      slideIds: Type.Optional(
        Type.Array(SlideId, {
          description: "Slides the code modifies or deletes, snapshotted for undo first.",
        }),
      ),
    }),
    describeCall: (args: { explanation: string }) => args.explanation,
    async execute(_id, { code, slideIds = [] }, signal) {
      if (slideIds.length > 0) await guard.checkpoint(slideIds);
      const run = await env.host.runCode(code, { timeoutMs: CODE_TIMEOUT_MS, ...(signal && { signal }) });
      return {
        content: [{ type: "text", text: formatCodeRun(run) }],
        details: run,
        ...(!run.ok && { isError: true }),
      };
    },
  });
}
