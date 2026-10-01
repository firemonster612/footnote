// Runs inside the Office realm: the add-in task pane page, or the extension's MAIN-world bridge in the Claude add-in frame.
// Keep it free of agent-side code and chrome.* APIs.

import type { CodeRunResult, OfficeErrorInfo, OpRegistry } from "@footnote/core/contracts";
import { createFootnoteHelpers } from "@footnote/powerpoint/ops";

/** What the realm knows about the Office document it's attached to. */
export interface OfficeInfo {
  /** `Office.context.host`, e.g. "PowerPoint". Undefined when the page isn't running inside Office. */
  hostName?: string;
  documentUrl?: string;
  apiVersions: Record<string, string>;
}

export type OpResponse = { ok: true; value: unknown } | { ok: false; error: OfficeErrorInfo };

type CodeConsole = Pick<Console, "log" | "info" | "warn" | "error">;

/** A compiled `execute_office_js` body. */
export type CodeBody = (
  context: PowerPoint.RequestContext,
  footnote: ReturnType<typeof createFootnoteHelpers>,
  console: CodeConsole,
) => Promise<unknown>;

/** Highest minor version probed per requirement set. */
const MAX_API_MINOR = 20;

export function readOfficeInfo(): OfficeInfo {
  // Typed as a numeric enum, but office.js sets string values such as "PowerPoint" (and null outside Office).
  const hostName = Office.context.host ? String(Office.context.host) : undefined;
  const apiSet = hostName && `${hostName}Api`;
  const minor = apiSet ? highestSupportedMinor(apiSet) : 0;
  return {
    ...(hostName && { hostName }),
    ...(Office.context.document?.url && { documentUrl: Office.context.document.url }),
    apiVersions: apiSet && minor > 0 ? { [apiSet]: `1.${minor}` } : {},
  };
}

function highestSupportedMinor(apiSet: string): number {
  for (let minor = MAX_API_MINOR; minor > 0; minor--) {
    if (Office.context.requirements.isSetSupported(apiSet, `1.${minor}`)) return minor;
  }
  return 0;
}

export async function callOp(ops: OpRegistry, op: string, args: unknown): Promise<OpResponse> {
  const run = ops[op];
  if (!run) return { ok: false, error: { message: `Unknown op "${op}"` } };
  try {
    return { ok: true, value: await run(args) };
  } catch (error) {
    return { ok: false, error: toErrorInfo(error) };
  }
}

/** Source of an async function expression that compiles to a `CodeBody`. */
export function codeBodySource(code: string): string {
  return `async (context, footnote, console) => {\n${code}\n}`;
}

export async function runCodeBody(body: CodeBody): Promise<CodeRunResult> {
  const logs: string[] = [];
  const log = (...values: unknown[]) => {
    logs.push(values.map(formatLogValue).join(" "));
  };
  const console = { log, info: log, warn: log, error: log };
  try {
    const result = await PowerPoint.run((context) =>
      body(context, createFootnoteHelpers(context), console),
    );
    return { ok: true, result: toJsonSafe(result), logs };
  } catch (error) {
    return { ok: false, logs, error: toErrorInfo(error) };
  }
}

export function toErrorInfo(error: unknown): OfficeErrorInfo {
  if (error instanceof OfficeExtension.Error) {
    return { message: error.message, code: error.code, debugInfo: toJsonSafe(error.debugInfo) };
  }
  return { message: error instanceof Error ? error.message : String(error) };
}

function formatLogValue(value: unknown): string {
  if (typeof value === "string") return value;
  try {
    return JSON.stringify(value) ?? String(value);
  } catch {
    return String(value);
  }
}

/** Results cross a structured-clone boundary; JSON round-trip drops Office proxies' internals (they define toJSON). */
function toJsonSafe(value: unknown): unknown {
  if (value === undefined) return undefined;
  try {
    return JSON.parse(JSON.stringify(value));
  } catch {
    return String(value);
  }
}
