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

/** `syncs` counts the context.sync() calls the realm made while the op ran (approximate when ops overlap). */
export type OpResponse = ({ ok: true; value: unknown } | { ok: false; error: OfficeErrorInfo }) & {
  syncs?: number;
};

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

let syncCount = 0;
let countingSyncs = false;

/** Counts every RequestContext.sync in this realm, for the opt-in timing log (see logOpTiming). */
function countSyncs(): void {
  if (countingSyncs) return;
  // Timing is optional: a realm without the class (an old or stubbed Office.js) must still run ops.
  const prototype = (globalThis as { PowerPoint?: typeof PowerPoint }).PowerPoint?.RequestContext
    ?.prototype;
  if (!prototype) return;
  countingSyncs = true;
  const sync = prototype.sync;
  prototype.sync = function <T>(this: PowerPoint.RequestContext, passThroughValue?: T) {
    syncCount += 1;
    return sync.call<PowerPoint.RequestContext, [T | undefined], Promise<T>>(
      this,
      passThroughValue,
    );
  };
}

export async function callOp(ops: OpRegistry, op: string, args: unknown): Promise<OpResponse> {
  const run = ops[op];
  if (!run) return { ok: false, error: { message: `Unknown op "${op}"` } };
  countSyncs();
  const before = syncCount;
  try {
    return { ok: true, value: await run(args), syncs: syncCount - before };
  } catch (error) {
    return { ok: false, error: toErrorInfo(error), syncs: syncCount - before };
  }
}

/** Source of an async function expression that compiles to a `CodeBody`. */
/** Names the model's code in stack traces, so errors can point at its line. */
export const CODE_SOURCE_URL = "footnote-code.js";
const sourceUrlComment = `\n//# sourceURL=${CODE_SOURCE_URL}`;

export function codeBodySource(code: string): string {
  return `async (context, footnote, console) => {\n${code}\n}`;
}

/** Where the model's first line lands in stack traces for each way of compiling the code. */
export const codeLineOffsets = {
  /** `${prefix}${codeBodySource(code)}…` evaluated as a script whose first line holds the arrow's opening. */
  script: 1,
  /** `new Function(\`return ${codeBodySource(code)}\`)`: V8 adds the `function anonymous(` header lines. */
  newFunction: 3,
};

export function withSourceUrl(source: string): string {
  return source + sourceUrlComment;
}

export interface CodeSource {
  code: string;
  lineOffset: number;
}

export async function runCodeBody(body: CodeBody, source?: CodeSource): Promise<CodeRunResult> {
  const logs: string[] = [];
  const log = (...values: unknown[]) => {
    logs.push(values.map(formatLogValue).join(" "));
  };
  const console = { log, info: log, warn: log, error: log };
  let before: string[] | undefined;
  let run: CodeRunResult;
  try {
    const result = await PowerPoint.run(async (context) => {
      before = await slideIdsIn(context);
      return body(context, createFootnoteHelpers(context), console);
    });
    run = { ok: true, result: toJsonSafe(result), logs };
  } catch (error) {
    run = { ok: false, logs, error: withCodeLocation(toErrorInfo(error), error, source) };
  }
  // A fresh batch: after a failed sync the run's context is unusable.
  const after = await PowerPoint.run(slideIdsIn).catch(() => undefined);
  return before && after ? { ...run, slides: { before, after } } : run;
}

async function slideIdsIn(context: PowerPoint.RequestContext): Promise<string[]> {
  const slides = context.presentation.slides.load("items/id");
  await context.sync();
  return slides.items.map((slide) => slide.id);
}

/** "Cannot set properties of undefined" is useless without the line; quote the model's line that threw. */
function withCodeLocation(
  info: OfficeErrorInfo,
  error: unknown,
  source: CodeSource | undefined,
): OfficeErrorInfo {
  if (!source || !(error instanceof Error) || !error.stack) return info;
  const match = new RegExp(`${CODE_SOURCE_URL.replace(".", "\\.")}:(\\d+):\\d+`).exec(error.stack);
  if (!match) return info;
  const lineNumber = Number(match[1]) - source.lineOffset;
  const line = source.code.split("\n")[lineNumber - 1]?.trim();
  if (!line) return info;
  const quoted = line.length > 160 ? `${line.slice(0, 160)}…` : line;
  return { ...info, message: `${info.message} (line ${lineNumber}: ${quoted})` };
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
