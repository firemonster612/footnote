// Agent-side helpers shared by the shells' OfficeHost implementations.

import {
  OfficeOpError,
  type CodeRunResult,
  type OfficeCallOptions,
  type OfficeHostKind,
  type OfficeHostStatus,
} from "@footnote/core/contracts";
import type { OfficeInfo, OpResponse } from "./realm.ts";

const DEFAULT_TIMEOUT_MS = 60_000;

/** Office.HostType names. */
export const hostNames: Record<OfficeHostKind, string> = {
  powerpoint: "PowerPoint",
  word: "Word",
  excel: "Excel",
};

function hostKind(hostName: string | undefined): OfficeHostKind | undefined {
  return (["powerpoint", "word", "excel"] as const).find((kind) => hostNames[kind] === hostName);
}

export function statusFromInfo(
  info: OfficeInfo,
  expected: OfficeHostKind,
  fallbackDocumentId?: string,
): OfficeHostStatus {
  const host = hostKind(info.hostName);
  if (host !== expected) {
    const reason = info.hostName
      ? `Footnote works with ${hostNames[expected]}, but this add-in is open in ${info.hostName}.`
      : `Not running inside ${hostNames[expected]}.`;
    return { connected: false, reason, ...(host && { host }) };
  }
  const documentId = info.documentUrl || fallbackDocumentId;
  const documentName = info.documentUrl && fileName(info.documentUrl);
  return {
    connected: true,
    host,
    apiVersions: info.apiVersions,
    ...(documentId && { documentId }),
    ...(documentName && { documentName }),
  };
}

function fileName(url: string): string | undefined {
  const name = url.split(/[?#]/, 1)[0]?.split(/[\\/]/).pop();
  if (!name) return undefined;
  try {
    return decodeURIComponent(name);
  } catch {
    return name;
  }
}

/**
 * A request reached the Office realm but its answer never came back (stopped, or the bridge dropped), so its edits
 * may or may not have applied. Matched by `name` across packages, like the TimeoutError of a missed deadline.
 */
export class OutcomeUnknownError extends Error {
  constructor(message: string, options?: ErrorOptions) {
    super(message, options);
    this.name = "OutcomeUnknownError";
  }
}

/**
 * Waits for work already sent to the Office realm. Rejects with a TimeoutError after `timeoutMs` (default 60s), or
 * with an OutcomeUnknownError when the signal aborts.
 */
export function withDeadline<T>(
  work: Promise<T>,
  label: string,
  { signal, timeoutMs = DEFAULT_TIMEOUT_MS }: OfficeCallOptions = {},
): Promise<T> {
  return new Promise<T>((resolve, reject) => {
    const settle = (finish: () => void) => {
      clearTimeout(timer);
      signal?.removeEventListener("abort", onAbort);
      finish();
    };
    const onAbort = () =>
      settle(() =>
        reject(
          new OutcomeUnknownError(`${label} was stopped before it answered`, {
            cause: signal?.reason,
          }),
        ),
      );
    const timer = setTimeout(
      () =>
        settle(() =>
          reject(new DOMException(`${label} timed out after ${timeoutMs / 1000}s`, "TimeoutError")),
        ),
      timeoutMs,
    );
    if (signal?.aborted) return onAbort();
    signal?.addEventListener("abort", onAbort, { once: true });
    work.then(
      (value) => settle(() => resolve(value)),
      (error: unknown) => settle(() => reject(error)),
    );
  });
}

const DEBUG_KEY = "footnote:debug";

/**
 * Opt-in timing log for measuring on real PowerPoint: run `localStorage["footnote:debug"] = "1"` in the console of
 * the page that owns the OfficeHost, or any page of the same origin (the extension's side panel sets it for its engine).
 * Logs each call's name and duration to that page's console and, for ops, the sync count the realm reported.
 */
export function logOpTiming(name: string, startedAt: number, response?: OpResponse): void {
  if (localStorage.getItem(DEBUG_KEY) !== "1") return;
  const ms = Math.round(performance.now() - startedAt);
  const syncs = response?.syncs === undefined ? "" : `, ${response.syncs} syncs`;
  const failed = response && !response.ok ? ", failed" : "";
  console.debug(`[footnote] ${name}: ${ms} ms${syncs}${failed}`);
}

export function unwrapOpResponse<T>(op: string, response: OpResponse): T {
  if (!response.ok) throw new OfficeOpError(op, response.error);
  // OfficeHost.call<T> lets the caller name an op's result type; op results aren't validated at runtime.
  return response.value as T;
}

/** Result for a code run that timed out, was cancelled, or lost its bridge mid-flight. */
export function unknownOutcome(error: unknown): CodeRunResult {
  const message = error instanceof Error ? error.message : String(error);
  return {
    ok: false,
    logs: [],
    outcomeUnknown: true,
    error: {
      message: `${message}. Its edits may or may not have applied; re-read before retrying.`,
    },
  };
}
