import type { AgentToolResult } from "@earendil-works/pi-agent-core";
import type { WriteReceipt } from "../ops/types.ts";

/** Compact JSON for the model; the same value goes to `details` for the UI. */
export function jsonResult<T>(value: T): AgentToolResult<T> {
  return { content: [{ type: "text", text: JSON.stringify(value) }], details: value };
}

/** What the model sees of a write: changed IDs, read-back values and warnings. Fingerprints stay internal. */
export function receiptResult(
  receipt: WriteReceipt,
  extra: Record<string, unknown> = {},
): AgentToolResult<WriteReceipt> {
  const { changed, verified, warnings } = receipt;
  const value = { changed, verified, ...(warnings.length > 0 && { warnings }), ...extra };
  return { content: [{ type: "text", text: JSON.stringify(value) }], details: receipt };
}

export const plural = (count: number, noun: string): string =>
  `${count} ${noun}${count === 1 ? "" : "s"}`;
