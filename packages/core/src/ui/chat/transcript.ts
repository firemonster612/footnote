// Pure helpers that turn session state into what the message list shows.

import type { ImageContent, TextContent, ToolResultMessage } from "@earendil-works/pi-ai";
import type { AgentMessage, ApprovalRequest, ThinkingLevel } from "../../contracts.ts";

export type ToolCallStatus = "awaiting-approval" | "running" | "done" | "error" | "not-run";

export function indexToolResults(messages: AgentMessage[]): Map<string, ToolResultMessage> {
  const results = new Map<string, ToolResultMessage>();
  for (const message of messages) {
    if (message.role === "toolResult") results.set(message.toolCallId, message);
  }
  return results;
}

export function toolCallStatus(
  toolCallId: string,
  result: ToolResultMessage | undefined,
  pendingApprovals: ApprovalRequest[],
  isStreaming: boolean,
): ToolCallStatus {
  if (result) return result.isError ? "error" : "done";
  if (pendingApprovals.some((approval) => approval.toolCallId === toolCallId))
    return "awaiting-approval";
  return isStreaming ? "running" : "not-run";
}

export interface UserMessageParts {
  text: string;
  attachmentNames: string[];
  images: ImageContent[];
}

// Attachment excerpts (attachmentToContent) start with an <attachment name="..."> tag holding a JSON string.
const attachmentTagPattern = /^<attachment\b[^>]*?\bname=("(?:[^"\\]|\\.)*")/;
// Per-turn context blocks are for the model only.
const contextBlockPattern = /^<(deck_state|document_state|context)\b/;

export function splitUserContent(
  content: string | (TextContent | ImageContent)[],
): UserMessageParts {
  if (typeof content === "string") return { text: content, attachmentNames: [], images: [] };
  const texts: string[] = [];
  const attachmentNames: string[] = [];
  const images: ImageContent[] = [];
  for (const part of content) {
    if (part.type === "image") {
      images.push(part);
      continue;
    }
    const trimmed = part.text.trimStart();
    const attachment = attachmentTagPattern.exec(trimmed);
    if (attachment) attachmentNames.push(JSON.parse(attachment[1]!));
    else if (!contextBlockPattern.test(trimmed)) texts.push(part.text);
  }
  return { text: texts.join("\n\n"), attachmentNames, images };
}

export const thinkingLevelOrder: ThinkingLevel[] = [
  "off",
  "minimal",
  "low",
  "medium",
  "high",
  "xhigh",
  "max",
];

/** The supported level closest to `level`, preferring the lower one on a tie. */
export function clampThinkingLevel(
  level: ThinkingLevel,
  supported: ThinkingLevel[],
): ThinkingLevel | undefined {
  const target = thinkingLevelOrder.indexOf(level);
  let best: ThinkingLevel | undefined;
  let bestDistance = Infinity;
  for (const candidate of supported) {
    const distance = Math.abs(thinkingLevelOrder.indexOf(candidate) - target);
    if (
      distance < bestDistance ||
      (distance === bestDistance && thinkingLevelOrder.indexOf(candidate) < target)
    ) {
      best = candidate;
      bestDistance = distance;
    }
  }
  return best;
}

export const thinkingLevelLabels: Record<ThinkingLevel, string> = {
  off: "Off",
  minimal: "Minimal",
  low: "Low",
  medium: "Medium",
  high: "High",
  xhigh: "X-High",
  max: "Max",
};

export function formatTokens(tokens: number): string {
  if (tokens < 1000) return String(tokens);
  if (tokens < 1_000_000) return `${Math.round(tokens / 1000)}k`;
  return `${Number((tokens / 1_000_000).toFixed(1))}M`;
}
