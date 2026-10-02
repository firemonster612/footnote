// Pure helpers that turn session state into what the message list shows.

import type {
  AssistantMessage,
  ImageContent,
  TextContent,
  ToolResultMessage,
} from "@earendil-works/pi-ai";
import type {
  AgentMessage,
  ApprovalRequest,
  ChatSessionState,
  ThinkingLevel,
} from "../../contracts.ts";

export type ToolCallStatus =
  | "awaiting-approval"
  | "running"
  | "queued"
  | "done"
  | "error"
  | "not-run";

export function indexToolResults(messages: AgentMessage[]): Map<string, ToolResultMessage> {
  const results = new Map<string, ToolResultMessage>();
  for (const message of messages) {
    if (message.role === "toolResult") results.set(message.toolCallId, message);
  }
  return results;
}

/**
 * The assistant message whose tool calls can still run: the newest one of the running run. Calls without
 * results in any other message were cut off by a stop and will never run.
 */
export function liveAssistantMessage({
  messages,
  streamingMessage,
  isStreaming,
}: Pick<ChatSessionState, "messages" | "streamingMessage" | "isStreaming">):
  | AssistantMessage
  | undefined {
  if (!isStreaming) return undefined;
  if (streamingMessage?.role === "assistant") return streamingMessage;
  const lastUser = messages.findLastIndex((message) => message.role === "user");
  const lastAssistant = messages.findLastIndex((message) => message.role === "assistant");
  const message = messages[lastAssistant];
  return lastAssistant > lastUser && message?.role === "assistant" ? message : undefined;
}

/**
 * Tool calls run one at a time in order, so only the first call without a result is running (or waiting for
 * approval); the ones after it are queued.
 */
export function toolCallStatus(
  toolCallId: string,
  result: ToolResultMessage | undefined,
  pendingApprovals: ApprovalRequest[],
  inLiveMessage: boolean,
  firstUnfinishedCallId?: string,
): ToolCallStatus {
  if (result) return result.isError ? "error" : "done";
  if (pendingApprovals.some((approval) => approval.toolCallId === toolCallId))
    return "awaiting-approval";
  if (!inLiveMessage) return "not-run";
  return firstUnfinishedCallId === undefined || firstUnfinishedCallId === toolCallId
    ? "running"
    : "queued";
}

export interface UserMessageParts {
  text: string;
  attachmentNames: string[];
  images: ImageContent[];
}

// Attachment excerpts (attachmentToContent) start with an <attachment name="..."> tag holding a JSON string.
const attachmentTagPattern = /^<attachment\b[^>]*?\bname=("(?:[^"\\]|\\.)*")/;

/** Requests are `[typed text, ...attachment blocks]`, so only blocks after the first can be attachments. */
export function splitUserContent(
  content: string | (TextContent | ImageContent)[],
): UserMessageParts {
  if (typeof content === "string") return { text: content, attachmentNames: [], images: [] };
  const texts: string[] = [];
  const attachmentNames: string[] = [];
  const images: ImageContent[] = [];
  content.forEach((part, index) => {
    if (part.type === "image") return void images.push(part);
    const name = index === 0 ? undefined : attachmentName(part.text);
    if (name === undefined) texts.push(part.text);
    else attachmentNames.push(name);
  });
  return { text: texts.join("\n\n"), attachmentNames, images };
}

/** The name in an attachment excerpt's tag, or undefined when the text isn't a well-formed excerpt. */
function attachmentName(text: string): string | undefined {
  const quoted = attachmentTagPattern.exec(text.trimStart())?.[1];
  if (quoted === undefined) return undefined;
  try {
    return JSON.parse(quoted); // the pattern matched a quoted string, so this parses to a string or throws
  } catch {
    return undefined; // typed text that merely looks like a tag, e.g. a Windows path with backslashes
  }
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
