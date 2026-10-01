import type { Message } from "@earendil-works/pi-ai";
import { estimateContextTokens, estimateMessageTokens } from "@earendil-works/pi-ai/utils/estimate";
import type { AgentMessage } from "../contracts.ts";
import { toLlmMessages, type CompactionSummaryMessage } from "../context/messages.ts";

/** Room kept free below the window besides the output reserve. */
const headroomTokens = 13_000;
const maxOutputReserveTokens = 20_000;

export function compactionThresholdTokens(contextWindow: number, maxOutputTokens: number): number {
  return contextWindow - Math.min(maxOutputTokens, maxOutputReserveTokens) - headroomTokens;
}

/** Context tokens of a provider request: the last reported usage plus an estimate of what came after it. */
export function estimateRequestTokens(llmMessages: Message[]): number {
  return estimateContextTokens(llmMessages).tokens;
}

/** Recent history kept verbatim: 20k tokens, or a quarter of small windows. */
export function keepRecentTokensFor(contextWindow: number): number {
  return Math.min(20_000, Math.floor(contextWindow / 4));
}

/**
 * Index of the first message to keep verbatim, so that at least `keepRecentTokens` of recent history survives. The cut
 * lands before a user or assistant message, never between a tool call and its result, and leaves at least one
 * non-system message to summarize. Undefined when there's nothing to compact.
 */
export function findCutIndex(
  messages: AgentMessage[],
  keepRecentTokens: number,
): number | undefined {
  const firstConversationIndex = messages.findIndex((message) => message.role !== "system");
  if (firstConversationIndex === -1) return undefined;

  let keptTokens = 0;
  for (let index = messages.length - 1; index > firstConversationIndex; index -= 1) {
    const message = messages[index];
    if (!message) continue;
    keptTokens += toLlmMessages([message], 0).reduce(
      (sum, llmMessage) => sum + estimateMessageTokens(llmMessage),
      0,
    );
    if (keptTokens >= keepRecentTokens && isCutPoint(message)) return index;
  }
  return undefined;
}

function isCutPoint(message: AgentMessage): boolean {
  return message.role === "user" || message.role === "assistant";
}

export const summaryInstructions = `You summarize the earlier part of a conversation between a user and Footnote, an assistant that edits an Office document through tools. The summary replaces that history, so the assistant must be able to continue the work from it alone.

Write these sections in order, as terse bullet points:
1. Intent: what the user is trying to achieve overall.
2. User requests: every request the user made, in order and close to their wording, each marked done, partly done, or not started.
3. Decisions: choices agreed with the user (content, structure, style), and anything they rejected or denied.
4. Document changes: every slide, shape, or other object created, changed, or deleted, with exact IDs and what changed.
5. Verification: what was checked (renders, read-backs) and what is unverified or known to be wrong.
6. Outstanding work: what remains, next step first.
7. Working context: loaded skills, attachments in use (with IDs), and tool facts later steps depend on.

Keep IDs, numbers, and names exact. Leave out pleasantries and anything superseded. Don't invent details.`;

const summaryExcerptChars = 1_500;

function excerpt(text: string): string {
  return text.length <= summaryExcerptChars
    ? text
    : `${text.slice(0, summaryExcerptChars)} […${text.length - summaryExcerptChars} more chars]`;
}

/** Renders history as plain text for the summarizer: no images, tool output excerpted. */
export function transcriptForSummary(messages: AgentMessage[]): string {
  return messages
    .map((message) => {
      switch (message.role) {
        case "system":
          return "";
        case "user": {
          if (typeof message.content === "string") return `User: ${message.content}`;
          const text = message.content
            .flatMap((block) => (block.type === "text" ? [block.text] : []))
            .join("\n");
          const images = message.content.filter((block) => block.type === "image").length;
          return `User: ${text}${images > 0 ? ` [${images} image(s)]` : ""}`;
        }
        case "assistant":
          return message.content
            .flatMap((block) => {
              if (block.type === "text") return [`Assistant: ${block.text}`];
              if (block.type === "toolCall")
                return [
                  `Assistant called ${block.name}(${excerpt(JSON.stringify(block.arguments))})`,
                ];
              return [];
            })
            .join("\n");
        case "toolResult": {
          const text = message.content
            .flatMap((block) => (block.type === "text" ? [block.text] : []))
            .join("\n");
          return `${message.toolName} ${message.isError ? "failed" : "returned"}: ${excerpt(text)}`;
        }
        case "documentContext":
          return `Document state at that point: ${excerpt(message.text)}`;
        case "compactionSummary":
          return `Summary of the conversation before this point:\n${message.summary}`;
      }
    })
    .filter((line) => line.length > 0)
    .join("\n\n");
}

/** Writes a summary of the given transcript text. */
export type Summarize = (transcript: string, signal?: AbortSignal) => Promise<string>;

/**
 * Replaces everything before the cut with a summary, keeping system messages (prompt and tool declarations) and the
 * recent tail verbatim. Undefined when there's nothing to compact.
 */
export async function compactMessages(
  messages: AgentMessage[],
  keepRecentTokens: number,
  summarize: Summarize,
  signal?: AbortSignal,
): Promise<AgentMessage[] | undefined> {
  const cut = findCutIndex(messages, keepRecentTokens);
  if (cut === undefined) return undefined;
  const head = messages.slice(0, cut);
  const summary: CompactionSummaryMessage = {
    role: "compactionSummary",
    summary: await summarize(transcriptForSummary(head), signal),
    timestamp: Date.now(),
  };
  return [...head.filter((message) => message.role === "system"), summary, ...messages.slice(cut)];
}

/** Trips after `limit` compactions within `windowTurns` turns, so an oversized tail can't loop forever. */
export function createCompactionBreaker(limit = 3, windowTurns = 3) {
  const compactionTurns: number[] = [];
  return {
    tripped: (turn: number) =>
      compactionTurns.filter((compactedAt) => turn - compactedAt < windowTurns).length >= limit,
    record: (turn: number) => void compactionTurns.push(turn),
  };
}
