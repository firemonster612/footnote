import type { ImageContent, Message, TextContent } from "@earendil-works/pi-ai";
import type { AgentMessage } from "../contracts.ts";

/** Host state (e.g. `<deck_state>`) captured before a request. Sent to the model as user text. */
export interface DocumentContextMessage {
  role: "documentContext";
  text: string;
  timestamp: number;
}

/** Replaces compacted history. Sent to the model as user text. */
export interface CompactionSummaryMessage {
  role: "compactionSummary";
  summary: string;
  timestamp: number;
}

declare module "@earendil-works/pi-agent-core" {
  interface CustomAgentMessages {
    documentContext: DocumentContextMessage;
    compactionSummary: CompactionSummaryMessage;
  }
}

export function documentContextMessage(text: string): DocumentContextMessage {
  return { role: "documentContext", text, timestamp: Date.now() };
}

/** How many image-bearing messages from earlier turns keep their images; the current turn always keeps its own. */
export const recentImageMessages = 3;

const imagePlaceholder: TextContent = {
  type: "text",
  text: "[Earlier image removed to save context. Render or read it again if you need to see it.]",
};

/**
 * Converts the transcript to provider messages: custom messages become user text. Images from the current turn
 * (since the last user message) always survive, so a model that renders several slides sees all of them; images
 * from earlier turns survive only in the newest `keepImageMessages` messages that carry any. 0 drops every image,
 * for models without image input.
 */
export function toLlmMessages(messages: AgentMessage[], keepImageMessages: number): Message[] {
  const currentTurnStart = messages.findLastIndex((message) => message.role === "user");
  let earlierImageMessagesSeen = 0;
  const converted: Message[] = [];
  for (let index = messages.length - 1; index >= 0; index -= 1) {
    const message = messages[index]!;
    const llmMessage = withSendableImages(toLlmMessage(message));
    const carriesImages =
      (llmMessage.role === "user" || llmMessage.role === "toolResult") &&
      hasImages(llmMessage.content);
    if (carriesImages) {
      const inCurrentTurn = index >= currentTurnStart;
      if (!inCurrentTurn) earlierImageMessagesSeen += 1;
      const keep =
        keepImageMessages > 0 && (inCurrentTurn || earlierImageMessagesSeen <= keepImageMessages);
      if (!keep) {
        converted.push({
          ...llmMessage,
          content: withoutImages(llmMessage.content as (TextContent | ImageContent)[]),
        });
        continue;
      }
    }
    converted.push(llmMessage);
  }
  return converted.reverse();
}

function toLlmMessage(message: AgentMessage): Message {
  switch (message.role) {
    case "documentContext":
      return {
        role: "user",
        content: [{ type: "text", text: message.text }],
        timestamp: message.timestamp,
      };
    case "compactionSummary":
      return {
        role: "user",
        content: [
          {
            type: "text",
            text: `The earlier part of this conversation was compacted. Summary:\n<conversation_summary>\n${message.summary}\n</conversation_summary>`,
          },
        ],
        timestamp: message.timestamp,
      };
    default:
      return message;
  }
}

type MessageContent = string | (TextContent | ImageContent)[];

function hasImages(content: MessageContent): content is (TextContent | ImageContent)[] {
  return typeof content !== "string" && content.some((block) => block.type === "image");
}

/** Images with missing or non-string data (a failed render) would make the provider reject the whole request. */
function isSendableImage(block: ImageContent): boolean {
  return typeof block.data === "string" && block.data.length > 0;
}

const brokenImagePlaceholder: TextContent = {
  type: "text",
  text: "[Image unavailable: it failed to load when it was captured.]",
};

function withSendableImages<T extends Message>(message: T): T {
  if (typeof message.content === "string") return message;
  const content = message.content as (TextContent | ImageContent | { type: string })[];
  if (!content.some((block) => block.type === "image" && !isSendableImage(block as ImageContent)))
    return message;
  return {
    ...message,
    content: content.map((block) =>
      block.type === "image" && !isSendableImage(block as ImageContent)
        ? brokenImagePlaceholder
        : block,
    ),
  };
}

function withoutImages(content: (TextContent | ImageContent)[]): TextContent[] {
  return content.map((block) => (block.type === "image" ? imagePlaceholder : block));
}
