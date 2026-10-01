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

/** How many image-bearing messages keep their images in what the model sees. */
export const recentImageMessages = 3;

const imagePlaceholder: TextContent = {
  type: "text",
  text: "[Image omitted: only the newest images are kept in context, and none for models without image input.]",
};

/**
 * Converts the transcript to provider messages: custom messages become user text, and images survive only in the
 * newest `keepImageMessages` user/tool-result messages that carry any.
 */
export function toLlmMessages(messages: AgentMessage[], keepImageMessages: number): Message[] {
  let imageMessagesSeen = 0;
  const converted: Message[] = [];
  for (const message of messages.toReversed()) {
    const llmMessage = toLlmMessage(message);
    if (
      (llmMessage.role === "user" || llmMessage.role === "toolResult") &&
      hasImages(llmMessage.content)
    ) {
      imageMessagesSeen += 1;
      if (imageMessagesSeen > keepImageMessages) {
        converted.push({ ...llmMessage, content: withoutImages(llmMessage.content) });
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

function withoutImages(content: (TextContent | ImageContent)[]): TextContent[] {
  return content.map((block) => (block.type === "image" ? imagePlaceholder : block));
}
