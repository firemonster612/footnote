import type { AssistantMessage, ToolResultMessage, UserMessage } from "@earendil-works/pi-ai";
import { ChevronRight, FileText, LoaderCircle, TriangleAlert } from "lucide-react";
import { useLayoutEffect, useMemo, useRef } from "react";
import type { AgentMessage, ChatSession, ChatSessionState, FootnoteTool } from "../../contracts.ts";
import { Badge } from "../components/badge.tsx";
import {
  Collapsible,
  CollapsibleContent,
  CollapsibleTrigger,
  disclosureChevron,
} from "../components/collapsible.tsx";
import { ImageThumbnail } from "../components/ImageThumbnail.tsx";
import { MessageActions } from "./MessageActions.tsx";
import { Markdown } from "../markdown/Markdown.tsx";
import { ApprovalCard } from "./ApprovalCard.tsx";
import { ToolCard } from "./ToolCard.tsx";
import { indexToolResults, splitUserContent, toolCallStatus } from "./transcript.ts";

export type ToolIndex = ReadonlyMap<string, FootnoteTool>;

/** Distance from the bottom within which new content keeps the list pinned to the end. */
const stickToBottomPx = 80;

export function MessageList({
  state,
  session,
  tools,
  onRevert,
}: {
  state: ChatSessionState;
  session: ChatSession;
  tools: ToolIndex;
  onRevert: (messageTimestamp: number) => void;
}) {
  const results = useMemo(() => indexToolResults(state.messages), [state.messages]);
  const scrollRef = useRef<HTMLDivElement>(null);
  const pinnedRef = useRef(true);

  useLayoutEffect(() => {
    const element = scrollRef.current;
    if (element && pinnedRef.current) element.scrollTop = element.scrollHeight;
  }, [state]);

  const renderMessage = (message: AgentMessage, key: string | number, streaming = false) => {
    switch (message.role) {
      case "user":
        return (
          <UserMessageView
            key={key}
            message={message}
            {...(state.revertibleRequests.includes(message.timestamp) && {
              onRevert: () => onRevert(message.timestamp),
            })}
            {...(state.isStreaming && { revertBlockedReason: "Stop the response to revert" })}
          />
        );
      case "assistant":
        return (
          <AssistantMessageView
            key={key}
            message={message}
            streaming={streaming}
            state={state}
            results={results}
            tools={tools}
          />
        );
      case "compactionSummary":
        return <CompactionDivider key={key} summary={message.summary} />;
      default:
        // Tool results render inside their tool cards; document context and unknown roles are model-only.
        return null;
    }
  };

  // Stays up for the whole run, including while text streams, so the only sign of work isn't the Stop button.
  const working = state.isStreaming && state.pendingApprovals.length === 0;

  return (
    <div
      ref={scrollRef}
      onScroll={(event) => {
        const element = event.currentTarget;
        pinnedRef.current =
          element.scrollHeight - element.scrollTop - element.clientHeight < stickToBottomPx;
      }}
      className="min-h-0 flex-1 overflow-y-auto"
    >
      <div role="log" aria-live="polite" className="flex flex-col gap-3 p-3">
        {state.messages.map((message, index) => renderMessage(message, index))}
        {state.streamingMessage && renderMessage(state.streamingMessage, "streaming", true)}
        {working && <WorkingIndicator />}
        {state.pendingApprovals.map((request) => (
          <ApprovalCard
            key={request.id}
            request={request}
            onDecide={(decision) => session.resolveApproval(request.id, decision)}
          />
        ))}
        {state.error && <ErrorNote text={state.error} />}
      </div>
    </div>
  );
}

function WorkingIndicator() {
  return (
    <div role="status" className="flex items-center gap-2 text-small">
      <LoaderCircle size={14} aria-hidden className="animate-spin text-subtle-foreground" />
      <span className="text-shimmer font-medium">Working…</span>
    </div>
  );
}

function UserMessageView({
  message,
  onRevert,
  revertBlockedReason,
}: {
  message: UserMessage;
  onRevert?: () => void;
  revertBlockedReason?: string;
}) {
  const { text, attachmentNames, images } = splitUserContent(message.content);
  return (
    <div className="group ml-8 flex flex-col items-end gap-1.5 self-end">
      {(attachmentNames.length > 0 || images.length > 0) && (
        <div className="flex flex-wrap items-center justify-end gap-1.5">
          {attachmentNames.map((name, index) => (
            <Badge key={index} className="max-w-48">
              <FileText size={12} />
              <span className="truncate">{name}</span>
            </Badge>
          ))}
          {images.map((image, index) => (
            <ImageThumbnail key={index} image={image} alt={`Attached image ${index + 1}`} />
          ))}
        </div>
      )}
      {text && (
        <div className="rounded-lg rounded-br-sm bg-muted px-2.5 py-1.5 whitespace-pre-wrap">
          {text}
        </div>
      )}
      <MessageActions
        text={text}
        align="end"
        {...(onRevert && { onRevert })}
        {...(revertBlockedReason && { revertBlockedReason })}
      />
    </div>
  );
}

function AssistantMessageView({
  message,
  streaming,
  state,
  results,
  tools,
}: {
  message: AssistantMessage;
  streaming: boolean;
  state: ChatSessionState;
  results: Map<string, ToolResultMessage>;
  tools: ToolIndex;
}) {
  const replyText = message.content
    .flatMap((part) => (part.type === "text" && part.text.trim() ? [part.text] : []))
    .join("\n\n");
  return (
    <div className="group flex flex-col gap-2">
      {message.content.map((part, index) => {
        switch (part.type) {
          case "text":
            return part.text.trim() ? <Markdown key={index} text={part.text} /> : null;
          case "thinking":
            return (
              <Thinking
                key={index}
                text={part.thinking}
                redacted={part.redacted === true}
                active={streaming && index === message.content.length - 1}
              />
            );
          case "toolCall": {
            const result = results.get(part.id);
            return (
              <ToolCard
                key={part.id || index}
                call={part}
                tool={tools.get(part.name)}
                result={result}
                status={toolCallStatus(part.id, result, state.pendingApprovals, state.isStreaming)}
              />
            );
          }
          default:
            return part satisfies never;
        }
      })}
      {message.stopReason === "error" && (
        <ErrorNote text={message.errorMessage ?? "The model request failed."} />
      )}
      {message.stopReason === "aborted" && (
        <span className="text-small text-muted-foreground">Stopped</span>
      )}
      {!streaming && replyText && <MessageActions text={replyText} align="start" />}
    </div>
  );
}

function CompactionDivider({ summary }: { summary: string }) {
  return (
    <Collapsible className="text-small text-muted-foreground">
      <CollapsibleTrigger className="group flex w-full cursor-default items-center gap-2 rounded-sm select-none hover:text-foreground">
        <span className="h-px flex-1 bg-border" />
        Earlier conversation summarized
        <ChevronRight size={13} className={disclosureChevron} />
        <span className="h-px flex-1 bg-border" />
      </CollapsibleTrigger>
      <CollapsibleContent>
        <div className="mt-2 rounded-md bg-muted p-2 text-foreground">
          <Markdown text={summary} />
        </div>
      </CollapsibleContent>
    </Collapsible>
  );
}

function Thinking({
  text,
  redacted,
  active,
}: {
  text: string;
  redacted: boolean;
  active: boolean;
}) {
  return (
    <Collapsible className="text-small text-muted-foreground">
      <CollapsibleTrigger className="group flex cursor-default items-center gap-1 rounded-sm select-none hover:text-foreground">
        <ChevronRight size={13} className={disclosureChevron} />
        {active ? "Thinking…" : "Thought"}
      </CollapsibleTrigger>
      <CollapsibleContent>
        <div className="mt-1 ml-1.5 border-l-2 pl-2.5 whitespace-pre-wrap">
          {redacted ? "Hidden by the provider." : text}
        </div>
      </CollapsibleContent>
    </Collapsible>
  );
}

function ErrorNote({ text }: { text: string }) {
  return (
    <div
      role="alert"
      className="flex items-start gap-2 rounded-md bg-danger-subtle p-2 text-small text-danger"
    >
      <TriangleAlert size={14} className="mt-0.5 shrink-0" />
      <span className="min-w-0 break-words">{text}</span>
    </div>
  );
}
