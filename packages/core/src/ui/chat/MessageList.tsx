import type { AssistantMessage, ToolResultMessage, UserMessage } from "@earendil-works/pi-ai";
import { ChevronRight, FileText, LoaderCircle, TriangleAlert } from "lucide-react";
import { useLayoutEffect, useMemo, useRef } from "react";
import type { AgentMessage, ChatSession, ChatSessionState, FootnoteTool } from "../../contracts.ts";
import { ImageThumbnail } from "../components/ImageThumbnail.tsx";
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
}: {
  state: ChatSessionState;
  session: ChatSession;
  tools: ToolIndex;
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
        return <UserMessageView key={key} message={message} />;
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

  const waiting =
    state.isStreaming && !state.streamingMessage && state.pendingApprovals.length === 0;

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
        {waiting && (
          <LoaderCircle size={16} className="animate-spin text-neutral-400" aria-label="Working" />
        )}
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

function UserMessageView({ message }: { message: UserMessage }) {
  const { text, attachmentNames, images } = splitUserContent(message.content);
  return (
    <div className="ml-8 flex flex-col items-end gap-1.5 self-end">
      {(attachmentNames.length > 0 || images.length > 0) && (
        <div className="flex flex-wrap items-center justify-end gap-1.5">
          {attachmentNames.map((name, index) => (
            <span
              key={index}
              className="inline-flex max-w-48 items-center gap-1 rounded border border-neutral-200 px-1.5 py-0.5 text-[11px] text-neutral-600 dark:border-neutral-700 dark:text-neutral-400"
            >
              <FileText size={12} className="shrink-0" />
              <span className="truncate">{name}</span>
            </span>
          ))}
          {images.map((image, index) => (
            <ImageThumbnail key={index} image={image} alt={`Attached image ${index + 1}`} />
          ))}
        </div>
      )}
      {text && (
        <div className="rounded-lg bg-neutral-100 px-2.5 py-1.5 whitespace-pre-wrap dark:bg-neutral-800">
          {text}
        </div>
      )}
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
  return (
    <div className="flex flex-col gap-2">
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
        <span className="text-[12px] text-neutral-500">Stopped</span>
      )}
    </div>
  );
}

function CompactionDivider({ summary }: { summary: string }) {
  return (
    <details className="group text-[12px] text-neutral-500 dark:text-neutral-400">
      <summary className="flex cursor-pointer list-none items-center gap-2 select-none hover:text-neutral-700 dark:hover:text-neutral-200 [&::-webkit-details-marker]:hidden">
        <span className="h-px flex-1 bg-neutral-200 dark:bg-neutral-700" />
        Earlier conversation summarized
        <ChevronRight size={13} className="transition-transform group-open:rotate-90" />
        <span className="h-px flex-1 bg-neutral-200 dark:bg-neutral-700" />
      </summary>
      <div className="mt-2 rounded-md bg-neutral-50 p-2 text-neutral-700 dark:bg-neutral-800/60 dark:text-neutral-300">
        <Markdown text={summary} />
      </div>
    </details>
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
    <details className="group text-[12px] text-neutral-500 dark:text-neutral-400">
      <summary className="flex cursor-pointer list-none items-center gap-1 select-none hover:text-neutral-700 dark:hover:text-neutral-200 [&::-webkit-details-marker]:hidden">
        <ChevronRight size={13} className="transition-transform group-open:rotate-90" />
        {active ? "Thinking…" : "Thought"}
      </summary>
      <div className="mt-1 border-l-2 border-neutral-200 pl-2.5 whitespace-pre-wrap dark:border-neutral-700">
        {redacted ? "Hidden by the provider." : text}
      </div>
    </details>
  );
}

function ErrorNote({ text }: { text: string }) {
  return (
    <div
      role="alert"
      className="flex items-start gap-2 rounded-md bg-red-50 p-2 text-[12px] text-red-800 dark:bg-red-950/60 dark:text-red-300"
    >
      <TriangleAlert size={14} className="mt-0.5 shrink-0" />
      <span className="min-w-0 break-words">{text}</span>
    </div>
  );
}
