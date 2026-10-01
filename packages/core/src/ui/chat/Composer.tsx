import {
  ArrowUp,
  CornerDownRight,
  FileText,
  Paperclip,
  ShieldCheck,
  ShieldOff,
  Square,
  Undo2,
  X,
  type LucideIcon,
} from "lucide-react";
import { useRef, type RefObject } from "react";
import type {
  ChatSession,
  ChatSessionState,
  ModelInfo,
  PermissionMode,
  QueuedMessage,
} from "../../contracts.ts";
import { Badge } from "../components/badge.tsx";
import { Button } from "../components/button.tsx";
import { IconButton } from "../components/icon-button.tsx";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "../components/select.tsx";
import { ToggleGroup, ToggleGroupItem } from "../components/toggle-group.tsx";
import { Tooltip, TooltipContent, TooltipTrigger } from "../components/tooltip.tsx";
import { cn } from "../lib/utils.ts";
import { clampThinkingLevel, formatTokens, thinkingLevelLabels } from "./transcript.ts";

export function Composer({
  session,
  state,
  models,
  onFiles,
  onUndo,
  onSendError,
  text,
  onTextChange,
  textareaRef,
}: {
  session: ChatSession;
  state: ChatSessionState;
  models: ModelInfo[];
  onFiles: (files: File[]) => void;
  onUndo: () => void;
  onSendError: (cause: unknown) => void;
  /** The draft lives in ChatView so a revert can put the request back here. */
  text: string;
  onTextChange: (text: string) => void;
  textareaRef: RefObject<HTMLTextAreaElement | null>;
}) {
  const setText = onTextChange;
  const fileInputRef = useRef<HTMLInputElement>(null);
  const currentModel = models.find((model) => model.id === state.modelId);
  const thinkingLevels = currentModel?.thinkingLevels ?? [state.thinkingLevel];
  const canSend = text.trim() !== "" || state.stagedAttachments.length > 0;

  function submit() {
    if (!canSend) return;
    const message = text.trim();
    // While a run is going, Enter queues; a queued message's Steer button injects it now.
    if (state.isStreaming) session.queue(message);
    else session.send(message).catch(onSendError);
    setText("");
  }

  function changeModel(modelId: string) {
    session.setModel(modelId);
    const levels = models.find((model) => model.id === modelId)?.thinkingLevels ?? [];
    if (levels.includes(state.thinkingLevel)) return;
    const level = clampThinkingLevel(state.thinkingLevel, levels);
    if (level) session.setThinkingLevel(level);
  }

  return (
    <div className="flex flex-col gap-1.5 border-t p-2">
      {state.stagedAttachments.length > 0 && (
        <div className="flex flex-wrap gap-1">
          {state.stagedAttachments.map((attachment) => (
            <Badge
              key={attachment.id}
              title={
                attachment.summary ? `${attachment.name} · ${attachment.summary}` : attachment.name
              }
              className="pr-0.5 text-foreground"
            >
              <FileText size={12} className="text-subtle-foreground" />
              <span className="truncate">{attachment.name}</span>
              {attachment.summary && (
                <span className="shrink-0 text-subtle-foreground">{attachment.summary}</span>
              )}
              <IconButton
                icon={X}
                label={`Remove ${attachment.name}`}
                onClick={() => session.unstageAttachment(attachment.id)}
                className="size-5 rounded-sm [&_svg]:size-3"
              />
            </Badge>
          ))}
        </div>
      )}
      <div className="flex flex-col">
        {state.queuedMessages.length > 0 && (
          <QueueBar messages={state.queuedMessages} running={state.isStreaming} session={session} />
        )}
        <div className="flex flex-col rounded-lg border border-input bg-background shadow-control transition-[border-color,box-shadow] duration-100 hover:border-input-hover has-[textarea:focus]:border-ring has-[textarea:focus]:ring-2 has-[textarea:focus]:ring-ring/20">
          <textarea
            ref={textareaRef}
            value={text}
            onChange={(event) => setText(event.target.value)}
            onKeyDown={(event) => {
              if (event.key !== "Enter" || event.shiftKey || event.nativeEvent.isComposing) return;
              event.preventDefault();
              submit();
            }}
            onPaste={(event) => {
              const files = [...event.clipboardData.files];
              if (files.length === 0) return;
              event.preventDefault();
              onFiles(files);
            }}
            rows={2}
            placeholder={state.isStreaming ? "Queue a message" : "Ask Footnote"}
            aria-label="Message"
            className="max-h-48 min-h-12 resize-none bg-transparent px-2.5 pt-2 pb-1 outline-none [field-sizing:content] placeholder:text-subtle-foreground focus-visible:outline-none"
          />
          <div className="flex items-center gap-0.5 px-1 pb-1">
            <IconButton
              icon={Paperclip}
              label="Attach files"
              onClick={() => fileInputRef.current?.click()}
            />
            <input
              ref={fileInputRef}
              type="file"
              multiple
              hidden
              onChange={(event) => {
                onFiles([...(event.target.files ?? [])]);
                event.target.value = "";
              }}
            />
            <Select
              value={state.modelId ?? ""}
              onValueChange={changeModel}
              disabled={models.length === 0}
            >
              <SelectTrigger variant="ghost" aria-label="Model" className="max-w-40">
                <SelectValue placeholder="No model" />
              </SelectTrigger>
              <SelectContent>
                {state.modelId && !currentModel && (
                  <SelectItem value={state.modelId}>{state.modelId}</SelectItem>
                )}
                {models.map((model) => (
                  <SelectItem key={model.id} value={model.id}>
                    {model.id}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
            <Select
              value={state.thinkingLevel}
              onValueChange={(value) => {
                const level = thinkingLevels.find((candidate) => candidate === value);
                if (level) session.setThinkingLevel(level);
              }}
            >
              <SelectTrigger variant="ghost" aria-label="Thinking effort" className="shrink-0">
                <SelectValue />
              </SelectTrigger>
              <SelectContent>
                {thinkingLevels.map((level) => (
                  <SelectItem key={level} value={level}>
                    {thinkingLevelLabels[level]}
                  </SelectItem>
                ))}
              </SelectContent>
            </Select>
            <ToggleGroup
              type="single"
              aria-label="Permissions"
              value={state.permissionMode}
              onValueChange={(mode) => {
                if (mode === "ask" || mode === "full") session.setPermissionMode(mode);
              }}
              className="ml-0.5"
            >
              {permissionModes.map(({ mode, icon: Icon, label, className }) => (
                <Tooltip key={mode}>
                  <TooltipTrigger asChild>
                    <ToggleGroupItem value={mode} aria-label={label} className={className}>
                      <Icon size={14} />
                    </ToggleGroupItem>
                  </TooltipTrigger>
                  <TooltipContent>{label}</TooltipContent>
                </Tooltip>
              ))}
            </ToggleGroup>
            {state.canUndo && !state.isStreaming && (
              <IconButton icon={Undo2} label="Undo last turn" onClick={onUndo} />
            )}
            <span className="flex-1" />
            {state.contextUsage && <ContextUsage {...state.contextUsage} />}
            {state.isStreaming && (
              <IconButton icon={Square} label="Stop" onClick={() => session.abort()} />
            )}
            {(!state.isStreaming || canSend) && (
              <IconButton
                icon={ArrowUp}
                label={state.isStreaming ? "Queue message" : "Send"}
                variant="primary"
                onClick={submit}
                disabled={!canSend}
              />
            )}
          </div>
        </div>
      </div>
    </div>
  );
}

/** Queued messages, drawn as a tab attached to the composer's top edge. */
function QueueBar({
  messages,
  running,
  session,
}: {
  messages: QueuedMessage[];
  running: boolean;
  session: ChatSession;
}) {
  return (
    <ol
      aria-label="Queued messages"
      className="mx-2 flex max-h-32 flex-col overflow-y-auto rounded-t-md border border-b-0 border-input bg-muted/60 py-0.5"
    >
      {messages.map((message, index) => (
        <li key={message.id} className="flex items-center gap-1.5 py-0.5 pr-0.5 pl-2 text-small">
          <CornerDownRight size={12} aria-hidden className="shrink-0 text-subtle-foreground" />
          <span className="min-w-0 flex-1 truncate" title={message.text}>
            {message.text || "(attachments only)"}
          </span>
          {message.attachments.length > 0 && (
            <span className="shrink-0 text-caption text-subtle-foreground">
              +{message.attachments.length} file{message.attachments.length === 1 ? "" : "s"}
            </span>
          )}
          <Tooltip>
            <TooltipTrigger asChild>
              <Button
                variant="ghost"
                className="h-6 shrink-0 px-2"
                onClick={() => session.steerQueued(message.id)}
              >
                {running ? "Steer" : "Send"}
              </Button>
            </TooltipTrigger>
            <TooltipContent>
              {running
                ? "Send now: Footnote reads it after its current step"
                : index === 0
                  ? "Send now"
                  : "Send now, ahead of the others"}
            </TooltipContent>
          </Tooltip>
          <IconButton
            icon={X}
            label="Remove from queue"
            onClick={() => session.removeQueued(message.id)}
            className="size-6 [&_svg]:size-3"
          />
        </li>
      ))}
    </ol>
  );
}

const permissionModes: {
  mode: PermissionMode;
  icon: LucideIcon;
  label: string;
  className?: string;
}[] = [
  { mode: "ask", icon: ShieldCheck, label: "Ask before edits" },
  {
    mode: "full",
    icon: ShieldOff,
    label: "Full access: edits run without asking",
    className: "aria-checked:text-warning",
  },
];

const ringRadius = 6;
const ringCircumference = 2 * Math.PI * ringRadius;

function ContextUsage({ tokens, window }: { tokens: number; window: number }) {
  const fraction = Math.min(tokens / window, 1);
  const label = `Context ${formatTokens(tokens)} of ${formatTokens(window)} tokens`;
  return (
    <Tooltip>
      <TooltipTrigger asChild>
        <span
          role="img"
          aria-label={label}
          className={cn(
            "inline-flex size-7 items-center justify-center text-subtle-foreground",
            fraction > 0.85 && "text-warning",
          )}
        >
          <svg width="16" height="16" viewBox="0 0 16 16" className="-rotate-90">
            <circle
              cx="8"
              cy="8"
              r={ringRadius}
              fill="none"
              stroke="currentColor"
              strokeOpacity="0.25"
              strokeWidth="2"
            />
            <circle
              cx="8"
              cy="8"
              r={ringRadius}
              fill="none"
              stroke="currentColor"
              strokeWidth="2"
              strokeDasharray={ringCircumference}
              strokeDashoffset={ringCircumference * (1 - fraction)}
            />
          </svg>
        </span>
      </TooltipTrigger>
      <TooltipContent>{label}</TooltipContent>
    </Tooltip>
  );
}
