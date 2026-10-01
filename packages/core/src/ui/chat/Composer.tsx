import {
  ArrowUp,
  FileText,
  Paperclip,
  ShieldCheck,
  ShieldOff,
  Square,
  Undo2,
  X,
} from "lucide-react";
import { useRef, useState } from "react";
import type { ChatSession, ChatSessionState, ModelInfo } from "../../contracts.ts";
import { IconButton, Select } from "../components/controls.tsx";
import { clampThinkingLevel, formatTokens, thinkingLevelLabels } from "./transcript.ts";

export function Composer({
  session,
  state,
  models,
  onFiles,
  onUndo,
  onSendError,
}: {
  session: ChatSession;
  state: ChatSessionState;
  models: ModelInfo[];
  onFiles: (files: File[]) => void;
  onUndo: () => void;
  onSendError: (cause: unknown) => void;
}) {
  const [text, setText] = useState("");
  const fileInputRef = useRef<HTMLInputElement>(null);
  const currentModel = models.find((model) => model.id === state.modelId);
  const thinkingLevels = currentModel?.thinkingLevels ?? [state.thinkingLevel];
  const canSend = text.trim() !== "" || (!state.isStreaming && state.stagedAttachments.length > 0);

  function submit() {
    if (!canSend) return;
    const message = text.trim();
    if (state.isStreaming) session.steer(message);
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

  const fullAccess = state.permissionMode === "full";

  return (
    <div className="flex flex-col gap-1.5 border-t border-neutral-200 p-2 dark:border-neutral-800">
      {state.stagedAttachments.length > 0 && (
        <div className="flex flex-wrap gap-1.5">
          {state.stagedAttachments.map((attachment) => (
            <span
              key={attachment.id}
              title={
                attachment.summary ? `${attachment.name} · ${attachment.summary}` : attachment.name
              }
              className="inline-flex max-w-full items-center gap-1 rounded border border-neutral-200 py-0.5 pr-0.5 pl-1.5 text-[11px] dark:border-neutral-700"
            >
              <FileText size={12} className="shrink-0 text-neutral-500" />
              <span className="truncate">{attachment.name}</span>
              {attachment.summary && (
                <span className="shrink-0 text-neutral-500">{attachment.summary}</span>
              )}
              <button
                type="button"
                onClick={() => session.unstageAttachment(attachment.id)}
                aria-label={`Remove ${attachment.name}`}
                className="rounded p-0.5 text-neutral-500 hover:bg-neutral-100 dark:hover:bg-neutral-800"
              >
                <X size={12} />
              </button>
            </span>
          ))}
        </div>
      )}
      <div className="flex flex-col rounded-lg border border-neutral-300 bg-white focus-within:border-accent dark:border-neutral-700 dark:bg-neutral-900">
        <textarea
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
          placeholder={state.isStreaming ? "Add to the current task" : "Ask Footnote"}
          aria-label="Message"
          className="max-h-48 min-h-12 resize-none bg-transparent px-2.5 pt-2 pb-1 outline-none [field-sizing:content] placeholder:text-neutral-400"
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
            aria-label="Model"
            value={state.modelId ?? ""}
            onChange={(event) => changeModel(event.target.value)}
            disabled={models.length === 0}
            className="max-w-40"
          >
            {!currentModel && (
              <option value={state.modelId ?? ""}>{state.modelId ?? "No model"}</option>
            )}
            {models.map((model) => (
              <option key={model.id} value={model.id}>
                {model.id}
              </option>
            ))}
          </Select>
          <Select
            aria-label="Thinking effort"
            value={state.thinkingLevel}
            onChange={(event) => {
              const level = thinkingLevels.find((candidate) => candidate === event.target.value);
              if (level) session.setThinkingLevel(level);
            }}
            className="shrink-0"
          >
            {thinkingLevels.map((level) => (
              <option key={level} value={level}>
                {thinkingLevelLabels[level]}
              </option>
            ))}
          </Select>
          <IconButton
            icon={fullAccess ? ShieldOff : ShieldCheck}
            label={fullAccess ? "Full access: edits run without asking" : "Ask before edits"}
            aria-pressed={fullAccess}
            onClick={() => session.setPermissionMode(fullAccess ? "ask" : "full")}
            className={fullAccess ? "text-amber-600 dark:text-amber-400" : ""}
          />
          {state.canUndo && !state.isStreaming && (
            <IconButton icon={Undo2} label="Undo last turn" onClick={onUndo} />
          )}
          <span className="flex-1" />
          {state.contextUsage && <ContextUsage {...state.contextUsage} />}
          {state.isStreaming && (
            <IconButton icon={Square} label="Stop" onClick={() => session.abort()} />
          )}
          {(!state.isStreaming || canSend) && (
            <button
              type="button"
              onClick={submit}
              disabled={!canSend}
              aria-label={state.isStreaming ? "Send to the running task" : "Send"}
              title={state.isStreaming ? "Send to the running task" : "Send"}
              className="inline-flex size-7 shrink-0 items-center justify-center rounded-md bg-accent text-white hover:bg-accent-hover disabled:opacity-40"
            >
              <ArrowUp size={16} />
            </button>
          )}
        </div>
      </div>
    </div>
  );
}

const ringRadius = 6;
const ringCircumference = 2 * Math.PI * ringRadius;

function ContextUsage({ tokens, window }: { tokens: number; window: number }) {
  const fraction = Math.min(tokens / window, 1);
  const label = `Context ${formatTokens(tokens)} of ${formatTokens(window)} tokens`;
  const tone = fraction > 0.85 ? "text-amber-600 dark:text-amber-400" : "text-neutral-500";
  return (
    <span
      role="img"
      aria-label={label}
      title={label}
      className={`inline-flex size-7 items-center justify-center ${tone}`}
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
  );
}
