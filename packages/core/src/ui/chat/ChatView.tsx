import { X } from "lucide-react";
import { useRef, useState, type DragEvent } from "react";
import type { ChatSession, ModelInfo, UndoReport } from "../../contracts.ts";
import { IconButton } from "../components/icon-button.tsx";
import { errorMessage, useSessionState } from "../hooks.ts";
import { cn } from "../lib/utils.ts";
import { Composer } from "./Composer.tsx";
import { MessageList, type ToolIndex } from "./MessageList.tsx";

interface Notice {
  tone: "info" | "error";
  text: string;
}

export function ChatView({
  session,
  models,
  tools,
}: {
  session: ChatSession;
  models: ModelInfo[];
  tools: ToolIndex;
}) {
  const state = useSessionState(session);
  const [notice, setNotice] = useState<Notice>();
  const [dragging, setDragging] = useState(false);
  const [draft, setDraft] = useState("");
  const textareaRef = useRef<HTMLTextAreaElement>(null);
  const showError = (cause: unknown) => setNotice({ tone: "error", text: errorMessage(cause) });

  async function stageFiles(files: File[]) {
    for (const file of files) {
      try {
        await session.stageAttachment(file);
      } catch (cause) {
        setNotice({ tone: "error", text: `${file.name}: ${errorMessage(cause)}` });
      }
    }
  }

  async function undo() {
    try {
      setNotice({ tone: "info", text: formatUndoReport(await session.undoLastTurn()) });
    } catch (cause) {
      showError(cause);
    }
  }

  async function revert(messageTimestamp: number) {
    try {
      const { text, undo } = await session.revertTo(messageTimestamp);
      setDraft(text);
      setNotice({ tone: "info", text: formatRevertReport(undo) });
      textareaRef.current?.focus();
    } catch (cause) {
      showError(cause);
    }
  }

  const hasFiles = (event: DragEvent) => event.dataTransfer.types.includes("Files");

  return (
    <div
      className="relative flex min-h-0 flex-1 flex-col"
      onDragOver={(event) => {
        if (!hasFiles(event)) return;
        event.preventDefault();
        setDragging(true);
      }}
      onDragLeave={(event) => {
        const target = event.relatedTarget;
        if (!(target instanceof Node && event.currentTarget.contains(target))) setDragging(false);
      }}
      onDrop={(event) => {
        if (!hasFiles(event)) return;
        event.preventDefault();
        setDragging(false);
        void stageFiles([...event.dataTransfer.files]);
      }}
    >
      <MessageList
        state={state}
        session={session}
        tools={tools}
        onRevert={(timestamp) => void revert(timestamp)}
      />
      {notice && (
        <div
          role={notice.tone === "error" ? "alert" : "status"}
          className={cn(
            "mx-2 flex items-start gap-2 rounded-md py-1 pr-1 pl-2.5 text-small whitespace-pre-line",
            notice.tone === "error"
              ? "bg-danger-subtle text-danger"
              : "bg-muted text-muted-foreground",
          )}
        >
          <span className="min-w-0 flex-1 py-1">{notice.text}</span>
          <IconButton
            icon={X}
            label="Dismiss"
            onClick={() => setNotice(undefined)}
            className="size-6 text-current [&_svg]:size-3.5"
          />
        </div>
      )}
      <Composer
        session={session}
        state={state}
        models={models}
        onFiles={(files) => void stageFiles(files)}
        onUndo={() => void undo()}
        onSendError={showError}
        text={draft}
        onTextChange={setDraft}
        textareaRef={textareaRef}
      />
      {dragging && (
        <div className="pointer-events-none absolute inset-2 flex items-center justify-center rounded-lg border-2 border-dashed border-accent bg-background/90 font-medium text-accent-text">
          Drop files to attach
        </div>
      )}
    </div>
  );
}

function formatUndoReport({ restored, removed, warnings }: UndoReport): string {
  return [`Undid the last turn (restored ${restored}, removed ${removed}).`, ...warnings].join(
    "\n",
  );
}

function formatRevertReport({ restored, removed, warnings }: UndoReport): string {
  const changes =
    restored + removed > 0
      ? ` Slides: ${restored} restored, ${removed} removed.`
      : " No slide changes needed undoing.";
  return [
    `Reverted to that request; it's back in the box below.${changes}`,
    ...warnings.filter((warning) => warning !== "Nothing to undo."),
  ].join("\n");
}
