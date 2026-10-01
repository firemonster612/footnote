import { X } from "lucide-react";
import { useState, type DragEvent } from "react";
import type { ChatSession, ModelInfo, UndoReport } from "../../contracts.ts";
import { errorMessage, useSessionState } from "../hooks.ts";
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
      <MessageList state={state} session={session} tools={tools} />
      {notice && (
        <div
          role={notice.tone === "error" ? "alert" : "status"}
          className={`mx-2 flex items-start gap-2 rounded-md px-2 py-1.5 text-[12px] whitespace-pre-line ${
            notice.tone === "error"
              ? "bg-red-50 text-red-800 dark:bg-red-950/60 dark:text-red-300"
              : "bg-neutral-100 text-neutral-700 dark:bg-neutral-800 dark:text-neutral-300"
          }`}
        >
          <span className="min-w-0 flex-1">{notice.text}</span>
          <button
            type="button"
            aria-label="Dismiss"
            onClick={() => setNotice(undefined)}
            className="shrink-0 opacity-70 hover:opacity-100"
          >
            <X size={14} />
          </button>
        </div>
      )}
      <Composer
        session={session}
        state={state}
        models={models}
        onFiles={(files) => void stageFiles(files)}
        onUndo={() => void undo()}
        onSendError={showError}
      />
      {dragging && (
        <div className="pointer-events-none absolute inset-2 flex items-center justify-center rounded-lg border-2 border-dashed border-accent bg-white/85 text-accent-fg dark:bg-neutral-900/85">
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
