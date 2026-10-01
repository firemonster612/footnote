import { Trash } from "lucide-react";
import { useEffect, useState } from "react";
import type { ChatSummary, FootnoteApp } from "../contracts.ts";
import { Button } from "./components/controls.tsx";
import { errorMessage } from "./hooks.ts";

const relativeTime = new Intl.RelativeTimeFormat(undefined, { numeric: "auto" });
const dayMs = 86_400_000;

export function HistoryView({
  app,
  documentId,
  currentChatId,
  onOpen,
  onDelete,
}: {
  app: FootnoteApp;
  documentId: string;
  currentChatId?: string;
  onOpen: (chatId: string) => void;
  onDelete: (chatId: string) => Promise<void>;
}) {
  const [chats, setChats] = useState<ChatSummary[]>();
  const [error, setError] = useState<string>();
  const [confirmingId, setConfirmingId] = useState<string>();

  useEffect(() => {
    let active = true;
    app.chats.list(documentId).then(
      (listed) => active && setChats(listed.toSorted((a, b) => b.updatedAt - a.updatedAt)),
      (cause) => active && setError(errorMessage(cause)),
    );
    return () => {
      active = false;
    };
  }, [app, documentId]);

  async function remove(chatId: string) {
    try {
      await onDelete(chatId);
      setChats((current) => current?.filter((chat) => chat.id !== chatId));
    } catch (cause) {
      setError(errorMessage(cause));
    }
  }

  return (
    <div className="flex min-h-0 flex-1 flex-col">
      {error && <p className="px-3 pt-2 text-[12px] text-red-700 dark:text-red-400">{error}</p>}
      {chats?.length === 0 && (
        <p className="px-3 pt-2 text-neutral-500">No chats for this document yet.</p>
      )}
      <ul className="min-h-0 flex-1 overflow-y-auto px-1.5 py-2">
        {chats?.map((chat) => (
          <li
            key={chat.id}
            className="group flex items-center gap-1 rounded-md hover:bg-neutral-100 dark:hover:bg-neutral-800"
          >
            <button
              type="button"
              onClick={() => onOpen(chat.id)}
              aria-current={chat.id === currentChatId}
              className="flex min-w-0 flex-1 flex-col items-start px-1.5 py-1.5 text-left aria-[current=true]:font-semibold"
            >
              <span className="w-full truncate">{chat.title || "Untitled chat"}</span>
              <span className="text-[11px] font-normal text-neutral-500">
                {formatAge(chat.updatedAt)}
              </span>
            </button>
            {confirmingId === chat.id ? (
              <Button
                variant="danger"
                className="mr-1"
                autoFocus
                onBlur={() => setConfirmingId(undefined)}
                onClick={() => remove(chat.id)}
              >
                Delete
              </Button>
            ) : (
              <button
                type="button"
                aria-label={`Delete ${chat.title || "chat"}`}
                onClick={() => setConfirmingId(chat.id)}
                className="mr-1 rounded p-1.5 text-neutral-500 opacity-0 group-hover:opacity-100 hover:bg-neutral-200 hover:text-red-700 focus-visible:opacity-100 dark:hover:bg-neutral-700"
              >
                <Trash size={14} />
              </button>
            )}
          </li>
        ))}
      </ul>
    </div>
  );
}

function formatAge(timestamp: number): string {
  const days = Math.round((timestamp - Date.now()) / dayMs);
  if (days === 0)
    return new Date(timestamp).toLocaleTimeString(undefined, {
      hour: "numeric",
      minute: "2-digit",
    });
  return relativeTime.format(days, "day");
}
