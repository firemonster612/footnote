import { Trash } from "lucide-react";
import { useEffect, useState } from "react";
import type { ChatSummary, FootnoteApp } from "../contracts.ts";
import { Button } from "./components/button.tsx";
import { HighlightGroup, highlightItem } from "./components/highlight-group.tsx";
import { IconButton } from "./components/icon-button.tsx";
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
      {error && <p className="px-3 pt-2 text-small text-danger">{error}</p>}
      {chats?.length === 0 && (
        <p className="px-3 pt-2 text-muted-foreground">No chats for this document yet.</p>
      )}
      <div className="min-h-0 flex-1 overflow-y-auto px-1.5 py-1.5">
        <HighlightGroup>
          <ul className="flex flex-col gap-px">
            {chats?.map((chat) => (
              <li
                key={chat.id}
                {...highlightItem}
                className="group flex items-center gap-1 rounded-sm pr-1"
              >
                <button
                  type="button"
                  onClick={() => onOpen(chat.id)}
                  aria-current={chat.id === currentChatId}
                  className="group/open flex min-w-0 flex-1 cursor-default flex-col items-start rounded-sm px-2 py-1.5 text-left focus-visible:outline-offset-[-2px]"
                >
                  <span className="w-full truncate group-aria-[current=true]/open:font-semibold group-aria-[current=true]/open:text-accent-text">
                    {chat.title || "Untitled chat"}
                  </span>
                  <span className="text-caption text-subtle-foreground">
                    {formatAge(chat.updatedAt)}
                  </span>
                </button>
                {confirmingId === chat.id ? (
                  <Button
                    variant="danger"
                    autoFocus
                    onBlur={() => setConfirmingId(undefined)}
                    onClick={() => remove(chat.id)}
                  >
                    Delete
                  </Button>
                ) : (
                  <IconButton
                    icon={Trash}
                    label={`Delete ${chat.title || "chat"}`}
                    onClick={() => setConfirmingId(chat.id)}
                    className="opacity-0 group-hover:opacity-100 hover:bg-pressed hover:text-danger focus-visible:opacity-100"
                  />
                )}
              </li>
            ))}
          </ul>
        </HighlightGroup>
      </div>
    </div>
  );
}

/** Today's chats show their time; older ones count calendar days ("yesterday" means the previous date). */
export function formatAge(timestamp: number): string {
  const days = Math.round((startOfDay(new Date(timestamp)) - startOfDay(new Date())) / dayMs);
  if (days === 0)
    return new Date(timestamp).toLocaleTimeString(undefined, {
      hour: "numeric",
      minute: "2-digit",
    });
  return relativeTime.format(days, "day");
}

// Rounding the difference absorbs the 23- and 25-hour days around DST changes.
function startOfDay(date: Date): number {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate()).getTime();
}
