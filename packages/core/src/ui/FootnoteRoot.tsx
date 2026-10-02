import {
  ArrowLeft,
  Clock,
  MessageSquarePlus,
  Settings as SettingsIcon,
  Unplug,
} from "lucide-react";
import { useCallback, useEffect, useMemo, useRef, useState } from "react";
import type { AttachmentStore, ChatSession, FootnoteApp, OfficeHostStatus } from "../contracts.ts";
import type { ToolIndex } from "./chat/MessageList.tsx";
import { ChatView } from "./chat/ChatView.tsx";
import { Button } from "./components/button.tsx";
import { IconButton } from "./components/icon-button.tsx";
import { TooltipProvider } from "./components/tooltip.tsx";
import { HistoryView } from "./HistoryView.tsx";
import { errorMessage, useHostStatus, useModels, useSettings } from "./hooks.ts";
import { EndpointSection } from "./settings/EndpointSection.tsx";
import { SettingsView } from "./settings/SettingsView.tsx";

type View = "setup" | "chat" | "history" | "settings";

const viewTitles: Partial<Record<View, string>> = { history: "Chats", settings: "Settings" };

export function FootnoteRoot({ app }: { app: FootnoteApp }) {
  const settings = useSettings(app);
  const modelList = useModels(app, settings);
  const status = useHostStatus(app);
  const lastDocument = useLastDocument(status);
  const documentId = lastDocument?.id;
  const tools = useToolIndex(app);
  const [view, setView] = useState<View>(() => (settings.endpoint.baseUrl ? "chat" : "setup"));
  const [session, setSession] = useState<ChatSession>();
  const [chatError, setChatError] = useState<string>();

  const latestLoad = useRef(0);

  /** Shows the loaded chat unless a later load started meanwhile (open A, then B: a slow A mustn't win). */
  const loadSession = useCallback(async (load: () => Promise<ChatSession>) => {
    const request = ++latestLoad.current;
    try {
      const loaded = await load();
      if (request !== latestLoad.current) return;
      setSession(loaded);
      setChatError(undefined);
    } catch (cause) {
      if (request === latestLoad.current) setChatError(errorMessage(cause));
    }
  }, []);

  useEffect(() => {
    if (documentId) void loadSession(() => openLatestChat(app, documentId));
  }, [app, documentId, loadSession]);

  function newChat() {
    if (!documentId) return;
    setView("chat");
    const currentIsEmpty = session?.getState().messages.length === 0;
    if (!currentIsEmpty) void loadSession(() => app.chats.create(documentId));
  }

  async function deleteChat(chatId: string) {
    await app.chats.delete(chatId);
    if (chatId === session?.getState().id) setSession(undefined);
  }

  function leaveHistory() {
    if (!session && documentId) void loadSession(() => openLatestChat(app, documentId));
    setView("chat");
  }

  return (
    <TooltipProvider>
      <div className="flex h-full min-h-0 flex-col">
        <header className="flex h-10 shrink-0 items-center gap-0.5 border-b px-2">
          {view === "settings" || view === "history" ? (
            <IconButton
              icon={ArrowLeft}
              label="Back to chat"
              onClick={view === "history" ? leaveHistory : () => setView("chat")}
            />
          ) : (
            <LogoMark />
          )}
          <span className="min-w-0 flex-1 truncate px-1.5 font-semibold">
            {viewTitles[view] ?? lastDocument?.name ?? "Footnote"}
          </span>
          {view !== "setup" && (
            <>
              <IconButton
                icon={MessageSquarePlus}
                label="New chat"
                onClick={newChat}
                disabled={!documentId}
              />
              <IconButton
                icon={Clock}
                label="Chats"
                onClick={() => setView("history")}
                disabled={!documentId}
              />
              <IconButton
                icon={SettingsIcon}
                label="Settings"
                onClick={() => setView("settings")}
              />
            </>
          )}
        </header>
        {status && !status.connected && <ConnectionBanner status={status} />}
        {view === "setup" && (
          <div className="flex flex-col gap-3 overflow-y-auto p-3">
            <EndpointSection app={app} settings={settings} modelList={modelList} />
            <Button
              variant="primary"
              className="self-start"
              disabled={modelList.models.length === 0}
              onClick={() => setView("chat")}
            >
              Continue
            </Button>
          </div>
        )}
        {view === "settings" && (
          <SettingsView app={app} settings={settings} modelList={modelList} />
        )}
        {view === "history" && documentId && (
          <HistoryView
            app={app}
            documentId={documentId}
            currentChatId={session?.getState().id}
            onOpen={(chatId) =>
              void loadSession(() => app.chats.open(chatId)).then(() => setView("chat"))
            }
            onDelete={deleteChat}
          />
        )}
        {view === "chat" &&
          (session ? (
            <ChatView
              key={session.getState().id}
              session={session}
              models={modelList.models}
              tools={tools}
            />
          ) : (
            <div className="flex flex-1 items-center justify-center p-6 text-center text-muted-foreground">
              {chatError}
            </div>
          ))}
      </div>
    </TooltipProvider>
  );
}

interface DocumentRef {
  id: string;
  name?: string;
}

/** The last connected document, kept while the host is briefly disconnected so the open chat stays put. */
function useLastDocument(status: OfficeHostStatus | undefined): DocumentRef | undefined {
  const [document, setDocument] = useState<DocumentRef>();
  const id = status?.documentId;
  const name = status?.documentName;
  useEffect(() => {
    if (id) setDocument({ id, name });
  }, [id, name]);
  return document;
}

const noAttachments: AttachmentStore = { list: () => [], get: () => undefined };

/** Host tools by name, for tool-card labels. Built from the host module because the session doesn't expose its tools. */
function useToolIndex(app: FootnoteApp): ToolIndex {
  return useMemo(() => {
    const env = { host: app.host, attachments: noAttachments, settings: () => app.settings.get() };
    return new Map(app.hostModule.createTools(env).map((tool) => [tool.name, tool]));
  }, [app]);
}

async function openLatestChat(app: FootnoteApp, documentId: string): Promise<ChatSession> {
  const chats = await app.chats.list(documentId);
  const latest = chats.toSorted((a, b) => b.updatedAt - a.updatedAt)[0];
  return latest ? app.chats.open(latest.id) : app.chats.create(documentId);
}

function ConnectionBanner({ status }: { status: OfficeHostStatus }) {
  return (
    <div
      role="status"
      className="flex items-center gap-2 border-b border-warning-border bg-warning-subtle px-3 py-1.5 text-small text-warning"
    >
      <Unplug size={14} className="shrink-0" />
      <span className="min-w-0">{status.reason ?? "Not connected to an Office document"}</span>
    </div>
  );
}

function LogoMark() {
  return (
    <svg viewBox="0 0 128 128" className="size-6 shrink-0" aria-hidden="true">
      <rect width="128" height="128" rx="28" fill="#0F766E" />
      <g stroke="#fff" strokeWidth="15" strokeLinecap="round">
        <line x1="64" y1="53" x2="64" y2="25" />
        <line x1="64" y1="53" x2="88.2" y2="39" />
        <line x1="64" y1="53" x2="88.2" y2="67" />
        <line x1="64" y1="53" x2="64" y2="81" />
        <line x1="64" y1="53" x2="39.8" y2="67" />
        <line x1="64" y1="53" x2="39.8" y2="39" />
      </g>
      <rect x="32" y="99" width="64" height="11" rx="5.5" fill="#fff" />
    </svg>
  );
}
