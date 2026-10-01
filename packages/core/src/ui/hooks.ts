import { useCallback, useEffect, useRef, useState } from "react";
import type {
  ChatSession,
  ChatSessionState,
  FootnoteApp,
  ModelInfo,
  OfficeHostStatus,
  Settings,
} from "../contracts.ts";

export function useSessionState(session: ChatSession): ChatSessionState {
  const [state, setState] = useState(() => session.getState());
  useEffect(() => {
    setState(session.getState());
    return session.subscribe(setState);
  }, [session]);
  return state;
}

export function useSettings(app: FootnoteApp): Settings {
  const [settings, setSettings] = useState(() => app.settings.get());
  useEffect(() => app.settings.subscribe(setSettings), [app]);
  return settings;
}

export function useHostStatus(app: FootnoteApp): OfficeHostStatus | undefined {
  const [status, setStatus] = useState<OfficeHostStatus>();
  useEffect(() => {
    let active = true;
    void app.host.status().then((initial) => active && setStatus(initial));
    const unsubscribe = app.host.onStatusChange(setStatus);
    return () => {
      active = false;
      unsubscribe();
    };
  }, [app]);
  return status;
}

export interface ModelList {
  models: ModelInfo[];
  error?: string;
  loading: boolean;
  /** Refetches the list; failures land in `error`. */
  reload(): Promise<void>;
}

const RELOAD_DEBOUNCE_MS = 500;

/** Lists the endpoint's models, refetching when the endpoint changes. */
export function useModels(app: FootnoteApp, settings: Settings): ModelList {
  const [models, setModels] = useState<ModelInfo[]>([]);
  const [error, setError] = useState<string>();
  const [loading, setLoading] = useState(false);
  const latestRequest = useRef(0);

  const reload = useCallback(async () => {
    const request = ++latestRequest.current;
    setLoading(true);
    const outcome = await app.models.list().then(
      (listed) => ({ listed, error: undefined }),
      (cause: unknown) => ({ listed: [], error: errorMessage(cause) }),
    );
    if (request !== latestRequest.current) return; // a newer request superseded this one
    setModels(outcome.listed);
    setError(outcome.error);
    setLoading(false);
  }, [app]);

  const { baseUrl, apiKey } = settings.endpoint;
  useEffect(() => {
    // Without a key the proxy just answers 401; keyless servers can still use Test connection.
    if (!baseUrl || !apiKey) return;
    // Settings save per keystroke; fetch once typing pauses instead of once per character.
    const timer = setTimeout(() => void reload(), RELOAD_DEBOUNCE_MS);
    return () => clearTimeout(timer);
  }, [reload, baseUrl, apiKey]);

  return { models, error, loading, reload };
}

export function errorMessage(cause: unknown): string {
  return cause instanceof Error ? cause.message : String(cause);
}
