import type { FootnoteApp, KeyValueStore, Settings } from "../contracts.ts";

export const defaultSettings: Settings = {
  endpoint: { baseUrl: "", apiKey: "" },
  defaultPermissionMode: "ask",
  thinkingLevel: "medium",
  apiOverrides: {},
  customSkills: [],
};

const settingsKey = "settings";

export type SettingsHandle = FootnoteApp["settings"];

/** Loads persisted settings over the defaults; every update is persisted before listeners run. */
export async function loadSettings(store: KeyValueStore): Promise<SettingsHandle> {
  let current: Settings = {
    ...defaultSettings,
    ...(await store.get<Partial<Settings>>(settingsKey)),
  };
  const listeners = new Set<(settings: Settings) => void>();

  return {
    get: () => current,
    async update(patch) {
      current = { ...current, ...patch };
      await store.set(settingsKey, current);
      for (const listener of listeners) listener(current);
    },
    subscribe(listener) {
      listeners.add(listener);
      return () => listeners.delete(listener);
    },
  };
}
