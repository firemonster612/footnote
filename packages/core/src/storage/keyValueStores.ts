import type { KeyValueStore } from "../contracts.ts";

/** The slice of chrome.storage.local we use; core doesn't depend on @types/chrome. */
interface ChromeStorageArea {
  get(key: string): Promise<Record<string, unknown>>;
  set(items: Record<string, unknown>): Promise<void>;
  remove(key: string): Promise<void>;
}

declare const chrome: { storage: { local: ChromeStorageArea } };

/** chrome.storage.local, for the extension side panel. */
export function createChromeStore(): KeyValueStore {
  const area = chrome.storage.local;
  return {
    async get<T>(key: string) {
      const items = await area.get(key);
      // chrome.storage returns what we stored under this key; set<T> is the only writer.
      return items[key] as T | undefined;
    },
    set: (key, value) => area.set({ [key]: value }),
    delete: (key) => area.remove(key),
  };
}

/** localStorage with JSON values, for the add-in task pane. */
export function createLocalStore(prefix = "footnote:"): KeyValueStore {
  return {
    async get<T>(key: string) {
      const raw = localStorage.getItem(prefix + key);
      // Values under our prefix are only written by set<T> below.
      return raw === null ? undefined : (JSON.parse(raw) as T);
    },
    async set(key, value) {
      localStorage.setItem(prefix + key, JSON.stringify(value));
    },
    async delete(key) {
      localStorage.removeItem(prefix + key);
    },
  };
}

const objectStoreName = "kv";

function requestResult<T>(request: IDBRequest<T>): Promise<T> {
  return new Promise((resolve, reject) => {
    request.onsuccess = () => resolve(request.result);
    request.onerror = () => reject(request.error);
  });
}

/**
 * A key-value store in IndexedDB, for data too large for chrome.storage/localStorage quotas (chat transcripts with
 * images). Resolves to undefined when IndexedDB is unavailable or refuses to open, so callers can fall back.
 */
export async function openIndexedDbStore(databaseName: string): Promise<KeyValueStore | undefined> {
  if (typeof indexedDB === "undefined") return undefined;
  const request = indexedDB.open(databaseName, 1);
  request.onupgradeneeded = () => request.result.createObjectStore(objectStoreName);
  const db = await requestResult(request).catch(() => undefined);
  if (!db) return undefined;

  const objectStore = (mode: IDBTransactionMode) =>
    db.transaction(objectStoreName, mode).objectStore(objectStoreName);
  return {
    async get<T>(key: string) {
      // Values in this object store are only written by set<T> below.
      return (await requestResult(objectStore("readonly").get(key))) as T | undefined;
    },
    async set(key, value) {
      await requestResult(objectStore("readwrite").put(value, key));
    },
    async delete(key) {
      await requestResult(objectStore("readwrite").delete(key));
    },
  };
}
