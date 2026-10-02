// Live smoke test against a real endpoint: `FOOTNOTE_ENDPOINT=http://localhost:8317 FOOTNOTE_API_KEY=… bun
// packages/core/test/runtime/liveSmoke.ts [modelId:level ...]`. Streams text and runs one tool call per model through
// createFootnoteApp, and checks the requested effort reached the request body. Without FOOTNOTE_API_KEY it reads the
// first key from a local CLIProxyAPI config; it never prints the key.
import { readFileSync } from "node:fs";
import { homedir } from "node:os";
import { createFootnoteApp, type ThinkingLevel } from "../../src/index.ts";
import { fakeHost, fakeHostModule, memoryStore } from "./fakes.ts";

const baseUrl = process.env.FOOTNOTE_ENDPOINT ?? "http://localhost:8317";
const targets =
  process.argv.length > 2 ? process.argv.slice(2) : ["claude-sonnet-5-5:low", "gpt-6-luna:high"];

function readApiKey(): string {
  if (process.env.FOOTNOTE_API_KEY) return process.env.FOOTNOTE_API_KEY;
  const lines = readFileSync(`${homedir()}/.local/share/cliproxy-api/config.yaml`, "utf8").split(
    "\n",
  );
  const start = lines.findIndex((line) => line.startsWith("api-keys:"));
  const key = lines
    .slice(start + 1)
    .map((line) => /^\s*-\s*["']?([^"'\s]+)["']?\s*$/.exec(line)?.[1])
    .find((value) => value !== undefined);
  if (!key) throw new Error("No api-keys entry in the CLIProxyAPI config.");
  return key;
}

const requestBodies: unknown[] = [];
const realFetch = globalThis.fetch;
globalThis.fetch = (async (input: RequestInfo | URL, init?: RequestInit) => {
  if (typeof init?.body === "string") requestBodies.push(JSON.parse(init.body));
  return realFetch(input, init);
}) as typeof fetch;

/** Where each API puts the effort: Anthropic output_config (top-level or per system message), OpenAI reasoning. */
function effortsIn(body: unknown): string[] {
  const found: string[] = [];
  JSON.stringify(body, (key, value) => {
    if (key === "output_config" || key === "reasoning") found.push(value?.effort);
    if (key === "reasoning_effort") found.push(value);
    return value;
  });
  return found.filter((effort): effort is string => typeof effort === "string");
}

const store = memoryStore();
await store.set("settings", {
  endpoint: { baseUrl, apiKey: readApiKey() },
  defaultPermissionMode: "full",
});
const { hostModule, contextCalls } = fakeHostModule();
const app = await createFootnoteApp({ host: fakeHost, hostModule, store });
const models = await app.models.list();
console.log(`listed ${models.length} chat models`);

let failed = false;
for (const target of targets) {
  const [modelId, level] = target.split(":") as [string, ThinkingLevel];
  requestBodies.length = 0;
  const session = await app.chats.create("smoke-doc");
  session.setModel(modelId);
  session.setThinkingLevel(level);
  let streamedTextUpdates = 0;
  session.subscribe((state) => {
    const streaming = state.streamingMessage;
    if (
      streaming?.role === "assistant" &&
      streaming.content.some((block) => block.type === "text" && block.text)
    ) {
      streamedTextUpdates += 1;
    }
  });

  await session.send("Call the get_deck tool once, then reply with exactly the word FINISHED.");

  const state = session.getState();
  const toolResults = state.messages.filter((message) => message.role === "toolResult");
  const final = state.messages.at(-1);
  const finalText =
    final?.role === "assistant"
      ? final.content.flatMap((block) => (block.type === "text" ? [block.text] : [])).join("")
      : "";
  const efforts = requestBodies.flatMap(effortsIn);
  const ok =
    !state.error &&
    toolResults.length >= 1 &&
    /FINISHED/.test(finalText) &&
    streamedTextUpdates > 1 &&
    efforts.includes(level);
  failed ||= !ok;
  console.log(
    JSON.stringify({
      model: modelId,
      requestedLevel: level,
      ok,
      error: state.error,
      toolCalls: toolResults.map((message) => message.toolName),
      finalText,
      streamedTextUpdates,
      effortsInRequestBodies: efforts,
      requests: requestBodies.length,
      contextUsage: state.contextUsage,
    }),
  );
}
console.log(`context blocks fetched: ${contextCalls.length}`);
process.exit(failed ? 1 : 0);
