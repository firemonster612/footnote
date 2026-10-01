import type { StreamFn } from "@earendil-works/pi-agent-core";
import {
  createAssistantMessageEventStream,
  type AssistantMessage,
  type AssistantMessageEventStream,
} from "@earendil-works/pi-ai";

/**
 * Models occasionally mangle a tool name (Sonnet 5.5 sent `someone_execute_office_js` five times in a row).
 * Maps an unknown name to the one registered tool it ends with; leaves anything ambiguous alone so the
 * normal "tool not found" error still reaches the model.
 */
export function repairToolName(name: string, knownNames: readonly string[]): string {
  if (knownNames.includes(name)) return name;
  const matches = knownNames.filter(
    (known) => name.endsWith(`_${known}`) || name.endsWith(`.${known}`),
  );
  return matches.length === 1 ? matches[0]! : name;
}

function repairMessage(message: AssistantMessage, knownNames: readonly string[]): void {
  for (const block of message.content) {
    if (block.type === "toolCall") block.name = repairToolName(block.name, knownNames);
  }
}

/** Wraps a stream function so tool calls reach the agent loop with repaired names. */
export function withToolNameRepair(
  streamFn: StreamFn,
  knownNames: () => readonly string[],
): StreamFn {
  return async (model, context, options) => {
    const source = await streamFn(model, context, options);
    const repaired: AssistantMessageEventStream = createAssistantMessageEventStream();
    void (async () => {
      try {
        for await (const event of source) {
          if (event.type === "toolcall_end") {
            event.toolCall.name = repairToolName(event.toolCall.name, knownNames());
            repairMessage(event.partial, knownNames());
          }
          if (event.type === "done") repairMessage(event.message, knownNames());
          repaired.push(event);
        }
      } finally {
        repaired.end(await source.result());
      }
    })();
    return repaired;
  };
}
