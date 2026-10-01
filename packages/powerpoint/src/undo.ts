import type { OfficeHost, UndoReport } from "@footnote/core/contracts";
import { type ChatState, turnHasChanges } from "./chats.ts";
import type { SlideState } from "./ops/types.ts";

interface RestoreResult {
  removedSlideIds: string[];
  restoredSlideIds: string[];
}

/** Restores the slides the last changed turn touched, removes the slides it created, and reinserts the ones it deleted. */
export async function undoLastTurn(host: OfficeHost, chat: ChatState): Promise<UndoReport> {
  const turnIndex = chat.turns.findLastIndex(turnHasChanges);
  const turn = chat.turns[turnIndex];
  if (!turn) return { restored: 0, removed: 0, warnings: ["Nothing to undo."] };

  const touched = [...turn.snapshots.keys(), ...turn.created];
  const current = await host.call<Record<string, SlideState>>("get_slide_states", {
    slideIds: touched,
  });
  const warnings = [...turn.after].flatMap(([id, fingerprint]) => {
    const slide = current[id];
    return slide && slide.fingerprint !== fingerprint
      ? [`Slide ${slide.index + 1} was edited after that turn; undo discarded those later edits.`]
      : [];
  });

  const result = await host.call<RestoreResult>("restore_slides", {
    deleteSlideIds: touched.filter((id) => current[id]),
    inserts: [...turn.snapshots.values()],
  });
  chat.turns.splice(turnIndex, 1);
  for (const id of result.removedSlideIds) chat.observed.delete(id);
  if (turn.snapshots.size > 0) warnings.push("Restored slides have new slide IDs.");
  return {
    restored: result.restoredSlideIds.length,
    removed: result.removedSlideIds.filter((id) => turn.created.has(id)).length,
    warnings,
  };
}
