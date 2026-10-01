import type { OfficeHost, UndoReport } from "@footnote/core/contracts";
import { type ChatState, type Turn, turnHasChanges } from "./chats.ts";
import type { SlideState } from "./ops/types.ts";

interface RestoreResult {
  removedSlideIds: string[];
  restoredSlideIds: string[];
  /** Original slide ID → ID of its restored copy (re-inserted slides always get new IDs). */
  idMap: Record<string, string>;
}

const nothingToUndo: UndoReport = { restored: 0, removed: 0, warnings: ["Nothing to undo."] };

/** Restores the slides the last changed turn touched, removes the slides it created, and reinserts the ones it deleted. */
export async function undoLastTurn(host: OfficeHost, chat: ChatState): Promise<UndoReport> {
  const turnIndex = chat.turns.findLastIndex(turnHasChanges);
  if (turnIndex === -1) return nothingToUndo;
  return undoTurnAt(host, chat, turnIndex);
}

/** Undoes the given turns newest first, so each restore sees the deck as the later turns left it. */
export async function undoTurns(
  host: OfficeHost,
  chat: ChatState,
  turnIds: string[],
): Promise<UndoReport> {
  const requested = new Set(turnIds);
  const total: UndoReport = { restored: 0, removed: 0, warnings: [] };
  const lost = turnIds.filter((id) => !chat.seenTurnIds.has(id)).length;
  for (let index = chat.turns.length - 1; index >= 0; index -= 1) {
    const turn = chat.turns[index]!;
    if (!requested.has(turn.id) || !turnHasChanges(turn)) continue;
    const report = await undoTurnAt(host, chat, index);
    total.restored += report.restored;
    total.removed += report.removed;
    total.warnings.push(...report.warnings);
  }
  if (lost > 0)
    total.warnings.push(
      `Slide changes from ${lost} earlier ${lost === 1 ? "turn" : "turns"} couldn't be undone: their checkpoints are kept only until the panel reloads.`,
    );
  total.warnings = [...new Set(total.warnings)];
  return total;
}

async function undoTurnAt(host: OfficeHost, chat: ChatState, turnIndex: number): Promise<UndoReport> {
  const turn = chat.turns[turnIndex]!;
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
    inserts: [...turn.snapshots].map(([slideId, snapshot]) => ({ slideId, ...snapshot })),
  });
  chat.turns.splice(turnIndex, 1);
  for (const id of result.removedSlideIds) chat.observed.delete(id);
  // Older turns still name the slides by their pre-restore IDs; without this, undoing them next would
  // re-insert a second copy instead of replacing the restored slide.
  for (const earlier of chat.turns) remapSlideIds(earlier, result.idMap);
  if (turn.snapshots.size > 0) warnings.push("Restored slides have new slide IDs.");
  return {
    restored: result.restoredSlideIds.length,
    removed: result.removedSlideIds.filter((id) => turn.created.has(id)).length,
    warnings,
  };
}

function remapSlideIds(turn: Turn, idMap: Record<string, string>): void {
  const rename = (id: string) => idMap[id] ?? id;
  turn.snapshots = new Map([...turn.snapshots].map(([id, snapshot]) => [rename(id), snapshot]));
  turn.created = new Set([...turn.created].map(rename));
  turn.after = new Map([...turn.after].map(([id, fingerprint]) => [rename(id), fingerprint]));
}
