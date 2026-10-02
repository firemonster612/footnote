import type { OfficeHost, UndoReport } from "@footnote/core/contracts";
import { type ChatState, MAX_UNDO_TURNS, type Turn, turnHasChanges } from "./chats.ts";
import type { RestoreArgs, RestoreResult, SlideState } from "./ops/types.ts";
import { isOutcomeUnknown, WRITE_TIMEOUT_MS } from "./writeGuard.ts";

const nothingToUndo: UndoReport = { restored: 0, removed: 0, warnings: ["Nothing to undo."] };

/** Undo of one turn, and whether it finished; an unfinished turn stays for another try. */
interface TurnUndo {
  report: UndoReport;
  complete: boolean;
}

/** Restores the slides the last changed turn touched, removes the slides it created, and reinserts the ones it deleted. */
export async function undoLastTurn(host: OfficeHost, chat: ChatState): Promise<UndoReport> {
  const turnIndex = chat.turns.findLastIndex(turnHasChanges);
  if (turnIndex === -1) return nothingToUndo;
  return (await undoTurnAt(host, chat, turnIndex)).report;
}

/**
 * Undoes the given turns newest first, so each restore sees the deck as the later turns left it. Stops at a turn that
 * didn't finish: older turns build on it.
 */
export async function undoTurns(
  host: OfficeHost,
  chat: ChatState,
  turnIds: string[],
): Promise<UndoReport> {
  const requested = new Set(turnIds);
  const total: UndoReport = { restored: 0, removed: 0, warnings: [] };
  for (let index = chat.turns.length - 1; index >= 0; index -= 1) {
    const turn = chat.turns[index]!;
    if (!requested.has(turn.id) || !turnHasChanges(turn)) continue;
    const { report, complete } = await undoTurnAt(host, chat, index);
    total.restored += report.restored;
    total.removed += report.removed;
    total.warnings.push(...report.warnings);
    if (!complete) break;
  }
  const lost = turnIds.filter((id) => !chat.seenTurnIds.has(id)).length;
  if (lost > 0)
    total.warnings.push(
      `Slide changes from ${lost} earlier ${lost === 1 ? "turn" : "turns"} couldn't be undone: their checkpoints last only until the extension's engine closes or the add-in pane reloads.`,
    );
  const evicted = turnIds.filter((id) => chat.evictedTurnIds.has(id)).length;
  if (evicted > 0)
    total.warnings.push(
      `Slide changes from ${evicted} older ${evicted === 1 ? "turn" : "turns"} couldn't be undone: only the last ${MAX_UNDO_TURNS} turns that changed slides keep checkpoints.`,
    );
  total.warnings = [...new Set(total.warnings)];
  return total;
}

async function undoTurnAt(host: OfficeHost, chat: ChatState, turnIndex: number): Promise<TurnUndo> {
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

  const args: RestoreArgs = {
    deleteSlideIds: touched.filter((id) => current[id]),
    inserts: [...turn.snapshots].map(([slideId, snapshot]) => ({ slideId, ...snapshot })),
  };
  let result: RestoreResult;
  try {
    result = await host.call<RestoreResult>("restore_slides", args, {
      timeoutMs: WRITE_TIMEOUT_MS,
    });
  } catch (error) {
    if (!isOutcomeUnknown(error)) throw error;
    const report = {
      restored: 0,
      removed: 0,
      warnings: [
        `Undo's outcome is unknown (${error.message}): it may have applied. Check the deck before undoing again.`,
      ],
    };
    return { report, complete: false };
  }

  const removed = new Set(result.removedSlideIds);
  // A copy replaced its original once the original is gone; a copy next to a surviving original is a leftover.
  const replaced = Object.fromEntries(
    Object.entries(result.idMap).filter(([id]) => !current[id] || removed.has(id)),
  );
  const createdRemoved = result.removedSlideIds.filter((id) => turn.created.has(id)).length;
  for (const id of removed) chat.observed.delete(id);
  if (result.error) {
    // Keep what's left of the turn so undoing again finishes the job: replaced slides and removed ones are done,
    // and leftover copies join the slides to remove.
    for (const id of Object.keys(replaced)) turn.snapshots.delete(id);
    for (const id of removed) turn.created.delete(id);
    for (const [id, copy] of Object.entries(result.idMap))
      if (!(id in replaced)) turn.created.add(copy);
    warnings.push(`Undo stopped part-way (${result.error}). Undo again to finish.`);
  } else {
    chat.turns.splice(turnIndex, 1);
  }
  // Older turns still name the slides by their pre-restore IDs; without this, undoing them next would
  // re-insert a second copy instead of replacing the restored slide.
  for (const earlier of chat.turns) if (earlier !== turn) remapSlideIds(earlier, replaced);
  if (Object.keys(replaced).length > 0) warnings.push("Restored slides have new slide IDs.");
  const report = { restored: Object.keys(replaced).length, removed: createdRemoved, warnings };
  return { report, complete: !result.error };
}

function remapSlideIds(turn: Turn, idMap: Record<string, string>): void {
  const rename = (id: string) => idMap[id] ?? id;
  turn.snapshots = new Map([...turn.snapshots].map(([id, snapshot]) => [rename(id), snapshot]));
  turn.created = new Set([...turn.created].map(rename));
  turn.after = new Map([...turn.after].map(([id, fingerprint]) => [rename(id), fingerprint]));
}
