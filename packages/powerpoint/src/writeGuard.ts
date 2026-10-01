import { type OfficeHost, OfficeOpError } from "@footnote/core/contracts";
import type { AdvancedToolDeps } from "./advanced/deps.ts";
import { type ChatState, currentTurn, type Turn } from "./chats.ts";
import { staleSlideMessage } from "./ops/guard.ts";
import { slideNotFound } from "./ops/presentation.ts";
import type { SlideState, WriteGuardArgs, WriteReceipt } from "./ops/types.ts";

export interface WriteGuard extends AdvancedToolDeps {
  /**
   * Guard args for a write op on these slides: the fingerprints the model last read (the op fails on a stale slide
   * before changing anything) and the slides this turn hasn't snapshotted yet (the op exports them for undo).
   */
  writeArgs(slideIds: string[]): Required<WriteGuardArgs>;
  /** Snapshot for undo without the freshness check (code runs). */
  checkpoint(slideIds: string[]): Promise<void>;
  /** Applies a write op's receipt: snapshots and created slides join the turn, fingerprints become the model's view. */
  record(receipt: WriteReceipt): void;
  /** The model has read this slide at this fingerprint. */
  observe(slideId: string, fingerprint: string): void;
}

/**
 * Errors an op threw itself carry no Office error code and are already written for the model ("Slide 2 changed since
 * you last read it; …"), so they lose the "op failed:" wrapper. Office errors keep it, with their code and statement.
 */
export function modelFacingError(error: unknown): unknown {
  return error instanceof OfficeOpError && error.info.code === undefined
    ? new Error(error.info.message, { cause: error })
    : error;
}

const needsSnapshot = (turn: Turn, slideIds: string[]): string[] =>
  slideIds.filter((id) => !turn.snapshots.has(id) && !turn.created.has(id));

/** Read-before-write and undo checkpoints for one chat. `chat` resolves lazily because the env-to-chat binding can arrive after tool creation. */
export function createWriteGuard(host: OfficeHost, chat: () => ChatState): WriteGuard {
  async function prepare(slideIds: string[], checkFresh: boolean): Promise<void> {
    const state = chat();
    const turn = currentTurn(state);
    const exportSlideIds = needsSnapshot(turn, slideIds);
    const slides = await host.call<Record<string, SlideState>>("get_slide_states", {
      slideIds,
      exportSlideIds,
    });
    for (const id of slideIds) {
      const slide = slides[id];
      if (!slide) throw slideNotFound(id);
      const seen = state.observed.get(id);
      if (checkFresh && seen !== undefined && seen !== slide.fingerprint)
        throw new Error(staleSlideMessage(slide.index + 1));
    }
    for (const id of exportSlideIds) {
      const slide = slides[id];
      if (slide?.base64) turn.snapshots.set(id, { index: slide.index, base64: slide.base64 });
    }
  }

  function recordFingerprints(fingerprints: Record<string, string>): void {
    const state = chat();
    const turn = currentTurn(state);
    for (const [id, fingerprint] of Object.entries(fingerprints)) {
      state.observed.set(id, fingerprint);
      turn.after.set(id, fingerprint);
    }
  }

  return {
    beforeWrite: (slideIds) => prepare(slideIds, true),
    checkpoint: (slideIds) => prepare(slideIds, false),

    writeArgs(slideIds) {
      const state = chat();
      return {
        expectedFingerprints: Object.fromEntries(
          slideIds.flatMap((id) => {
            const seen = state.observed.get(id);
            return seen === undefined ? [] : [[id, seen]];
          }),
        ),
        snapshotSlideIds: needsSnapshot(currentTurn(state), slideIds),
      };
    },

    record(receipt) {
      const state = chat();
      const turn = currentTurn(state);
      for (const [id, snapshot] of Object.entries(receipt.snapshots ?? {}))
        if (!turn.snapshots.has(id)) turn.snapshots.set(id, snapshot);
      for (const id of receipt.createdSlideIds ?? []) turn.created.add(id);
      for (const id of receipt.deletedSlideIds ?? []) state.observed.delete(id);
      recordFingerprints(receipt.fingerprints);
    },

    async afterWrite({ createdSlideIds = [], changedSlideIds = [] }) {
      const turn = currentTurn(chat());
      for (const id of createdSlideIds) turn.created.add(id);
      const slideIds = [...createdSlideIds, ...changedSlideIds];
      if (slideIds.length === 0) return;
      const slides = await host.call<Record<string, SlideState>>("get_slide_states", { slideIds });
      recordFingerprints(
        Object.fromEntries(Object.entries(slides).map(([id, slide]) => [id, slide.fingerprint])),
      );
    },

    observe(slideId, fingerprint) {
      chat().observed.set(slideId, fingerprint);
    },
  };
}
