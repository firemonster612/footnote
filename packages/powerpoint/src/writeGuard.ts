import type { OfficeHost } from "@footnote/core/contracts";
import type { AdvancedToolDeps } from "./advanced/deps.ts";
import { type ChatState, currentTurn } from "./chats.ts";
import type { SlideState, WriteReceipt } from "./ops/types.ts";

export interface WriteGuard extends AdvancedToolDeps {
  /** Snapshot for undo without the freshness check (moves, code runs). */
  checkpoint(slideIds: string[]): Promise<void>;
  /** Applies a core op's receipt: created slides join the turn, fingerprints become the model's view. */
  record(receipt: WriteReceipt): void;
  /** The model has read this slide at this fingerprint. */
  observe(slideId: string, fingerprint: string): void;
}

/** Read-before-write and undo checkpoints for one chat. `chat` resolves lazily because the env-to-chat binding can arrive after tool creation. */
export function createWriteGuard(host: OfficeHost, chat: () => ChatState): WriteGuard {
  async function prepare(slideIds: string[], checkFresh: boolean): Promise<void> {
    const state = chat();
    const turn = currentTurn(state);
    const exportSlideIds = slideIds.filter(
      (id) => !turn.snapshots.has(id) && !turn.created.has(id),
    );
    const slides = await host.call<Record<string, SlideState>>("get_slide_states", {
      slideIds,
      exportSlideIds,
    });
    for (const id of slideIds) {
      const slide = slides[id];
      if (!slide)
        throw new Error(`Slide ${id} not found. Use slide IDs from <deck_state> or get_deck.`);
      const seen = state.observed.get(id);
      if (checkFresh && seen !== undefined && seen !== slide.fingerprint) {
        throw new Error(
          `Slide ${slide.index + 1} changed since you last read it; call get_slide again.`,
        );
      }
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

    record(receipt) {
      const state = chat();
      const turn = currentTurn(state);
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
