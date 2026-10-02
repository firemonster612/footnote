// The write-guard hooks the advanced tools use. The WriteGuard (src/writeGuard.ts) implements them for the current chat.

export interface AdvancedToolDeps {
  /**
   * Call before writing to existing slides (including slides you're about to replace or delete).
   * Throws the model-facing "Slide N changed since you last read it; call get_slide again" when a slide differs from
   * the fingerprint the model last observed. Otherwise snapshots each slide for undo, once per turn.
   */
  beforeWrite(slideIds: string[]): Promise<void>;
  /**
   * Call after a write committed.
   * - createdSlideIds: slides the write inserted, including replacements (a chart merge or notes edit re-inserts the
   *   slide under a new ID). Undo deletes them.
   * - changedSlideIds: existing slides edited in place.
   * Refreshes the observed fingerprints of both sets so the model's next write isn't flagged as stale. Never throws:
   * the write already happened, so a failed refresh comes back as warnings for the result.
   */
  afterWrite(changes: {
    createdSlideIds?: string[];
    changedSlideIds?: string[];
  }): Promise<string[]>;
}
