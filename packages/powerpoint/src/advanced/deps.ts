// Owned by B1. B2 reads this; don't edit it from src/advanced.
//
// `createTools(env)` builds one AdvancedToolDeps per chat and passes it to `createAdvancedTools(env, deps)`,
// so every hook is already bound to the chat and its current turn.

export interface AdvancedToolDeps {
  /**
   * Call before writing to existing slides (including slides you're about to replace or delete).
   * Throws an Error with a model-facing message ("Slide N changed since you last read it; call get_slide again")
   * when a slide's fingerprint differs from the one the model last observed. Otherwise snapshots each slide
   * (exportAsBase64 + index) for undo, once per turn. Let the error propagate so the tool call fails.
   */
  beforeWrite(slideIds: string[]): Promise<void>;
  /**
   * Call after a successful write.
   * - createdSlideIds: slides the write inserted, including replacements (a chart merge or notes edit that
   *   re-inserts a slide yields a new ID). Undo deletes them.
   * - changedSlideIds: existing slides edited in place.
   * Refreshes the observed fingerprints of both sets so the model's next write isn't flagged as stale.
   */
  afterWrite(changes: { createdSlideIds?: string[]; changedSlideIds?: string[] }): Promise<void>;
}

// Core ops B2 may call through `env.host.call(name, args)` (defined in src/ops/):
// - "get_deck" → { slideWidth, slideHeight, slides: [{ id, index, layout, title }], ... } (points)
// - "export_slide" { slideId } → base64 PPTX of that slide
// - "insert_slides" { base64, targetSlideId?, formatting?: "KeepSourceFormatting" | "UseDestinationTheme", sourceSlideIds? } → { createdSlideIds }
// - "delete_slides" { slideIds }
