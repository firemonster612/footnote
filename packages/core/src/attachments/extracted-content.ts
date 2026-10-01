import type { ProcessedAttachment } from "../contracts.ts";

/** What a format processor adds on top of the file's metadata and raw bytes. */
export type ExtractedContent = Pick<ProcessedAttachment, "text" | "images" | "summary">;

export type AttachmentImage = NonNullable<ProcessedAttachment["images"]>[number];
