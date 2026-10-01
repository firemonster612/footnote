import type { ImageContent, TextContent } from "@earendil-works/pi-ai";
import type { ProcessedAttachment } from "../contracts.ts";

/** ~6k tokens at ~4 characters per token. */
const EXCERPT_CHARS = 24_000;
const MAX_IMAGES = 6;

/** The message content the model sees when the user sends an attachment: a text excerpt plus up to 6 images. */
export function attachmentToContent(
  attachment: ProcessedAttachment,
): (TextContent | ImageContent)[] {
  const { id, name, kind, summary, text } = attachment;
  const images = (attachment.images ?? []).slice(0, MAX_IMAGES);
  const attributes = [`id=${JSON.stringify(id)}`, `name=${JSON.stringify(name)}`, `kind="${kind}"`];
  if (summary) attributes.push(`summary=${JSON.stringify(summary)}`);

  const lines = [`<attachment ${attributes.join(" ")}>`];
  if (text) lines.push(text.slice(0, EXCERPT_CHARS));
  lines.push("</attachment>");
  if (text && text.length > EXCERPT_CHARS) {
    lines.push(
      `[Excerpt: first ${EXCERPT_CHARS.toLocaleString("en-US")} of ${text.length.toLocaleString("en-US")} characters. ` +
        `Call read_attachment with id "${id}" and offset ${EXCERPT_CHARS} to read more.]`,
    );
  }
  if (images.length > 0) {
    const more = kind === "pdf" ? " read_attachment with pages renders any other page." : "";
    lines.push(`[Images below: ${images.map((image) => image.label).join(", ")}.${more}]`);
  }

  return [
    { type: "text", text: lines.join("\n") },
    ...images.map((image): ImageContent => ({
      type: "image",
      data: image.base64,
      mimeType: image.mimeType,
    })),
  ];
}
