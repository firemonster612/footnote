import mammoth from "mammoth";
import type { ExtractedContent } from "./extracted-content.ts";

/**
 * mammoth 1.x still exports convertToMarkdown but its typings dropped it. Its Node build reads
 * `buffer` and its browser build reads `arrayBuffer`; both accept an ArrayBuffer.
 */
interface MammothMarkdown {
  convertToMarkdown(
    input: { arrayBuffer: ArrayBuffer; buffer: ArrayBuffer },
    options: { convertImage: unknown },
  ): Promise<{ value: string }>;
}

// Inline images would otherwise become data URIs in the markdown; keep a short marker instead.
const imagePlaceholder = mammoth.images.imgElement(async () => ({ src: "embedded-image" }));

// mammoth backslash-escapes every period, hyphen, and parenthesis. The text is for the model to read,
// not to re-render, so the escapes are noise.
const ESCAPED_PUNCTUATION = /\\([\\`*_{}[\]()#+\-.!])/g;

export async function processDocx(bytes: ArrayBuffer): Promise<ExtractedContent> {
  // See MammothMarkdown for why this cast is sound.
  const { value } = await (mammoth as typeof mammoth & MammothMarkdown).convertToMarkdown(
    { arrayBuffer: bytes, buffer: bytes },
    { convertImage: imagePlaceholder },
  );
  const text = value.replace(ESCAPED_PUNCTUATION, "$1").trim();
  // Tokens with a letter or digit, so markdown markers like "#" and "-" don't count as words.
  const words = text.match(/\S*[\p{L}\p{N}]\S*/gu)?.length ?? 0;
  return { text, summary: `${words.toLocaleString("en-US")} words` };
}
