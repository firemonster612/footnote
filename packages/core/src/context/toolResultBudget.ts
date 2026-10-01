import type { ImageContent, TextContent } from "@earendil-works/pi-ai";

/** About 4k tokens of text per tool result. */
export const toolResultTextBudgetChars = 16_000;

/**
 * Caps the combined text of a tool result, keeping images, and appends a note telling the model what was cut.
 * Returns undefined when the result already fits.
 */
export function capToolResultText(
  content: (TextContent | ImageContent)[],
  budgetChars = toolResultTextBudgetChars,
): (TextContent | ImageContent)[] | undefined {
  const totalChars = content.reduce(
    (sum, block) => sum + (block.type === "text" ? block.text.length : 0),
    0,
  );
  if (totalChars <= budgetChars) return undefined;

  let remaining = budgetChars;
  const capped = content.map((block) => {
    if (block.type !== "text") return block;
    const text = block.text.slice(0, remaining);
    remaining -= text.length;
    return { ...block, text };
  });
  const note = `[Truncated: ${totalChars - budgetChars} of ${totalChars} characters omitted. Request less at once (pagination, offset/limit, or a narrower target) to see the rest.]`;
  return [
    ...capped.filter((block) => block.type !== "text" || block.text.length > 0),
    { type: "text", text: note },
  ];
}
