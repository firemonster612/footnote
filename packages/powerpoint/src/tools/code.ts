import {
  describeOfficeError,
  type CodeRunResult,
  type FootnoteTool,
  type ToolEnv,
} from "@footnote/core/contracts";
import { Type } from "typebox";
import { codeHelpersReference } from "../advanced/index.ts";
import type { WriteGuard } from "../writeGuard.ts";
import { defineTool, SlideId } from "./schemas.ts";

const MAX_CODE_CHARS = 20_000;
const MAX_RESULT_CHARS = 8_000;
/** PowerPoint for the web applies large batches slowly; a 60s cap reported real, still-landing edits as failures. */
const CODE_TIMEOUT_MS = 300_000;

const clip = (text: string): string =>
  text.length > MAX_RESULT_CHARS
    ? `${text.slice(0, MAX_RESULT_CHARS)}\n… truncated (${text.length} chars); return less data.`
    : text;

/** Model-facing text for a code run. Exported for tests. */
export function formatCodeRun(run: CodeRunResult): string {
  const logs = run.logs.length > 0 ? `\nLogs:\n${run.logs.join("\n")}` : "";
  if (run.outcomeUnknown) {
    const reason = run.error ? ` (${run.error.message})` : "";
    return clip(
      `Outcome unknown${reason}: the code timed out or the connection dropped, so edits may or may not have applied. Do not rerun it: call get_deck or get_slide first, because the edits often landed.${logs}`,
    );
  }
  if (!run.ok) {
    const detail = run.error ? describeOfficeError(run.error) : "unknown error";
    return clip(
      `Error: ${detail}${logs}\nEdits synced before the failure stay applied. Check the slide with get_slide before rerunning.`,
    );
  }
  const result = run.result === undefined ? "undefined" : JSON.stringify(run.result);
  return clip(`Result: ${result}${logs}`);
}

export function createCodeTool(env: ToolEnv, guard: WriteGuard): FootnoteTool {
  return defineTool({
    name: "execute_office_js",
    label: "Run Office.js",
    access: "code",
    executionMode: "sequential",
    description: `Runs an async function body inside PowerPoint with \`context\` (PowerPoint.RequestContext) and \`footnote\` (helpers) in scope. \
Use it for what the structured tools can't do (bulk edits across many slides, properties they don't expose, reading something unusual), not for edits a structured tool covers. \
\`return\` a small JSON-safe value; \`console.log\` output is returned too. Results over ${MAX_RESULT_CHARS} characters are cut. \
List the slides the code will change in slideIds so undo can restore them. If the outcome is unknown, re-read before retrying.

Batch so the edits land together (every \`context.sync()\` is a round trip and a visible redraw):
1. Load everything you need first, for all slides at once (\`footnote.textShapes([...slides])\`, \`footnote.load(objects, "props")\`, or \`.load()\` on each, then one \`await context.sync()\`).
2. Queue all the writes, across every slide, without syncing in between.
3. \`await context.sync()\` once at the end.
Split into more syncs only when a later step needs a value from an earlier write (e.g. the ID of a shape you just added). Never sync inside a loop over slides or shapes.

Keep each run small: about 150 lines at most. Writing a long script takes the model minutes before anything appears, and one mistake wastes all of it. For a deck-wide redesign, build one or two slides per run, render them, then continue; define shared helpers again in each run rather than relying on earlier ones. Before setting properties on a shape you looked up (\`find\`, by name), check it exists and throw a clear error if it doesn't.

PowerPoint for the web rules (breaking them fails the run):
- Batches aren't transactional: when a sync fails, everything queued before the failing statement may already be applied. After an error, read the slide (get_slide) before running again, and make scripts safe to rerun (prefixed shape names + \`footnote.removeShapes\`).
- Never swallow sync errors (\`.catch(() => {})\`); a failed sync leaves the loaded values unusable and the next statements fail confusingly.
- Only GeometricShape, TextBox and Placeholder shapes have a \`textFrame\`. Load \`type\` first or use \`footnote.textShapes\`; touching \`textFrame\` on a line, picture, group or table fails the sync.
- Don't use a shape after \`delete()\`, and don't \`getItem\` an ID you haven't just loaded; IDs are strings scoped to their slide and change when a slide is re-inserted.
- Use string enum values ("Ellipse", "SendToBack", "Center"). Colors are "#RRGGBB"; \`fill.transparency\` is 0–1.
- There is no API for blur, gradients, shadows or glow; draw them with \`footnote.canvasImage\` and an image fill, or say it isn't possible.

${codeHelpersReference}`,
    parameters: Type.Object({
      code: Type.String({
        maxLength: MAX_CODE_CHARS,
        description: "Body of an async function. Use await and return.",
      }),
      explanation: Type.String({ description: "One sentence for the user: what the code does." }),
      slideIds: Type.Optional(
        Type.Array(SlideId, {
          description: "Slides the code modifies or deletes, snapshotted for undo first.",
        }),
      ),
    }),
    describeCall: (args: { explanation: string }) => args.explanation,
    async execute(_id, { code, slideIds = [] }, signal) {
      if (slideIds.length > 0) await guard.checkpoint(slideIds);
      const run = await env.host.runCode(code, {
        timeoutMs: CODE_TIMEOUT_MS,
        ...(signal && { signal }),
      });
      const slideChanges = run.slides && diffSlides(run.slides);
      // Slides the code created join the turn, so undo and revert remove them.
      if (slideChanges?.created.length)
        await guard.afterWrite({ createdSlideIds: slideChanges.created });
      return {
        content: [
          { type: "text", text: formatCodeRun(run) + formatSlideChanges(slideChanges, slideIds) },
        ],
        details: run,
        ...(!run.ok && { isError: true }),
      };
    },
  });
}

function diffSlides({ before, after }: { before: string[]; after: string[] }) {
  return {
    created: after.filter((id) => !before.includes(id)),
    deleted: before.filter((id) => !after.includes(id)),
  };
}

function formatSlideChanges(
  changes: ReturnType<typeof diffSlides> | undefined,
  snapshotted: string[],
): string {
  if (!changes) return "";
  const lines = [];
  if (changes.created.length > 0) lines.push(`Created slides: ${changes.created.join(", ")}.`);
  const unrecoverable = changes.deleted.filter((id) => !snapshotted.includes(id));
  if (changes.deleted.length > 0) lines.push(`Deleted slides: ${changes.deleted.join(", ")}.`);
  if (unrecoverable.length > 0)
    lines.push(
      `Undo can't bring back ${unrecoverable.join(", ")}: list slides in slideIds before deleting them in code.`,
    );
  return lines.length > 0 ? `\n${lines.join("\n")}` : "";
}
