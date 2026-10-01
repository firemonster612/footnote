import type { SkillDefinition } from "../contracts.ts";

const DEFAULT_BUDGET_CHARS = 4_000;
const MIN_DESCRIPTION_CHARS = 40;
const HEADER =
  "## Skills\n" +
  "Skills hold instructions for specific kinds of work. When a task matches a skill below, call load_skill with " +
  "its name before starting, and follow it.";

/**
 * The system-prompt section listing skills by name and description, kept within `budgetChars`.
 * Over budget, descriptions are shortened evenly; if names alone don't fit, the tail is summarized.
 */
export function formatSkillListing(
  skills: SkillDefinition[],
  budgetChars = DEFAULT_BUDGET_CHARS,
): string {
  if (skills.length === 0) return "";
  const render = (lines: string[]) => [HEADER, ...lines].join("\n");

  const full = render(skills.map(({ name, description }) => `- ${name}: ${description}`));
  if (full.length <= budgetChars) return full;

  const nameChars = skills.reduce((sum, { name }) => sum + `\n- ${name}: `.length, HEADER.length);
  const descriptionChars = Math.floor((budgetChars - nameChars) / skills.length);
  if (descriptionChars >= MIN_DESCRIPTION_CHARS) {
    return render(
      skills.map(
        ({ name, description }) => `- ${name}: ${truncate(description, descriptionChars)}`,
      ),
    );
  }

  const names = skills.map(({ name }) => `- ${name}`);
  const listNames = (shown: number) =>
    render(
      shown < names.length
        ? [...names.slice(0, shown), `- …and ${names.length - shown} more`]
        : names,
    );
  let shown = names.length;
  while (shown > 0 && listNames(shown).length > budgetChars) shown--;
  return listNames(shown);
}

function truncate(text: string, maxChars: number): string {
  return text.length <= maxChars ? text : `${text.slice(0, maxChars - 1).trimEnd()}…`;
}
