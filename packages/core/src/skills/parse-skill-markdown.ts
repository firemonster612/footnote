import type { SkillDefinition } from "../contracts.ts";

const FRONT_MATTER = /^﻿?---\r?\n([\s\S]*?)\r?\n---[ \t]*(?:\r?\n|$)/;
const KEY_LINE = /^([A-Za-z][\w-]*):[ \t]*(.*)$/;
const BLOCK_SCALAR = /^[|>][+-]?$/;

/** Reads a SKILL.md: front matter with `name` and `description`, then the markdown body. */
export function parseSkillMarkdown(
  markdown: string,
  source: SkillDefinition["source"],
): SkillDefinition {
  const match = FRONT_MATTER.exec(markdown);
  const fields = match?.[1] ? parseFrontMatter(match[1]) : new Map<string, string>();
  const name = fields.get("name");
  const description = fields.get("description");
  if (!name || !description) {
    throw new Error(
      "Not a valid SKILL.md: it must start with front matter (---) containing name and description.",
    );
  }
  return { name, description, body: markdown.slice(match?.[0].length).trim(), source };
}

/**
 * The YAML subset skills use: `key: value` lines, quoted values, indented continuation lines, and
 * `|` / `>` block scalars. Nested structures are skipped.
 */
function parseFrontMatter(yaml: string): Map<string, string> {
  const entries: { key: string; style: string; lines: string[] }[] = [];
  for (const line of yaml.split(/\r?\n/)) {
    const keyLine = /^\s/.test(line) ? null : KEY_LINE.exec(line);
    if (keyLine) {
      const [, key = "", value = ""] = keyLine;
      const isBlock = BLOCK_SCALAR.test(value.trim());
      entries.push({
        key,
        style: isBlock ? value.trim()[0]! : "plain",
        lines: isBlock ? [] : [value],
      });
    } else {
      entries.at(-1)?.lines.push(line.trim());
    }
  }
  return new Map(entries.map(({ key, style, lines }) => [key, joinScalar(style, lines)]));
}

function joinScalar(style: string, lines: string[]): string {
  if (style === "|") return lines.join("\n").trim();
  const value = lines.filter(Boolean).join(" ").trim();
  const quoted = /^(["'])([\s\S]*)\1$/.exec(value);
  if (!quoted) return value;
  return quoted[1] === '"' ? quoted[2]!.replace(/\\"/g, '"') : quoted[2]!.replace(/''/g, "'");
}
