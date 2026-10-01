import { Type } from "typebox";
import { defineTool } from "../attachments/define-tool.ts";
import type { FootnoteTool, SkillDefinition } from "../contracts.ts";

const parameters = Type.Object({
  name: Type.String({ description: "Skill name from the skill listing." }),
  file: Type.Optional(
    Type.String({
      description: "A supporting file the skill names, instead of its main instructions.",
    }),
  ),
});

/** `load_skill`: returns a skill's instructions, or one of its supporting files. */
export function createSkillTool(getSkills: () => SkillDefinition[]): FootnoteTool {
  return defineTool({
    name: "load_skill",
    label: "Load skill",
    access: "read",
    description:
      "Load a skill's full instructions by name (see Skills in the system prompt). Load it before starting work the " +
      "skill covers. Pass file to load one of the supporting files a skill lists.",
    parameters,
    describeCall: (args: { name: string; file?: string }) =>
      args.file ? `Load ${args.file} from skill ${args.name}` : `Load skill ${args.name}`,
    execute: async (_toolCallId, { name, file }) => {
      const skills = getSkills();
      const skill = skills.find((candidate) => candidate.name === name);
      if (!skill) {
        throw new Error(
          `No skill named "${name}". Available: ${skills.map((s) => s.name).join(", ") || "none"}.`,
        );
      }
      return {
        content: [{ type: "text", text: file ? readFile(skill, file) : formatSkill(skill) }],
        details: undefined,
      };
    },
  });
}

function formatSkill({ name, body, files }: SkillDefinition): string {
  const fileNames = Object.keys(files ?? {});
  const fileNote =
    fileNames.length > 0
      ? `\n\nSupporting files (load_skill with name "${name}" and file): ${fileNames.join(", ")}`
      : "";
  return `# Skill: ${name}\n\n${body}${fileNote}`;
}

function readFile({ name, files = {} }: SkillDefinition, file: string): string {
  const content = files[file];
  if (content === undefined) {
    throw new Error(
      `Skill "${name}" has no file "${file}". Files: ${Object.keys(files).join(", ") || "none"}.`,
    );
  }
  return `# ${name}/${file}\n\n${content}`;
}
