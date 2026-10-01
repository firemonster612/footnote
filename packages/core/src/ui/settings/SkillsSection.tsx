import { Trash } from "lucide-react";
import { useState } from "react";
import type { FootnoteApp, Settings, SkillDefinition } from "../../contracts.ts";
import { fetchSkillFromUrl, parseSkillMarkdown } from "../../skills/index.ts";
import { Button, IconButton, TextInput } from "../components/controls.tsx";
import { errorMessage } from "../hooks.ts";

export function SkillsSection({ app, settings }: { app: FootnoteApp; settings: Settings }) {
  const [markdown, setMarkdown] = useState("");
  const [url, setUrl] = useState("");
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string>();
  const skills = settings.customSkills;

  /** Adds a skill, replacing any custom skill with the same name. */
  async function addSkill(load: () => Promise<SkillDefinition>, clearInput: () => void) {
    setBusy(true);
    setError(undefined);
    try {
      const skill: SkillDefinition = { ...(await load()), source: "custom" };
      await app.settings.update({
        customSkills: [...skills.filter((existing) => existing.name !== skill.name), skill],
      });
      clearInput();
    } catch (cause) {
      setError(errorMessage(cause));
    } finally {
      setBusy(false);
    }
  }

  const removeSkill = (name: string) =>
    app.settings.update({ customSkills: skills.filter((skill) => skill.name !== name) });

  return (
    <div className="flex flex-col gap-2.5">
      {skills.length > 0 && (
        <ul className="flex flex-col divide-y divide-neutral-200 rounded-md border border-neutral-200 dark:divide-neutral-700 dark:border-neutral-700">
          {skills.map((skill) => (
            <li key={skill.name} className="flex items-start gap-2 py-1.5 pr-1 pl-2">
              <div className="min-w-0 flex-1">
                <div className="truncate font-medium">{skill.name}</div>
                <div className="line-clamp-2 text-[12px] text-neutral-500">{skill.description}</div>
              </div>
              <IconButton
                icon={Trash}
                label={`Remove ${skill.name}`}
                onClick={() => void removeSkill(skill.name)}
              />
            </li>
          ))}
        </ul>
      )}
      <textarea
        value={markdown}
        onChange={(event) => setMarkdown(event.target.value)}
        placeholder={"---\nname: my-skill\ndescription: When to use it\n---\n\nInstructions"}
        aria-label="SKILL.md content"
        rows={4}
        spellCheck={false}
        className="w-full resize-y rounded-md border border-neutral-300 bg-white px-2 py-1.5 font-mono text-[12px] placeholder:text-neutral-400 dark:border-neutral-600 dark:bg-neutral-800"
      />
      <Button
        className="self-start"
        disabled={busy || !markdown.trim()}
        onClick={() =>
          addSkill(
            async () => parseSkillMarkdown(markdown, "custom"),
            () => setMarkdown(""),
          )
        }
      >
        Add skill
      </Button>
      <div className="flex gap-1.5">
        <TextInput
          type="url"
          value={url}
          onChange={(event) => setUrl(event.target.value)}
          placeholder="SKILL.md or GitHub URL"
          aria-label="Skill URL"
          spellCheck={false}
        />
        <Button
          disabled={busy || !url.trim()}
          onClick={() =>
            addSkill(
              () => fetchSkillFromUrl(url.trim()),
              () => setUrl(""),
            )
          }
          className="h-8"
        >
          Import
        </Button>
      </div>
      {error && (
        <p role="alert" className="text-[12px] text-red-700 dark:text-red-400">
          {error}
        </p>
      )}
    </div>
  );
}
