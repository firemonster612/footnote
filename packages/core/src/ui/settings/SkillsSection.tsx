import { Trash } from "lucide-react";
import { useState } from "react";
import type { FootnoteApp, Settings, SkillDefinition } from "../../contracts.ts";
import { fetchSkillFromUrl, parseSkillMarkdown } from "../../skills/index.ts";
import { Button } from "../components/button.tsx";
import { IconButton } from "../components/icon-button.tsx";
import { Input } from "../components/input.tsx";
import { Textarea } from "../components/textarea.tsx";
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
        <ul className="flex flex-col divide-y rounded-md border">
          {skills.map((skill) => (
            <li key={skill.name} className="flex items-start gap-2 py-1.5 pr-1 pl-2">
              <div className="min-w-0 flex-1">
                <div className="truncate font-medium">{skill.name}</div>
                <div className="line-clamp-2 text-small text-muted-foreground">
                  {skill.description}
                </div>
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
      <Textarea
        value={markdown}
        onChange={(event) => setMarkdown(event.target.value)}
        placeholder={"---\nname: my-skill\ndescription: When to use it\n---\n\nInstructions"}
        aria-label="SKILL.md content"
        rows={4}
        spellCheck={false}
        className="font-mono text-small"
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
        <Input
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
          size="md"
        >
          Import
        </Button>
      </div>
      {error && (
        <p role="alert" className="text-small text-danger">
          {error}
        </p>
      )}
    </div>
  );
}
