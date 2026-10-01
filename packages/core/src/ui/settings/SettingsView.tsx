import { ChevronRight } from "lucide-react";
import { useState, type ReactNode } from "react";
import type { ApiFormat, FootnoteApp, PermissionMode, Settings } from "../../contracts.ts";
import { clampThinkingLevel, thinkingLevelLabels, thinkingLevelOrder } from "../chat/transcript.ts";
import {
  Collapsible,
  CollapsibleContent,
  CollapsibleTrigger,
  disclosureChevron,
} from "../components/collapsible.tsx";
import { Input } from "../components/input.tsx";
import {
  Select,
  SelectContent,
  SelectItem,
  SelectTrigger,
  SelectValue,
} from "../components/select.tsx";
import type { ModelList } from "../hooks.ts";
import { EndpointSection } from "./EndpointSection.tsx";
import { Field } from "./Field.tsx";
import { SkillsSection } from "./SkillsSection.tsx";

const apiFormatLabels: Record<ApiFormat, string> = {
  "anthropic-messages": "Anthropic",
  "openai-responses": "Responses",
  "openai-completions": "Completions",
};

/** Select value for "no override"; Radix Select reserves the empty string for "nothing selected". */
const autoFormat = "auto";

const permissionModeLabels: Record<PermissionMode, string> = {
  ask: "Ask before edits",
  full: "Full access",
};

export function SettingsView({
  app,
  settings,
  modelList,
}: {
  app: FootnoteApp;
  settings: Settings;
  modelList: ModelList;
}) {
  const { models } = modelList;
  const defaultModel = models.find((model) => model.id === settings.modelId);
  const levels = defaultModel?.thinkingLevels ?? thinkingLevelOrder;

  function changeDefaultModel(modelId: string) {
    const supported =
      models.find((model) => model.id === modelId)?.thinkingLevels ?? thinkingLevelOrder;
    const thinkingLevel =
      clampThinkingLevel(settings.thinkingLevel, supported) ?? settings.thinkingLevel;
    void app.settings.update({ modelId, thinkingLevel });
  }

  async function changeApiOverride(modelId: string, format: ApiFormat | undefined) {
    const { [modelId]: _previous, ...rest } = settings.apiOverrides;
    await app.settings.update({ apiOverrides: format ? { ...rest, [modelId]: format } : rest });
    await modelList.reload();
  }

  return (
    <div className="min-h-0 flex-1 overflow-y-auto">
      <div className="flex flex-col gap-6 p-3">
        <Section title="Connection">
          <EndpointSection app={app} settings={settings} modelList={modelList} />
          <SecretField
            label="Firecrawl API key"
            value={settings.firecrawlApiKey ?? ""}
            onSave={(firecrawlApiKey) =>
              app.settings.update({ firecrawlApiKey: firecrawlApiKey || undefined })
            }
          />
        </Section>

        <Section title="Defaults">
          <Field label="Permissions">
            {(id) => (
              <Select
                value={settings.defaultPermissionMode}
                onValueChange={(mode) =>
                  app.settings.update({ defaultPermissionMode: mode === "full" ? "full" : "ask" })
                }
              >
                <SelectTrigger id={id}>
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  {Object.entries(permissionModeLabels).map(([mode, label]) => (
                    <SelectItem key={mode} value={mode}>
                      {label}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            )}
          </Field>
          <Field label="Model">
            {(id) => (
              <Select value={settings.modelId ?? ""} onValueChange={changeDefaultModel}>
                <SelectTrigger id={id}>
                  <SelectValue placeholder="Choose a model" />
                </SelectTrigger>
                <SelectContent>
                  {settings.modelId && !defaultModel && (
                    <SelectItem value={settings.modelId}>{settings.modelId}</SelectItem>
                  )}
                  {models.map((model) => (
                    <SelectItem key={model.id} value={model.id}>
                      {model.id}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            )}
          </Field>
          <Field label="Thinking effort">
            {(id) => (
              <Select
                value={settings.thinkingLevel}
                onValueChange={(value) => {
                  const thinkingLevel = levels.find((level) => level === value);
                  if (thinkingLevel) void app.settings.update({ thinkingLevel });
                }}
              >
                <SelectTrigger id={id}>
                  <SelectValue />
                </SelectTrigger>
                <SelectContent>
                  {levels.map((level) => (
                    <SelectItem key={level} value={level}>
                      {thinkingLevelLabels[level]}
                    </SelectItem>
                  ))}
                </SelectContent>
              </Select>
            )}
          </Field>
        </Section>

        {models.length > 0 && (
          <Section title="API format">
            <Collapsible>
              <CollapsibleTrigger className="group flex cursor-default items-center gap-1 rounded-sm text-small text-muted-foreground select-none hover:text-foreground">
                <ChevronRight size={13} className={disclosureChevron} />
                {Object.keys(settings.apiOverrides).length} overridden of {models.length}
              </CollapsibleTrigger>
              <CollapsibleContent>
                <ul className="flex flex-col gap-1 pt-2">
                  {models.map((model) => (
                    <li key={model.id} className="flex items-center gap-2">
                      <span
                        className="min-w-0 flex-1 truncate font-mono text-small"
                        title={model.id}
                      >
                        {model.id}
                      </span>
                      <Select
                        value={settings.apiOverrides[model.id] ?? autoFormat}
                        onValueChange={(value) =>
                          changeApiOverride(model.id, parseApiFormat(value))
                        }
                      >
                        <SelectTrigger
                          aria-label={`API format for ${model.id}`}
                          className="w-44 shrink-0 text-small"
                        >
                          <SelectValue />
                        </SelectTrigger>
                        <SelectContent>
                          <SelectItem value={autoFormat}>
                            {settings.apiOverrides[model.id]
                              ? "Auto"
                              : `Auto · ${apiFormatLabels[model.api]}`}
                          </SelectItem>
                          {Object.entries(apiFormatLabels).map(([format, label]) => (
                            <SelectItem key={format} value={format}>
                              {label}
                            </SelectItem>
                          ))}
                        </SelectContent>
                      </Select>
                    </li>
                  ))}
                </ul>
              </CollapsibleContent>
            </Collapsible>
          </Section>
        )}

        <Section title="Skills">
          <SkillsSection app={app} settings={settings} />
        </Section>
      </div>
    </div>
  );
}

function parseApiFormat(value: string): ApiFormat | undefined {
  return Object.keys(apiFormatLabels).find((format): format is ApiFormat => format === value);
}

function Section({ title, children }: { title: string; children: ReactNode }) {
  return (
    <section className="flex flex-col gap-2.5">
      <h2 className="text-caption font-semibold tracking-wide text-subtle-foreground uppercase">
        {title}
      </h2>
      {children}
    </section>
  );
}

function SecretField({
  label,
  value,
  onSave,
}: {
  label: string;
  value: string;
  onSave: (value: string) => Promise<void>;
}) {
  const [draft, setDraft] = useState(value);
  return (
    <Field label={label}>
      {(id) => (
        <Input
          id={id}
          type="password"
          value={draft}
          onChange={(event) => setDraft(event.target.value)}
          onBlur={() => draft.trim() !== value && void onSave(draft.trim())}
          autoComplete="off"
          spellCheck={false}
        />
      )}
    </Field>
  );
}
