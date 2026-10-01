import { useState, type ReactNode } from "react";
import type { ApiFormat, FootnoteApp, PermissionMode, Settings } from "../../contracts.ts";
import { clampThinkingLevel, thinkingLevelLabels, thinkingLevelOrder } from "../chat/transcript.ts";
import { Select, TextInput } from "../components/controls.tsx";
import type { ModelList } from "../hooks.ts";
import { EndpointSection } from "./EndpointSection.tsx";
import { Field } from "./Field.tsx";
import { SkillsSection } from "./SkillsSection.tsx";

const apiFormatLabels: Record<ApiFormat, string> = {
  "anthropic-messages": "Anthropic Messages",
  "openai-responses": "OpenAI Responses",
  "openai-completions": "Chat Completions",
};

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
      <div className="flex flex-col gap-5 p-3">
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
                id={id}
                variant="field"
                value={settings.defaultPermissionMode}
                onChange={(event) =>
                  app.settings.update({
                    defaultPermissionMode: event.target.value === "full" ? "full" : "ask",
                  })
                }
              >
                {Object.entries(permissionModeLabels).map(([mode, label]) => (
                  <option key={mode} value={mode}>
                    {label}
                  </option>
                ))}
              </Select>
            )}
          </Field>
          <Field label="Model">
            {(id) => (
              <Select
                id={id}
                variant="field"
                value={settings.modelId ?? ""}
                onChange={(event) => changeDefaultModel(event.target.value)}
              >
                {!defaultModel && (
                  <option value={settings.modelId ?? ""}>
                    {settings.modelId ?? "Choose a model"}
                  </option>
                )}
                {models.map((model) => (
                  <option key={model.id} value={model.id}>
                    {model.id}
                  </option>
                ))}
              </Select>
            )}
          </Field>
          <Field label="Thinking effort">
            {(id) => (
              <Select
                id={id}
                variant="field"
                value={settings.thinkingLevel}
                onChange={(event) => {
                  const thinkingLevel = levels.find((level) => level === event.target.value);
                  if (thinkingLevel) void app.settings.update({ thinkingLevel });
                }}
              >
                {levels.map((level) => (
                  <option key={level} value={level}>
                    {thinkingLevelLabels[level]}
                  </option>
                ))}
              </Select>
            )}
          </Field>
        </Section>

        {models.length > 0 && (
          <Section title="API format">
            <details className="group">
              <summary className="cursor-pointer text-[12px] text-neutral-600 select-none dark:text-neutral-400">
                {Object.keys(settings.apiOverrides).length} overridden of {models.length}
              </summary>
              <ul className="mt-2 flex flex-col gap-1">
                {models.map((model) => (
                  <li key={model.id} className="flex items-center gap-2">
                    <span
                      className="min-w-0 flex-1 truncate font-mono text-[12px]"
                      title={model.id}
                    >
                      {model.id}
                    </span>
                    <Select
                      variant="field"
                      aria-label={`API format for ${model.id}`}
                      value={settings.apiOverrides[model.id] ?? ""}
                      onChange={(event) =>
                        changeApiOverride(model.id, parseApiFormat(event.target.value))
                      }
                      className="w-40 shrink-0"
                    >
                      <option value="">
                        {settings.apiOverrides[model.id]
                          ? "Auto"
                          : `Auto (${apiFormatLabels[model.api]})`}
                      </option>
                      {Object.entries(apiFormatLabels).map(([format, label]) => (
                        <option key={format} value={format}>
                          {label}
                        </option>
                      ))}
                    </Select>
                  </li>
                ))}
              </ul>
            </details>
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
      <h2 className="text-[11px] font-semibold tracking-wide text-neutral-500 uppercase">
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
        <TextInput
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
