import { Eye, EyeOff } from "lucide-react";
import { useState } from "react";
import type { FootnoteApp, Settings } from "../../contracts.ts";
import { Button, IconButton, TextInput } from "../components/controls.tsx";
import type { ModelList } from "../hooks.ts";
import { Field } from "./Field.tsx";

export function EndpointSection({
  app,
  settings,
  modelList,
}: {
  app: FootnoteApp;
  settings: Settings;
  modelList: ModelList;
}) {
  const [baseUrl, setBaseUrl] = useState(settings.endpoint.baseUrl);
  const [apiKey, setApiKey] = useState(settings.endpoint.apiKey);
  const [showKey, setShowKey] = useState(false);
  const [tested, setTested] = useState(false);

  const save = () =>
    app.settings.update({ endpoint: { baseUrl: baseUrl.trim(), apiKey: apiKey.trim() } });

  async function test() {
    await save();
    setTested(true);
    await modelList.reload();
  }

  return (
    <div className="flex flex-col gap-2.5">
      <Field label="Endpoint URL">
        {(id) => (
          <TextInput
            id={id}
            type="url"
            value={baseUrl}
            onChange={(event) => setBaseUrl(event.target.value)}
            onBlur={save}
            placeholder="https://example.com"
            spellCheck={false}
          />
        )}
      </Field>
      <Field label="API key">
        {(id) => (
          <div className="flex gap-1">
            <TextInput
              id={id}
              type={showKey ? "text" : "password"}
              value={apiKey}
              onChange={(event) => setApiKey(event.target.value)}
              onBlur={save}
              autoComplete="off"
              spellCheck={false}
            />
            <IconButton
              icon={showKey ? EyeOff : Eye}
              label={showKey ? "Hide key" : "Show key"}
              onClick={() => setShowKey(!showKey)}
              className="size-8"
            />
          </div>
        )}
      </Field>
      <div className="flex items-center gap-2">
        <Button onClick={test} disabled={!baseUrl.trim() || modelList.loading}>
          Test connection
        </Button>
        <ConnectionResult
          modelList={modelList}
          visible={tested || settings.endpoint.baseUrl !== ""}
        />
      </div>
    </div>
  );
}

function ConnectionResult({ modelList, visible }: { modelList: ModelList; visible: boolean }) {
  if (!visible) return null;
  if (modelList.loading) return <span className="text-[12px] text-neutral-500">Connecting…</span>;
  if (modelList.error)
    return (
      <span role="alert" className="min-w-0 text-[12px] break-words text-red-700 dark:text-red-400">
        {modelList.error}
      </span>
    );
  const count = modelList.models.length;
  return (
    <span role="status" className="text-[12px] text-accent-fg">
      {count === 1 ? "1 model" : `${count} models`}
    </span>
  );
}
