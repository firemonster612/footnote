import type { StreamFn } from "@earendil-works/pi-agent-core";
import type { Api, Model, ProviderHeaders, ProviderStreams } from "@earendil-works/pi-ai";
import { anthropicMessagesApi } from "@earendil-works/pi-ai/api/anthropic-messages.lazy";
import { openAICompletionsApi } from "@earendil-works/pi-ai/api/openai-completions.lazy";
import { openAIResponsesApi } from "@earendil-works/pi-ai/api/openai-responses.lazy";
import type { ApiFormat, ModelInfo, Settings } from "../contracts.ts";
import {
  buildModel,
  endpointRoot,
  isChatModel,
  toModelInfo,
  type ListedModel,
} from "./modelCatalog.ts";

export interface ProviderClient {
  /** Fetches `/v1/models`. Throws with a user-readable message on failure. */
  listModels(): Promise<ModelInfo[]>;
  /** The Pi model for an ID offered by the endpoint, listing models first if needed. */
  resolveModel(modelId: string): Promise<Model<Api>>;
  /** Streams any model from resolveModel with the endpoint's credentials. */
  streamFn: StreamFn;
}

/**
 * "both": Authorization + x-api-key. "x-api-key": only x-api-key, for pages bound by CORS, where a wildcard
 * Access-Control-Allow-Headers (CLIProxyAPI's) doesn't cover Authorization.
 */
type AuthStyle = "both" | "x-api-key";

function authHeaders(style: AuthStyle, apiKey: string): ProviderHeaders {
  if (!apiKey) return {};
  return style === "both"
    ? { authorization: `Bearer ${apiKey}`, "x-api-key": apiKey }
    : { authorization: null, "x-api-key": apiKey };
}

function headersForFetch(headers: ProviderHeaders): Record<string, string> {
  return Object.fromEntries(
    Object.entries(headers).filter((entry): entry is [string, string] => entry[1] !== null),
  );
}

const apiStreams: Record<ApiFormat, ProviderStreams> = {
  "anthropic-messages": anthropicMessagesApi(),
  "openai-responses": openAIResponsesApi(),
  "openai-completions": openAICompletionsApi(),
};

export function createProviderClient(getSettings: () => Settings): ProviderClient {
  let authStyle: AuthStyle = "both";
  let listing: { endpointKey: string; models: ListedModel[] } | undefined;

  const endpointKey = ({ endpoint }: Settings) =>
    `${endpointRoot(endpoint.baseUrl)}\n${endpoint.apiKey}`;

  async function fetchListing(): Promise<ListedModel[]> {
    const settings = getSettings();
    const { baseUrl, apiKey } = settings.endpoint;
    if (!baseUrl.trim()) throw new Error("Set the endpoint URL in Settings.");
    const url = `${endpointRoot(baseUrl)}/v1/models`;
    const request = (style: AuthStyle) =>
      fetch(url, { headers: headersForFetch(authHeaders(style, apiKey)) });

    let response: Response;
    try {
      response = await request("both");
      authStyle = "both";
    } catch {
      // A rejected fetch here is a network or CORS failure; retry without Authorization before giving up.
      response = await request("x-api-key").catch((error: unknown) => {
        throw new Error(
          `Couldn't reach ${url}: ${error instanceof Error ? error.message : String(error)}`,
          {
            cause: error,
          },
        );
      });
      authStyle = "x-api-key";
    }
    if (!response.ok) {
      const detail = (await response.text()).slice(0, 300);
      throw new Error(`${url} returned ${response.status}${detail ? `: ${detail}` : ""}`);
    }
    const body: { data?: ListedModel[] } = await response.json();
    const models = (body.data ?? []).filter(isChatModel);
    listing = { endpointKey: endpointKey(settings), models };
    return models;
  }

  async function currentListing(): Promise<ListedModel[]> {
    if (listing?.endpointKey === endpointKey(getSettings())) return listing.models;
    return fetchListing();
  }

  return {
    async listModels() {
      const settings = getSettings();
      const models = await fetchListing();
      return models.map((listed) =>
        toModelInfo(
          buildModel(listed, settings.endpoint.baseUrl, settings.apiOverrides[listed.id]),
          listed.owned_by,
        ),
      );
    },

    async resolveModel(modelId) {
      const listed = (await currentListing()).find((model) => model.id === modelId);
      const settings = getSettings();
      if (!listed)
        throw new Error(
          `${modelId} isn't offered by ${settings.endpoint.baseUrl}. Pick another model.`,
        );
      return buildModel(listed, settings.endpoint.baseUrl, settings.apiOverrides[modelId]);
    },

    streamFn(model, context, options) {
      const { apiKey } = getSettings().endpoint;
      // resolveModel only builds models whose api is an ApiFormat.
      return apiStreams[model.api as ApiFormat].streamSimple(model, context, {
        ...options,
        apiKey,
        headers: { ...options?.headers, ...authHeaders(authStyle, apiKey) },
      });
    },
  };
}
