import {
  getSupportedThinkingLevels,
  type Api,
  type Model,
  type ThinkingLevelMap,
} from "@earendil-works/pi-ai";
import { ANTHROPIC_MODELS } from "@earendil-works/pi-ai/providers/anthropic.models";
import { OPENAI_MODELS } from "@earendil-works/pi-ai/providers/openai.models";
import type { ApiFormat, ModelInfo } from "../contracts.ts";

/** One entry of an OpenAI-style `/v1/models` listing. */
export interface ListedModel {
  id: string;
  owned_by?: string;
}

const piCatalog: readonly Model<Api>[] = [
  ...Object.values(ANTHROPIC_MODELS),
  ...Object.values(OPENAI_MODELS),
];

const nonChatModelPattern = /image|embedding|tts|whisper|transcribe|dall-e|moderation/;

export function isChatModel(listed: ListedModel): boolean {
  return !nonChatModelPattern.test(listed.id);
}

export function inferApi(ownedBy: string | undefined): ApiFormat {
  if (ownedBy === "anthropic") return "anthropic-messages";
  if (ownedBy === "openai") return "openai-responses";
  return "openai-completions";
}

/** The endpoint root without a trailing slash or `/v1`, so users can paste either form. */
export function endpointRoot(baseUrl: string): string {
  return baseUrl.trim().replace(/\/+$/, "").replace(/\/v1$/, "");
}

/** The Anthropic SDK appends `/v1/messages` itself; the OpenAI SDK appends only `/responses` or `/chat/completions`. */
function apiBaseUrl(endpointBaseUrl: string, api: ApiFormat): string {
  const root = endpointRoot(endpointBaseUrl);
  return api === "anthropic-messages" ? root : `${root}/v1`;
}

/**
 * A Pi model for a listed ID at the configured endpoint. Pi's built-in catalog supplies exact capabilities for known
 * models; unknown IDs get family defaults.
 */
export function buildModel(
  listed: ListedModel,
  endpointBaseUrl: string,
  apiOverride?: ApiFormat,
): Model<Api> {
  const api = apiOverride ?? inferApi(listed.owned_by);
  const known = piCatalog.find((model) => model.id === listed.id && model.api === api);
  const spec = known ? withoutServerFallbacks(known) : familyModel(listed.id, api);
  return { ...spec, baseUrl: apiBaseUrl(endpointBaseUrl, api) };
}

export function toModelInfo(model: Model<Api>, ownedBy: string | undefined): ModelInfo {
  return {
    id: model.id,
    ...(ownedBy !== undefined && { ownedBy }),
    // Every model built here has an ApiFormat api (see buildModel).
    api: model.api as ApiFormat,
    thinkingLevels: getSupportedThinkingLevels(model),
    contextWindow: model.contextWindow,
    supportsImages: model.input.includes("image"),
  };
}

/** Anthropic refusal fallbacks name upstream models a proxy may not route; never send them. */
function withoutServerFallbacks(model: Model<Api>): Model<Api> {
  if (!model.compat || !("allowedFallbackModels" in model.compat)) return model;
  const { allowedFallbackModels: _dropped, ...compat } = model.compat;
  return { ...model, compat };
}

const zeroCost = { input: 0, output: 0, cacheRead: 0, cacheWrite: 0 };

const providerByApi: Record<ApiFormat, string> = {
  "anthropic-messages": "anthropic",
  "openai-responses": "openai",
  "openai-completions": "footnote",
};

function familyModel(id: string, api: ApiFormat): Model<Api> {
  const unknown: Model<Api> = {
    id,
    name: id,
    api,
    provider: providerByApi[api],
    baseUrl: "",
    cost: zeroCost,
    input: ["text"],
    reasoning: true,
    thinkingLevelMap: { minimal: null },
    contextWindow: 128_000,
    maxTokens: 16_384,
    ...(api === "openai-completions" && { compat: { supportsReasoningEffort: true } }),
  };
  return { ...unknown, ...(claudeFamily(id) ?? gptFamily(id)) };
}

type FamilyTraits = Partial<Model<Api>>;

/** Parses "claude-opus-4-6", "claude-3-5-haiku-20241022", "claude-fable-5" into a numeric version (4.6, 3.5, 5). */
function claudeVersion(id: string): number | undefined {
  const match = /^claude-(?:[a-z]+-)?(\d+)(?:[-.](\d)(?!\d))?/.exec(id);
  if (!match) return undefined;
  return Number(`${match[1]}.${match[2] ?? 0}`);
}

function claudeFamily(id: string): FamilyTraits | undefined {
  const version = claudeVersion(id);
  if (version === undefined) return undefined;
  const claude: FamilyTraits = {
    input: ["text", "image"],
    contextWindow: 200_000,
    maxTokens: 64_000,
  };
  if (version < 3.7) return { ...claude, reasoning: false, maxTokens: 8_192 };
  // Budget-based extended thinking.
  if (version < 4.6) return { ...claude, thinkingLevelMap: {} };
  const levels: ThinkingLevelMap = version >= 4.7 ? { xhigh: "xhigh", max: "max" } : { max: "max" };
  return {
    ...claude,
    contextWindow: 1_000_000,
    maxTokens: 128_000,
    thinkingLevelMap: levels,
    compat: { forceAdaptiveThinking: true },
  };
}

function gptFamily(id: string): FamilyTraits | undefined {
  if (/^o\d/.test(id)) {
    return {
      input: ["text", "image"],
      contextWindow: 200_000,
      maxTokens: 100_000,
      thinkingLevelMap: { off: null, minimal: null },
    };
  }
  const match = /^gpt-(\d+)(?:\.(\d+))?/.exec(id);
  if (!match) return undefined;
  const version = Number(`${match[1]}.${match[2] ?? 0}`);
  if (version < 5) return { input: ["text", "image"], reasoning: false };
  return {
    input: ["text", "image"],
    contextWindow: 272_000,
    maxTokens: 128_000,
    thinkingLevelMap: {
      off: null,
      minimal: null,
      xhigh: "xhigh",
      max: version >= 5.6 ? "max" : null,
    },
  };
}
