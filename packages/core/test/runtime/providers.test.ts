import { afterEach, describe, expect, it, vi } from "vitest";
import { buildModel, isChatModel, toModelInfo } from "../../src/providers/modelCatalog.ts";
import { createProviderClient } from "../../src/providers/providerClient.ts";
import { defaultSettings } from "../../src/settings/settings.ts";

const endpoint = "https://proxy.test/v1/";
const info = (id: string, ownedBy?: string) =>
  toModelInfo(buildModel({ id, owned_by: ownedBy }, endpoint), ownedBy);

describe("model catalog", () => {
  it("maps owned_by to an API and a base URL the SDK can append to", () => {
    expect(buildModel({ id: "claude-sonnet-5-5", owned_by: "anthropic" }, endpoint)).toMatchObject({
      api: "anthropic-messages",
      baseUrl: "https://proxy.test",
    });
    expect(buildModel({ id: "gpt-6-luna", owned_by: "openai" }, endpoint)).toMatchObject({
      api: "openai-responses",
      baseUrl: "https://proxy.test/v1",
    });
    expect(buildModel({ id: "llama-4", owned_by: "meta" }, endpoint)).toMatchObject({
      api: "openai-completions",
      baseUrl: "https://proxy.test/v1",
    });
  });

  it("honours a per-model API override", () => {
    expect(
      buildModel({ id: "gpt-6-luna", owned_by: "openai" }, endpoint, "openai-completions").api,
    ).toBe("openai-completions");
  });

  it("gives each family its thinking levels, window, and image support", () => {
    expect(info("claude-sonnet-5-5", "anthropic")).toMatchObject({
      thinkingLevels: ["low", "medium", "high", "xhigh", "max"],
      contextWindow: 1_000_000,
      supportsImages: true,
    });
    expect(info("claude-opus-4-20250514", "anthropic")).toMatchObject({
      thinkingLevels: ["off", "minimal", "low", "medium", "high"],
      contextWindow: 200_000,
    });
    expect(info("claude-opus-7", "anthropic")).toMatchObject({
      thinkingLevels: ["off", "minimal", "low", "medium", "high", "xhigh", "max"],
      contextWindow: 1_000_000,
    });
    expect(info("gpt-6.9-nova", "openai")).toMatchObject({
      thinkingLevels: ["low", "medium", "high", "xhigh", "max"],
      contextWindow: 272_000,
    });
    expect(info("llama-4", "meta")).toMatchObject({
      thinkingLevels: ["off", "low", "medium", "high"],
      contextWindow: 128_000,
      supportsImages: false,
    });
  });

  it("drops image-generation models from the chat list", () => {
    expect(isChatModel({ id: "gpt-image-2" })).toBe(false);
    expect(isChatModel({ id: "gpt-6-sol" })).toBe(true);
  });
});

describe("provider client", () => {
  afterEach(() => vi.unstubAllGlobals());

  const settings = {
    ...defaultSettings,
    endpoint: { baseUrl: "https://proxy.test", apiKey: "secret" },
  };
  const listing = {
    data: [
      { id: "claude-sonnet-5-5", owned_by: "anthropic" },
      { id: "gpt-image-2", owned_by: "openai" },
    ],
  };

  it("lists chat models with both auth headers", async () => {
    const fetchMock = vi.fn(async () => Response.json(listing));
    vi.stubGlobal("fetch", fetchMock);

    const models = await createProviderClient(() => settings).listModels();

    expect(models.map((model) => model.id)).toEqual(["claude-sonnet-5-5"]);
    expect(fetchMock).toHaveBeenCalledWith("https://proxy.test/v1/models", {
      headers: { authorization: "Bearer secret", "x-api-key": "secret" },
    });
  });

  it("retries with x-api-key alone when the request with Authorization is blocked (CORS)", async () => {
    const fetchMock = vi.fn(async (_url: string, init: RequestInit) => {
      if (init.headers && "authorization" in init.headers) throw new TypeError("Failed to fetch");
      return Response.json(listing);
    });
    vi.stubGlobal("fetch", fetchMock);

    const models = await createProviderClient(() => settings).listModels();

    expect(models).toHaveLength(1);
    expect(fetchMock).toHaveBeenLastCalledWith("https://proxy.test/v1/models", {
      headers: { "x-api-key": "secret" },
    });
  });

  it("reports HTTP failures with the status and body", async () => {
    vi.stubGlobal("fetch", async () => new Response("bad key", { status: 401 }));

    await expect(createProviderClient(() => settings).listModels()).rejects.toThrow(
      "https://proxy.test/v1/models returned 401: bad key",
    );
  });
});
