import { afterEach, describe, expect, test, vi } from "vitest";
import type { FootnoteTool, Settings, ToolEnv } from "../../src/contracts.ts";
import { createWebTools, fetchImageAsBase64 } from "../../src/web/index.ts";

afterEach(() => vi.unstubAllGlobals());

// The web tools only read env.settings().
const envWith = (firecrawlApiKey?: string) =>
  ({ settings: () => ({ firecrawlApiKey }) as Settings }) as ToolEnv;

function stubFetch(respond: (url: string, init?: RequestInit) => Response) {
  const fetch = vi.fn(async (url: string, init?: RequestInit) => respond(url, init));
  vi.stubGlobal("fetch", fetch);
  return fetch;
}

async function run(tool: FootnoteTool | undefined, args: object): Promise<string> {
  const [part] = (await tool!.execute("call", args)).content;
  return part?.type === "text" ? part.text : "";
}

describe("createWebTools", () => {
  test("no tools without a Firecrawl key", () => {
    expect(createWebTools(envWith(undefined))).toEqual([]);
  });

  test("web_search posts to /v2/search and lists results compactly", async () => {
    const fetch = stubFetch(() =>
      Response.json({
        success: true,
        data: {
          web: [{ title: "Q3 report", url: "https://a.example/q3", description: "Revenue grew." }],
        },
      }),
    );
    const [search] = createWebTools(envWith("fc-key"));
    expect(await run(search, { query: "acme q3" })).toBe(
      "1. Q3 report\n   https://a.example/q3\n   Revenue grew.",
    );
    const [url, init] = fetch.mock.calls[0]!;
    expect(url).toBe("https://api.firecrawl.dev/v2/search");
    expect(new Headers(init?.headers).get("authorization")).toBe("Bearer fc-key");
    expect(JSON.parse(String(init?.body))).toEqual({ query: "acme q3", limit: 5 });
  });

  test("fetch_page returns markdown pages with a continuation note", async () => {
    const markdown = "a".repeat(12_000) + "tail";
    stubFetch(() =>
      Response.json({
        success: true,
        data: {
          markdown,
          metadata: { title: "Doc", url: "https://a.example/doc", statusCode: 200 },
        },
      }),
    );
    const [, fetchPage] = createWebTools(envWith("fc-key"));
    const first = await run(fetchPage, { url: "https://a.example/doc" });
    expect(first.startsWith("# Doc\nSource: https://a.example/doc\n\naaa")).toBe(true);
    expect(first).toContain("Call fetch_page with offset 12000 for more.");
    expect(await run(fetchPage, { url: "https://a.example/doc", offset: 12_000 })).toContain(
      "\n\ntail",
    );
  });

  test("Firecrawl errors and failed pages surface as tool errors", async () => {
    stubFetch(() =>
      Response.json({ success: false, error: "Insufficient credits" }, { status: 402 }),
    );
    const [search] = createWebTools(envWith("fc-key"));
    await expect(run(search, { query: "x" })).rejects.toThrow(
      "Firecrawl /search failed (HTTP 402): Insufficient credits",
    );

    stubFetch(() =>
      Response.json({
        success: true,
        data: { markdown: "", metadata: { statusCode: 404, url: "https://a.example/x" } },
      }),
    );
    const [, fetchPage] = createWebTools(envWith("fc-key"));
    await expect(run(fetchPage, { url: "https://a.example/x" })).rejects.toThrow(
      "returned HTTP 404",
    );
  });
});

describe("fetchImageAsBase64", () => {
  test("returns the image's type and bytes", async () => {
    stubFetch(
      () =>
        new Response(new Uint8Array([1, 2, 3]), {
          headers: { "content-type": "image/png; charset=binary" },
        }),
    );
    expect(await fetchImageAsBase64("https://a.example/i.png")).toEqual({
      mimeType: "image/png",
      base64: "AQID",
    });
  });

  test("rejects non-images and oversized images", async () => {
    stubFetch(() => new Response("<html>", { headers: { "content-type": "text/html" } }));
    await expect(fetchImageAsBase64("https://a.example/page")).rejects.toThrow(
      'not an image (content type "text/html")',
    );

    stubFetch(
      () =>
        new Response("x", {
          headers: { "content-type": "image/jpeg", "content-length": String(21 * 1024 * 1024) },
        }),
    );
    await expect(fetchImageAsBase64("https://a.example/big.jpg")).rejects.toThrow(
      "larger than 20 MB",
    );
  });
});
