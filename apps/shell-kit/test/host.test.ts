import { OfficeOpError } from "@footnote/core/contracts";
import { afterEach, describe, expect, it, vi } from "vitest";
import { statusFromInfo, unwrapOpResponse, withDeadline } from "../src/host.ts";

describe("statusFromInfo", () => {
  it("connects to the expected host and names the document from its URL", () => {
    const status = statusFromInfo(
      {
        hostName: "PowerPoint",
        documentUrl:
          "https://contoso.sharepoint.com/sites/team/Shared%20Documents/Q3%20Review.pptx?web=1",
        apiVersions: { PowerPointApi: "1.8" },
      },
      "powerpoint",
      "https://tab.example/",
    );
    expect(status).toEqual({
      connected: true,
      host: "powerpoint",
      apiVersions: { PowerPointApi: "1.8" },
      documentId:
        "https://contoso.sharepoint.com/sites/team/Shared%20Documents/Q3%20Review.pptx?web=1",
      documentName: "Q3 Review.pptx",
    });
  });

  it("falls back to the given document ID when Office has no URL", () => {
    const status = statusFromInfo(
      { hostName: "PowerPoint", apiVersions: {} },
      "powerpoint",
      "https://tab.example/",
    );
    expect(status.documentId).toBe("https://tab.example/");
    expect(status.documentName).toBeUndefined();
  });

  it("refuses a different Office host", () => {
    expect(statusFromInfo({ hostName: "Word", apiVersions: {} }, "powerpoint")).toEqual({
      connected: false,
      host: "word",
      reason: "Footnote works with PowerPoint, but this add-in is open in Word.",
    });
  });
});

describe("withDeadline", () => {
  afterEach(() => vi.useRealTimers());

  it("rejects with a TimeoutError when the work outlives the timeout", async () => {
    vi.useFakeTimers();
    const result = withDeadline(new Promise(() => {}), 'Op "get_deck"', { timeoutMs: 2_000 });
    vi.advanceTimersByTime(2_000);
    await expect(result).rejects.toMatchObject({
      name: "TimeoutError",
      message: 'Op "get_deck" timed out after 2s',
    });
  });

  it("rejects with an OutcomeUnknownError when stopped, since the work was already sent", async () => {
    const controller = new AbortController();
    const result = withDeadline(new Promise(() => {}), 'Op "add_slide"', {
      signal: controller.signal,
    });
    controller.abort();
    await expect(result).rejects.toMatchObject({
      name: "OutcomeUnknownError",
      message: 'Op "add_slide" was stopped before it answered',
    });
  });
});

it("unwrapOpResponse turns a failed op into an OfficeOpError", () => {
  expect(() =>
    unwrapOpResponse("get_slide", {
      ok: false,
      error: { message: "ItemNotFound", code: "ItemNotFound" },
    }),
  ).toThrow(new OfficeOpError("get_slide", { message: "ItemNotFound", code: "ItemNotFound" }));
});
