import { afterEach, describe, expect, it } from "vitest";
import {
  codeBodySource,
  codeLineOffsets,
  runCodeBody,
  withSourceUrl,
  type CodeBody,
} from "../src/realm.ts";

const stubContext = {
  presentation: { slides: { load: () => ({ items: [{ id: "256#1" }] }) } },
  sync: async () => {},
};

afterEach(() => {
  delete (globalThis as { PowerPoint?: unknown }).PowerPoint;
  delete (globalThis as { OfficeExtension?: unknown }).OfficeExtension;
});

describe("runCodeBody error location", () => {
  it("quotes the model's line that threw", async () => {
    (globalThis as { OfficeExtension?: unknown }).OfficeExtension = { Error: class {} };
    (globalThis as { PowerPoint?: unknown }).PowerPoint = {
      run: (batch: (context: unknown) => Promise<unknown>) => batch(stubContext),
    };
    const code = "const shapes = [];\nconst title = shapes[0];\ntitle.left = 48;";
    const body = new Function(withSourceUrl(`return ${codeBodySource(code)}`))() as CodeBody;

    const run = await runCodeBody(body, { code, lineOffset: codeLineOffsets.newFunction });

    expect(run.ok).toBe(false);
    expect(run.error?.message).toContain("(line 3: title.left = 48;)");
    expect(run.slides).toEqual({ before: ["256#1"], after: ["256#1"] });
  });
});
