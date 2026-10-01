import { afterEach, describe, expect, test, vi } from "vitest";
import type { SkillDefinition } from "../../src/contracts.ts";
import {
  createSkillTool,
  fetchSkillFromUrl,
  formatSkillListing,
  parseSkillMarkdown,
} from "../../src/skills/index.ts";

const skill = (name: string, description: string): SkillDefinition => ({
  name,
  description,
  body: "",
  source: "bundled",
});

afterEach(() => vi.unstubAllGlobals());

describe("parseSkillMarkdown", () => {
  test("reads name, description, and body", () => {
    const parsed = parseSkillMarkdown(
      `---\nname: deck-design\ndescription: "Visual system: grid, type, color"\nlicense: MIT\n---\n\n# Deck design\nBody.\n`,
      "custom",
    );
    expect(parsed).toEqual({
      name: "deck-design",
      description: "Visual system: grid, type, color",
      body: "# Deck design\nBody.",
      source: "custom",
    });
  });

  test("folded block and continuation lines", () => {
    const folded = parseSkillMarkdown(
      `---\nname: a\ndescription: >\n  Use when\n  building decks.\n---\nBody`,
      "custom",
    );
    const continued = parseSkillMarkdown(
      `---\nname: a\ndescription: Use when\n  building decks.\n---\nBody`,
      "custom",
    );
    expect(folded.description).toBe("Use when building decks.");
    expect(continued.description).toBe("Use when building decks.");
  });

  test("rejects markdown without name and description", () => {
    expect(() => parseSkillMarkdown("# Just a doc", "custom")).toThrow(
      "must start with front matter",
    );
    expect(() => parseSkillMarkdown("---\nname: a\n---\nBody", "custom")).toThrow(
      "must start with front matter",
    );
  });
});

describe("formatSkillListing", () => {
  const skills = Array.from({ length: 10 }, (_, index) =>
    skill(`skill-${index}`, `Use for task ${index}. `.repeat(20)),
  );

  test("lists every skill in full when it fits", () => {
    const listing = formatSkillListing([skill("a", "Use for A."), skill("b", "Use for B.")]);
    expect(listing).toContain("load_skill");
    expect(listing).toContain("- a: Use for A.\n- b: Use for B.");
    expect(formatSkillListing([])).toBe("");
  });

  test("shortens descriptions to stay within budget, keeping every name", () => {
    const listing = formatSkillListing(skills, 1_500);
    expect(listing.length).toBeLessThanOrEqual(1_500);
    for (const { name } of skills) expect(listing).toContain(`- ${name}: Use for task`);
    expect(listing).toContain("…");
  });

  test("falls back to names, then a count, when the budget is tiny", () => {
    const namesOnly = formatSkillListing(skills, 250);
    expect(namesOnly.length).toBeLessThanOrEqual(250);
    expect(namesOnly).toContain("- skill-0\n");
    expect(namesOnly).toMatch(/- …and \d+ more$/);
  });
});

describe("load_skill", () => {
  const tool = createSkillTool(() => [
    { ...skill("charts", "Charts."), body: "Pick a chart.", files: { "types.md": "Bar vs line." } },
  ]);
  const text = async (args: object) => {
    const [part] = (await tool.execute("call", args)).content;
    return part?.type === "text" ? part.text : "";
  };

  test("returns the body and lists supporting files", async () => {
    expect(await text({ name: "charts" })).toBe(
      '# Skill: charts\n\nPick a chart.\n\nSupporting files (load_skill with name "charts" and file): types.md',
    );
    expect(await text({ name: "charts", file: "types.md" })).toBe(
      "# charts/types.md\n\nBar vs line.",
    );
  });

  test("unknown names and files list what exists", async () => {
    await expect(tool.execute("call", { name: "nope" })).rejects.toThrow("Available: charts.");
    await expect(tool.execute("call", { name: "charts", file: "x.md" })).rejects.toThrow(
      "Files: types.md.",
    );
  });
});

describe("fetchSkillFromUrl", () => {
  const skillMd = (name: string) => `---\nname: ${name}\ndescription: Does ${name}.\n---\nBody`;

  function stubFetch(routes: Record<string, string | object>) {
    const fetch = vi.fn(async (url: string) => {
      const body = routes[url];
      if (body === undefined) return new Response("missing", { status: 404 });
      return new Response(typeof body === "string" ? body : JSON.stringify(body));
    });
    vi.stubGlobal("fetch", fetch);
    return fetch;
  }

  test("GitHub blob and tree URLs read from raw.githubusercontent.com", async () => {
    stubFetch({
      "https://raw.githubusercontent.com/o/r/main/skills/pptx/SKILL.md": skillMd("pptx"),
    });
    const fromBlob = await fetchSkillFromUrl(
      "https://github.com/o/r/blob/main/skills/pptx/SKILL.md",
    );
    const fromTree = await fetchSkillFromUrl("https://github.com/o/r/tree/main/skills/pptx");
    expect(fromBlob).toMatchObject({ name: "pptx", source: "custom" });
    expect(fromTree.name).toBe("pptx");
  });

  test("skills.sh pages resolve by directory, else by front-matter name", async () => {
    stubFetch({
      "https://api.github.com/repos/o/r/git/trees/HEAD?recursive=1": {
        tree: [
          { path: "skills/charts/SKILL.md" },
          { path: "skills/best-practices/SKILL.md" },
          { path: "README.md" },
        ],
      },
      "https://raw.githubusercontent.com/o/r/HEAD/skills/charts/SKILL.md": skillMd("charts"),
      "https://raw.githubusercontent.com/o/r/HEAD/skills/best-practices/SKILL.md":
        skillMd("vendor-best-practices"),
    });
    expect((await fetchSkillFromUrl("https://skills.sh/o/r/charts")).name).toBe("charts");
    expect((await fetchSkillFromUrl("https://skills.sh/o/r/vendor-best-practices")).name).toBe(
      "vendor-best-practices",
    );
    await expect(fetchSkillFromUrl("https://skills.sh/o/r/absent")).rejects.toThrow(
      'No skill named "absent"',
    );
  });

  test("HTML or other non-skill responses are rejected", async () => {
    stubFetch({ "https://example.com/skill": "<!doctype html><p>Hi</p>" });
    await expect(fetchSkillFromUrl("https://example.com/skill")).rejects.toThrow(
      "Not a valid SKILL.md",
    );
  });
});
