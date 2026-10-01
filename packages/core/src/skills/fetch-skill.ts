import type { SkillDefinition } from "../contracts.ts";
import { parseSkillMarkdown } from "./parse-skill-markdown.ts";

const MAX_REPO_CANDIDATES = 30;

/**
 * Imports a SKILL.md from a raw URL, a GitHub blob or tree URL, or a skills.sh page
 * (skills.sh/{owner}/{repo}/{skill}, resolved through that GitHub repo).
 */
export async function fetchSkillFromUrl(url: string): Promise<SkillDefinition> {
  const parsed = new URL(url);
  const [owner, repo, kind, ref, ...path] = parsed.pathname.split("/").filter(Boolean);

  if (parsed.hostname === "skills.sh" && owner && repo && kind)
    return findSkillInRepo(owner, repo, kind);
  if (
    parsed.hostname === "github.com" &&
    owner &&
    repo &&
    ref &&
    (kind === "blob" || kind === "tree")
  ) {
    const filePath = kind === "tree" ? [...path, "SKILL.md"] : path;
    return fetchSkill(rawGitHubUrl(owner, repo, ref, filePath.join("/")));
  }
  return fetchSkill(url);
}

/** skills.sh names don't always match the directory, so fall back to matching each SKILL.md's front-matter name. */
async function findSkillInRepo(
  owner: string,
  repo: string,
  skillName: string,
): Promise<SkillDefinition> {
  const response = await fetch(
    `https://api.github.com/repos/${owner}/${repo}/git/trees/HEAD?recursive=1`,
  );
  if (!response.ok)
    throw new Error(`Couldn't list github.com/${owner}/${repo}: HTTP ${response.status}.`);
  // Network boundary: GitHub's documented git-tree response.
  const { tree = [] } = (await response.json()) as { tree?: { path: string }[] };
  const skillPaths = tree
    .map(({ path }) => path)
    .filter((path) => path === "SKILL.md" || path.endsWith("/SKILL.md"));

  const byDirectory = skillPaths.find((path) => path.split("/").at(-2) === skillName);
  if (byDirectory) return fetchSkill(rawGitHubUrl(owner, repo, "HEAD", byDirectory));

  // While scanning, a file that fails to fetch or parse simply isn't the match.
  const scanned = await Promise.all(
    skillPaths
      .slice(0, MAX_REPO_CANDIDATES)
      .map((path) => fetchSkill(rawGitHubUrl(owner, repo, "HEAD", path)).catch(() => undefined)),
  );
  const skill = scanned.find((candidate) => candidate?.name === skillName);
  if (!skill) throw new Error(`No skill named "${skillName}" in github.com/${owner}/${repo}.`);
  return skill;
}

function rawGitHubUrl(owner: string, repo: string, ref: string, path: string): string {
  return `https://raw.githubusercontent.com/${owner}/${repo}/${ref}/${path}`;
}

async function fetchSkill(url: string): Promise<SkillDefinition> {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`Couldn't fetch ${url}: HTTP ${response.status}.`);
  return parseSkillMarkdown(await response.text(), "custom");
}
