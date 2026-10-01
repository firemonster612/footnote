# Footnote

Model-agnostic assistant that edits Office documents live. PowerPoint first; Word/Excel later. Design: `docs/design.md`. Shared types: `packages/core/src/contracts.ts` (every package builds against these; don't change them without the orchestrator).

## Layout

- `packages/core`: agent runtime (Pi `@earendil-works/pi-agent-core` 1.0), providers, settings, permissions, context, compaction, attachments, web tools, skills registry, React UI (`src/ui`).
- `packages/powerpoint`: HostModule for PowerPoint. `src/ops` run inside the Office realm (Office.js); `src/tools` run with the agent and call ops via `OfficeHost.call`.
- `apps/extension`: WXT Chrome MV3 extension. Side panel UI + bridge into the Claude add-in frame (`pivot.claude.ai`).
- `apps/addin`: Vite task-pane add-in with XML manifest.

## Commands

- Install: `bun install` (bun only; never npm/yarn)
- Typecheck a package: `cd <pkg> && bunx tsc -p .`
- Tests: `bunx vitest run` (from repo root or a package)
- Lint/format: `bunx oxlint`, `bunx oxfmt`

## Rules

- TypeScript strict, ESM, `.ts` import extensions inside packages.
- Ops (Office realm) take and return structured-clone-safe values. No closures, DOM, or chrome.* in ops.
- No native browser dialogs (alert/confirm/prompt) anywhere.
- Tailwind v4 + small local components; no heavy UI kits.
