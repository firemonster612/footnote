# Footnote design

Footnote is a model-agnostic assistant that edits Office documents live. PowerPoint first; Word and Excel reuse everything except the host tools. Research behind these decisions lives in `.scratch/research/`.

## Shells

One codebase, two thin loaders. Tool code is pure Office.js and runs wherever Office.js lives.

| Shell | Where the UI runs | Where Office.js runs | Use |
|---|---|---|---|
| Extension (Chrome MV3) | Chrome side panel (agent in an offscreen document) | The org-deployed Claude add-in frame (`pivot.claude.ai`), reached by a packaged MAIN-world script | Work tenant without IT |
| Add-in (XML manifest) | Office task pane | Same page | Personal account, future IT deployment |

The UI and agent loop never touch Office objects directly. They call an `OfficeHost` RPC (`call(op, args) -> result`). The add-in shell implements it in-page; the extension shell forwards it to the bridge script in the add-in frame. The bridge only runs packaged code, so the frame's CSP doesn't matter except for raw code execution (see below).

Model requests go straight from the extension's engine or the task pane to the configured endpoint. Extension pages with host permissions are not bound by CORS; the add-in sends `x-api-key` because CLIProxyAPI's wildcard CORS header doesn't cover `Authorization`.

## Extension runtime

The agent runs in an offscreen document (the engine, `entrypoints/engine`), not in the side panel, so a task keeps going when the panel closes. The service worker alone would be a poor host: Chrome stops it when idle, including during long model streams. The service worker creates the engine on demand (reason `WORKERS`); the engine closes itself after 10 minutes with no panel open and no chat running, so quick tab switches keep undo checkpoints.

- **Panel = view.** The side panel renders `FootnoteRoot` against `connectFootnoteApp` (`packages/core/src/remote`), a `FootnoteApp` mirror over a `chrome.runtime` port: the engine pushes settings, host status and session snapshots (transcripts as a tail patch), and the view's calls run as RPC. Files cross as base64. Approvals live in the engine, so they wait while the panel is closed.
- **Per-tab panel.** The panel is disabled globally and enabled for the tab whose toolbar button was clicked (`sidePanel.setOptions({ tabId, path: "sidepanel.html?tab=<id>" })`), so Chrome hides it on other tabs and shows it again on that one. The toolbar badge shows "…" while that tab's chat runs with the panel hidden and "!" while an approval waits.
- **Office access.** Each tab gets its own `OfficeHost` in the engine, and a chat stays on the tab it was running on whichever tab is active. The engine only has `chrome.runtime`, so the service worker does frame discovery (`scripting`), code runs (`userScripts.execute`), settings storage and the badge, each as a one-shot message that survives worker restarts. Op calls skip the worker: the relay in the add-in frame opens a port straight to the engine when asked.

## Packages

```
packages/core        agent loop (Pi), providers, settings, permissions, context assembly, compaction, chat UI
packages/powerpoint  PowerPoint ops (run in the Office realm), tool schemas, deck-state summary, skills
apps/extension       WXT: engine (offscreen), side panel view, service worker, bridge content scripts
apps/addin           Vite: task pane page, manifest.xml
```

Word and Excel become `packages/word` and `packages/excel` with the same shape.

## Models

- One endpoint: URL + API key (CLIProxyAPI or any OpenAI-compatible server). Optional Firecrawl key.
- Models come from `/v1/models`. API format per model: `owned_by: anthropic` → Anthropic Messages, `openai` → OpenAI Responses, anything else → Chat Completions. Manual override per model.
- Reasoning effort is a request parameter (Pi thinking levels → `output_config.effort`, `reasoning.effort`, or `reasoning_effort`), never a model-name suffix or prompt text. Levels shown per model family.

## Tools

Hybrid, following Claude Code, Codex, and Claude in Excel: structured tools for everyday work, code for everything else.

**Structured tools** (first cut, pruned by observed use):
- Read: `get_deck` (outline, theme, layouts), `get_slide` (shape tree with IDs, bounds, text, styles), `get_selection`, `render_slide` (image), `get_notes`
- Write: `add_slide`, `delete_slides`, `move_slide`, `duplicate_slide`, `apply_layout`, `add_shape` (text box / geometric / line / image), `update_shapes` (batch: text, font, fill, line, bounds, z-order, delete), `edit_table` (create / cells / format / rows & columns), `group_shapes`, `set_notes`, `insert_chart` (PptxGenJS), `insert_slides_from_file`
- Other: `execute_office_js`, `load_skill`, `web_search`, `fetch_page`, `read_attachment`

Every write reads back the affected objects and returns a compact receipt: changed IDs, what was verified, and any partial failure. A write against a slide whose fingerprint changed since the model last read it fails with "re-read slide N", the Office equivalent of Claude Code's read-before-edit check, so the agent can't silently overwrite edits you made mid-turn.

**Code mode** (`execute_office_js`): the model writes an async function body that receives a live `PowerPoint.RequestContext` and a `footnote` helper library (shape lookup by ID, unit conversion, bulk text helpers, read-back). It runs in the Office realm: `new Function` in the add-in shell, `chrome.userScripts.execute` in the extension shell (needs Chrome's "Allow user scripts" toggle; falls back to an error with setup instructions). Limits: 20k characters of code, 8k characters of result, timeout reported as "outcome unknown, re-read before retrying". Our API keys never live in that realm.

A QuickJS sandbox with a `footnote.*` RPC API (Codex/Cloudflare style) is the upgrade path if raw code proves risky; it isn't in v1.

## Permissions

Two modes, chosen per chat; the default is a setting.
- **Ask**: reads, renders, web search and skills run freely. Every document write asks (with "allow writes for this chat"). `execute_office_js` asks every time and shows the code.
- **Full access**: nothing asks.

Denials go back to the model as a tool result with your comment. The model is told not to retry a denied call unchanged.

## Context assembly

- **Static, cached**: system prompt (role, Office editing rules, tool guidance), tool schemas, skill listing (names + one-line descriptions, budgeted).
- **Appended each turn**, as a `<deck_state>` block after cached history: deck identity, slide count and order (ID, title, layout), selection with selected-slide shape summary, theme fonts and colors on the first turn and when they change, and edits made since the last turn, including your own manual edits (detected by fingerprint diff).
- **On demand via tools**: full shape trees, notes, other slides, renders, attachments.
- Deck text is treated as untrusted data in the prompt.

## Budgets and compaction

- Tool results: ~4k-token budget each, with explicit "truncated, N more shapes, call X with offset" notes. Slide renders at most 1280px wide.
- Old slide renders and large results are replaced with placeholders after they stop being recent (keep the last few).
- Compaction at the model's window minus output reserve minus ~13k, using Pi's compaction helpers with an Office summary prompt: intent, all user requests, decisions, changed IDs, verification state, outstanding work. Current deck state and loaded skills are re-injected after compaction. A breaker stops repeated compaction loops.

## Checkpoints and undo

Before the first write to a slide in a turn, the agent exports that slide (`exportAsBase64`). "Undo turn" restores exported slides (delete + `insertSlidesFromBase64` at the original index), deletes slides the turn created, and reinserts slides it deleted. Restored slides get new IDs; undo warns if you edited those slides after the turn.

## Attachments and web

- Images: model input + insertable. PDF: pdf.js text + page renders. DOCX: mammoth → markdown. XLSX/CSV: SheetJS → tables. PPTX: read content, insert slides.
- Firecrawl (only with a key): `web_search`, `fetch_page` (markdown), images from URLs insertable.

## Skills

Bundled skills are original (Anthropic's pptx skill is license-restricted). Listed by name and description; `load_skill` returns the body; supporting references load on demand. Users add skills by pasting a SKILL.md or importing a URL. Ship a starter set, run varied tasks, log which skills load and help, prune.

## Known risks

- The extension shell's frame hook is unverified against the live `pivot.claude.ai` frame. First thing to check in the work-account spike.
- Office.js has no chart or speaker-notes API. Charts and notes go through generated or edited PPTX slides, which changes slide IDs.
- `chrome.userScripts` needs a one-time manual toggle per browser.
