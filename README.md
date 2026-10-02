<p align="center"><img src="assets/logo.svg" width="72" alt="Footnote logo"></p>

# Footnote

Footnote is an AI assistant that edits PowerPoint decks while you watch, using whatever model you point it at. You chat with it in a side panel; it reads the open deck, changes slides, shapes, text, tables and charts, renders slides to check its own work, reads attachments and searches the web.

It talks to any Anthropic-compatible or OpenAI-compatible endpoint. It was built against [CLIProxyAPI](https://github.com/router-for-me/CLIProxyAPI), which lets you use Claude and GPT subscriptions through one URL and key, but OpenRouter, a local server or a provider's own API work the same way.

> **Alpha.** This is v0.0.1. It works on real decks, and it can also make a mess of them. Keep a copy of anything you care about, and expect rough edges.

## How it attaches to PowerPoint

PowerPoint only lets code edit a deck from inside an Office add-in. Footnote comes in two forms:

|          | Chrome extension                                                                                                                                                                          | Office add-in                                                                                     |
| -------- | ----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- | ------------------------------------------------------------------------------------------------- |
| Works in | PowerPoint for the web, in Chrome, Edge, Brave, Helium or another Chromium browser                                                                                                        | PowerPoint for the web and desktop PowerPoint                                                     |
| Needs    | An Office add-in already open in the deck: currently the Claude add-in (`pivot.claude.ai`). Footnote borrows that pane's connection to the deck; your organization must have deployed it. | Permission to sideload add-ins. Personal Microsoft accounts allow this; many work accounts don't. |
| UI       | Browser side panel; tasks keep running when you switch tabs                                                                                                                               | Office task pane                                                                                  |

If your work account blocks sideloading but has the Claude add-in, use the extension. Otherwise the add-in is the cleaner route.

## Quick start: Chrome extension

1. **Get the extension.** Download `footnote-0.0.1-chrome.zip` from the [latest release](../../releases) and unzip it. (Or build it, see [Build from source](#build-from-source).)
2. **Load it.** Open `chrome://extensions`, turn on **Developer mode** (top right), click **Load unpacked**, and pick the unzipped folder.
3. **Allow code mode** (optional but recommended). On Footnote's **Details** page, turn on **Allow user scripts**. Without it, everything works except the tool that runs model-written Office.js.
4. **Open your deck.** Open a presentation in PowerPoint for the web and open the **Claude** add-in pane. Leave the pane open. If the tab was already open before you installed Footnote, reload it.
5. **Open Footnote.** Click the Footnote icon in the toolbar. The panel shows the deck's name once it's connected.
6. **Connect a model.** Enter your endpoint URL and API key, click **Test connection**, then **Continue**. For CLIProxyAPI that's something like `http://localhost:8317` and one of the keys under `api-keys:` in its `config.yaml`.
7. **Ask for something.** Pick a model and effort level under the message box, then try "What's in this deck?" or "Tighten slide 2 and check it visually."

The panel stays attached to that tab: it hides when you switch away and comes back with the results when you return. The toolbar icon shows **…** while a task runs in the background and **!** when it's waiting for your approval.

## Quick start: Office add-in

You need [bun](https://bun.sh) and a Microsoft account that can upload add-ins.

```sh
git clone https://github.com/firemonster612/footnote.git
cd footnote
bun install
cd apps/addin
bun run certs   # trusts a localhost HTTPS certificate (asks for your password)
bun run dev     # serves the task pane at https://localhost:3443
```

Then, in PowerPoint for the web: **Home → Add-ins → More Add-ins → My Add-ins → Upload My Add-in**, and upload `apps/addin/manifest.xml`. A **Footnote** button appears on the Home tab.

To serve the task pane from another machine (for example over Tailscale), run the dev server there with `FOOTNOTE_ADDIN_URL=https://your-host:3443 bun run dev`, generate a matching manifest with `FOOTNOTE_ADDIN_URL=https://your-host:3443 bun run manifest`, and upload `manifest.local.xml` instead.

## Using it

- **Permissions.** In **Ask** mode (the shield toggle), Footnote asks before every edit; reading, rendering and web search never ask. Code runs always ask in Ask mode and show you the script. **Full access** runs everything without asking. Set the default in Settings.
- **Undo and revert.** The undo button reverts the last turn's slide changes. Hover any of your messages and click **Revert** (twice) to undo that request and everything after it; the text goes back in the message box so you can edit and resend.
- **Queue and steer.** Press Enter while Footnote is working to queue a message for when it finishes. **Steer** on a queued message sends it into the running task right away.
- **Attachments.** Drop or paste images, PDFs, Word, Excel/CSV and PowerPoint files into the panel.
- **Web search.** Add a [Firecrawl](https://firecrawl.dev) API key in Settings to enable `web_search` and `fetch_page`.
- **Skills.** Footnote ships with presentation skills it loads when needed. Add your own in Settings by pasting a `SKILL.md` or importing one from a URL.

### Which API format is used

Footnote asks your endpoint for `/v1/models`. Models owned by `anthropic` use the Anthropic Messages API, `openai` models use the OpenAI Responses API, and everything else uses Chat Completions. You can override this per model in Settings. Reasoning effort is sent as the provider's own parameter.

## Privacy

Footnote sends the deck content it reads, your messages and your attachments to the endpoint you configure, and to Firecrawl if you enable web search. Nothing goes anywhere else. Settings, API keys and chat history stay in your browser's extension storage. If you use it on work documents, check that your organization is fine with that content reaching your endpoint.

## Known limits

- The extension depends on the Claude add-in's frame at `pivot.claude.ai`. If Anthropic changes how that add-in loads, the extension stops connecting until it's updated.
- Office.js has no API for charts or speaker notes. Footnote handles both by rebuilding the slide from an edited copy, which gives the slide a new internal ID.
- Undo history lives in memory. It's lost when the extension's background engine closes (10 minutes after the panel closes with nothing running) or the add-in pane reloads. Revert still rewinds the chat but can't undo those older slide changes.
- Attachments aren't saved with chats. After a reload, the model can still see the excerpt it was sent, but it can't read further into the file or insert it.
- Undo can't restore slides that model-written code deleted unless the model listed them first.
- PowerPoint for the web applies edits more slowly than desktop PowerPoint; big redesigns take a while.
- In the add-in shell, model-written code runs in the same page that stores your settings, so it could read your API key. Ask mode always shows the code first. The extension keeps keys out of the page where code runs.
- Slide reads report text formatting per paragraph, not per word: Office.js has no API for formatting runs inside a paragraph.
- Footnote can't insert an image straight from a web address yet; attach the image instead.

## Build from source

```sh
bun install
cd apps/extension && bun run build   # unpacked extension in apps/extension/.output/chrome-mv3
bun run zip                          # zipped build in apps/extension/.output
```

To let the extension attach to another add-in's frame, list its origins at build time: `FOOTNOTE_ADDIN_MATCHES="https://addin.example.com/*" bun run build`.

### Development

| Task                                    | Command                                                                                                  |
| --------------------------------------- | -------------------------------------------------------------------------------------------------------- |
| Tests                                   | `bunx vitest run` (repo root)                                                                            |
| Typecheck one package                   | `cd packages/core && bunx tsc -p .`                                                                      |
| Lint / format                           | `bunx oxlint`, `bunx oxfmt`                                                                              |
| UI dev page with a fake chat            | `cd packages/core && bun run ui:dev`                                                                     |
| Live smoke test against a real endpoint | `FOOTNOTE_ENDPOINT=http://localhost:8317 FOOTNOTE_API_KEY=… bun packages/core/test/runtime/liveSmoke.ts` |

The code is a bun workspace:

- `packages/core`: agent runtime (built on [Pi](https://github.com/earendil-works)'s agent core), providers, permissions, context and compaction, attachments, web tools, skills, and the React UI.
- `packages/powerpoint`: everything PowerPoint-specific: Office.js operations, the tools the model sees, undo, deck summaries and bundled skills. Word and Excel would be sibling packages.
- `apps/extension`: the Chrome extension (side panel, offscreen engine, service worker, bridge into the add-in frame).
- `apps/addin`: the Office add-in task pane.

`docs/design.md` explains the architecture and the reasoning behind it.

## License

[MIT](LICENSE)
