import tailwindcss from "@tailwindcss/vite";
import { defineConfig } from "wxt";

/** Comma-separated match patterns for extra add-in frames the bridge should serve, e.g. "https://localhost:3443/*". */
const extraAddinMatches = (process.env.FOOTNOTE_ADDIN_MATCHES ?? "")
  .split(",")
  .map((pattern) => pattern.trim())
  .filter(Boolean);

export default defineConfig({
  modules: ["@wxt-dev/module-react"],
  imports: false,
  publicDir: "../../assets/icons",
  vite: () => ({ plugins: [tailwindcss()] }),
  manifest: {
    name: "Footnote",
    description: "Model-agnostic assistant that edits PowerPoint decks live.",
    action: { default_title: "Open Footnote" },
    permissions: ["sidePanel", "storage", "scripting", "userScripts", "tabs"],
    host_permissions: [
      "https://pivot.claude.ai/*",
      "https://*.officeapps.live.com/*",
      "https://*.sharepoint.com/*",
      "https://onedrive.live.com/*",
      "https://*.office.com/*",
      "https://*.cloud.microsoft/*",
      // Model endpoints and Firecrawl are user-configured, so requests can go anywhere.
      "<all_urls>",
    ],
  },
  hooks: {
    // Every content script here is an add-in frame script (bridge + relay), so all of them get the extra origins.
    // They share one matches array, so assign a fresh one instead of pushing.
    "build:manifestGenerated": (_wxt, manifest) => {
      for (const script of manifest.content_scripts ?? []) {
        script.matches = [...(script.matches ?? []), ...extraAddinMatches];
      }
    },
  },
});
