import tailwindcss from "@tailwindcss/vite";
import { fileURLToPath } from "node:url";
import { defineConfig } from "vite";

// Serves the UI dev page (src/ui/dev) with a fake FootnoteApp: `bun run ui:dev`.
export default defineConfig({
  root: fileURLToPath(new URL("../src/ui/dev", import.meta.url)),
  plugins: [tailwindcss()],
  server: { host: true },
});
