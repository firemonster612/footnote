import tailwindcss from "@tailwindcss/vite";
import basicSsl from "@vitejs/plugin-basic-ssl";
import react from "@vitejs/plugin-react";
import { existsSync, readFileSync } from "node:fs";
import { homedir } from "node:os";
import { join } from "node:path";
import { defineConfig, type ServerOptions } from "vite";

/** Public base URL when served from a remote dev box, e.g. https://devbox.example.ts.net:3443. Also read by `bun run manifest`. */
const publicHost =
  process.env.FOOTNOTE_ADDIN_URL && new URL(process.env.FOOTNOTE_ADDIN_URL).hostname;

/** Certificates from `bun run certs` (office-addin-dev-certs), which Office trusts. Otherwise a self-signed one. */
const devCertDir = join(homedir(), ".office-addin-dev-certs");
const devCerts = existsSync(join(devCertDir, "localhost.crt"))
  ? {
      cert: readFileSync(join(devCertDir, "localhost.crt")),
      key: readFileSync(join(devCertDir, "localhost.key")),
    }
  : undefined;

const server: ServerOptions = {
  host: "0.0.0.0",
  port: 3443,
  strictPort: true,
  ...(devCerts && { https: devCerts }),
  ...(publicHost && { allowedHosts: [publicHost] }),
};

export default defineConfig({
  plugins: [
    react(),
    tailwindcss(),
    ...(devCerts
      ? []
      : [
          basicSsl({
            name: "footnote",
            domains: ["localhost", ...(publicHost ? [publicHost] : [])],
          }),
        ]),
  ],
  publicDir: "../../assets/icons",
  server,
  preview: server,
});
