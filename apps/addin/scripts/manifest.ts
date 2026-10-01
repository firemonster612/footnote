// Writes manifest.local.xml: manifest.xml with every https://localhost:3443 URL pointed at FOOTNOTE_ADDIN_URL.
import { readFileSync, writeFileSync } from "node:fs";

const MANIFEST_BASE_URL = "https://localhost:3443";

const baseUrl = process.env.FOOTNOTE_ADDIN_URL?.replace(/\/+$/, "");
if (!baseUrl) {
  throw new Error(
    "Set FOOTNOTE_ADDIN_URL to the add-in's base URL, e.g. https://devbox.example.ts.net:3443",
  );
}

const appDir = new URL("../", import.meta.url);
const manifest = readFileSync(new URL("manifest.xml", appDir), "utf8").replaceAll(
  MANIFEST_BASE_URL,
  baseUrl,
);
writeFileSync(new URL("manifest.local.xml", appDir), manifest);
console.log(`Wrote manifest.local.xml for ${baseUrl}`);
