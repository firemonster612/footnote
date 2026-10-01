import { createChromeStore, createFootnoteApp } from "@footnote/core";
import { FootnoteRoot } from "@footnote/core/ui";
import { powerpointModule } from "@footnote/powerpoint";
import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { createExtensionHost } from "../../lib/extension-host.ts";
import "./style.css";

const app = await createFootnoteApp({
  host: createExtensionHost(powerpointModule.kind),
  hostModule: powerpointModule,
  store: createChromeStore(),
});

// index.html always contains #root.
createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <FootnoteRoot app={app} />
  </StrictMode>,
);
