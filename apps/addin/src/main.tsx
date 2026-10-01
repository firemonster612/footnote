import { createFootnoteApp, createLocalStore } from "@footnote/core";
import { FootnoteRoot } from "@footnote/core/ui";
import { powerpointModule } from "@footnote/powerpoint";
import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { createTaskPaneHost } from "./taskpane-host.ts";
import "./style.css";

await Office.onReady();

const app = await createFootnoteApp({
  host: createTaskPaneHost(powerpointModule),
  hostModule: powerpointModule,
  store: createLocalStore(),
});

// index.html always contains #root.
createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <FootnoteRoot app={app} />
  </StrictMode>,
);
