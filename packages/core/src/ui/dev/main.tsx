// Standalone dev page: FootnoteRoot on a fake app, framed at side-panel width.
// Query params: ?firstRun, ?disconnect, ?width=320
import { StrictMode } from "react";
import { createRoot } from "react-dom/client";
import { FootnoteRoot } from "../index.ts";
import "../styles.css";
import { createFakeApp } from "./fakeApp.ts";

const params = new URLSearchParams(location.search);
const app = createFakeApp({
  firstRun: params.has("firstRun"),
  disconnect: params.has("disconnect"),
});
const width = Number(params.get("width") ?? 400);

createRoot(document.getElementById("root")!).render(
  <StrictMode>
    <div
      className="mx-auto h-full border-x border-neutral-200 dark:border-neutral-800"
      style={{ width }}
    >
      <FootnoteRoot app={app} />
    </div>
  </StrictMode>,
);
