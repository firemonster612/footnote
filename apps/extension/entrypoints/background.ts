import { defineBackground } from "wxt/utils/define-background";

export default defineBackground(() => {
  chrome.sidePanel.setPanelBehavior({ openPanelOnActionClick: true }).catch((error: unknown) => {
    console.error("Footnote: couldn't make the toolbar button open the side panel", error);
  });
});
