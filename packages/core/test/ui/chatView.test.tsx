import { renderToStaticMarkup } from "react-dom/server";
import { expect, test } from "vitest";
import { ChatView, formatRevertReport } from "../../src/ui/chat/ChatView.tsx";
import { TooltipProvider } from "../../src/ui/components/tooltip.tsx";
import { createFakeApp } from "../../src/ui/dev/fakeApp.ts";

test("the demo chat renders every message, tool and approval state", async () => {
  const app = createFakeApp();
  const session = await app.chats.open("chat-demo");
  const tools = new Map(
    app.hostModule
      .createTools({
        host: app.host,
        attachments: { list: () => [], get: () => undefined },
        settings: app.settings.get,
      })
      .map((tool) => [tool.name, tool]),
  );
  const html = renderToStaticMarkup(
    <TooltipProvider>
      <ChatView session={session} models={await app.models.list()} tools={tools} />
    </TooltipProvider>,
  );

  const expected = [
    "Q3 revenue.xlsx", // attachment chip on the user message
    "Read slide 3", // describeCall label: positions, not slide IDs
    "Waiting for approval",
    "Queued",
    "Failed",
    "Provider returned 529", // assistant error
    "Allow writes for this chat", // write approval
    "footnote.pt(48)", // code approval shows the code
    "brand-guidelines.pdf", // staged attachment
    "Context 48k of 200k tokens",
    "<table",
    "Earlier conversation summarized",
    "Working…", // stays up while approvals wait
  ];
  expect(expected.filter((text) => !html.includes(text))).toEqual([]);
  expect(html).not.toContain("deck_state");
  // Only the write approval offers chat-wide approval; the code approval doesn't.
  expect(html.split("Allow writes for this chat")).toHaveLength(2);
});

test("the revert notice only reports what happened", () => {
  const lost = "Slide changes from 1 earlier turn couldn't be undone.";
  expect(formatRevertReport({ restored: 0, removed: 0, warnings: [lost] })).toBe(
    `Reverted to that request; it's back in the box below.\n${lost}`,
  );
  expect(formatRevertReport({ restored: 0, removed: 0, warnings: [] })).toBe(
    "Reverted to that request; it's back in the box below. No slide changes needed undoing.",
  );
  expect(formatRevertReport({ restored: 2, removed: 1, warnings: [] })).toBe(
    "Reverted to that request; it's back in the box below. Slides: 2 restored, 1 removed.",
  );
});
