import { afterEach, beforeAll, describe, expect, test, vi } from "vitest";
import type { ToolEnv } from "../../src/contracts.ts";
import {
  attachmentToContent,
  createAttachmentStore,
  createAttachmentTool,
  processAttachment,
} from "../../src/attachments/index.ts";
import { buildDocx, buildPdf, buildPptx, buildXlsx, fileOf } from "./fixtures.ts";

// Node lacks the browser features pdf.js's modern build needs; its legacy build runs the same code.
vi.mock("pdfjs-dist", () => import("pdfjs-dist/legacy/build/pdf.mjs"));
beforeAll(async () => {
  // pdf.js's documented hook for running its worker on the main thread.
  Object.assign(globalThis, {
    pdfjsWorker: await import("pdfjs-dist/legacy/build/pdf.worker.mjs"),
  });
});

function textOf(content: { type: string; text?: string }[]): string {
  return content
    .flatMap((part) => (part.type === "text" && part.text ? [part.text] : []))
    .join("\n");
}

describe("processAttachment", () => {
  test("PDF: text per page and the first pages rendered as PNG", async () => {
    const attachment = await processAttachment(
      fileOf(buildPdf(["Hello PDF", "Second page"]), "report.pdf"),
    );
    expect(attachment).toMatchObject({
      kind: "pdf",
      mimeType: "application/pdf",
      summary: "2 pages",
    });
    expect(attachment.text).toBe("## Page 1\n\nHello PDF\n\n## Page 2\n\nSecond page");
    expect(attachment.images?.map((image) => [image.label, image.mimeType])).toEqual([
      ["Page 1", "image/png"],
      ["Page 2", "image/png"],
    ]);
  });

  test("DOCX: markdown with headings and emphasis", async () => {
    const docx = await buildDocx(
      `<w:p><w:pPr><w:pStyle w:val="Heading1"/></w:pPr><w:r><w:t>Plan</w:t></w:r></w:p>` +
        `<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>Bold</w:t></w:r><w:r><w:t xml:space="preserve"> start.</w:t></w:r></w:p>`,
    );
    const attachment = await processAttachment(fileOf(docx, "plan.docx"));
    expect(attachment).toMatchObject({ kind: "docx", summary: "3 words" });
    expect(attachment.text).toBe("# Plan\n\n__Bold__ start.");
  });

  test("XLSX: one markdown table per sheet, truncation noted", async () => {
    const rows = [
      ["Region", "Revenue"],
      ...Array.from({ length: 501 }, (_, index) => [`R${index}`, index]),
    ];
    const xlsx = buildXlsx({ Sales: rows, Notes: [["a|b"]] });
    const attachment = await processAttachment(fileOf(xlsx, "sales.xlsx"));
    expect(attachment).toMatchObject({ kind: "spreadsheet", summary: "2 sheets" });
    expect(attachment.text).toContain(
      "## Sheet: Sales\n\n| Region | Revenue |\n| --- | --- |\n| R0 | 0 |",
    );
    expect(attachment.text).toContain(
      "| R499 | 499 |\n\n(Truncated: showing 500 of 501 data rows.)",
    );
    expect(attachment.text).toContain("## Sheet: Notes\n\n| a\\|b |");
  });

  test("CSV goes through the spreadsheet path", async () => {
    const attachment = await processAttachment(fileOf("name,score\nAda,9\n", "scores.csv"));
    expect(attachment.kind).toBe("spreadsheet");
    expect(attachment.text).toContain("| name | score |\n| --- | --- |\n| Ada | 9 |");
  });

  test("PPTX: titles, body text, and notes per slide in order", async () => {
    const pptx = await buildPptx([
      { title: "Quarterly review", body: "Revenue up 12%", notes: "Mention the & sign" },
      { title: "Next steps", body: "Hire two engineers" },
    ]);
    const attachment = await processAttachment(fileOf(pptx, "deck.pptx"));
    expect(attachment).toMatchObject({ kind: "pptx", summary: "2 slides" });
    expect(attachment.text).toBe(
      "## Slide 1: Quarterly review\nRevenue up 12%\n\nNotes:\nMention the & sign\n\n## Slide 2: Next steps\nHire two engineers",
    );
  });

  test("text files of any extension, binary rejected, oversized rejected", async () => {
    const text = await processAttachment(fileOf("line one\nline two", "notes.md"));
    expect(text).toMatchObject({ kind: "text", text: "line one\nline two", summary: "2 lines" });

    await expect(
      processAttachment(fileOf(new Uint8Array([0xd0, 0xcf, 0x11, 0xe0, 0xff]), "old.doc")),
    ).rejects.toThrow(`Can't read "old.doc"`);
    await expect(
      processAttachment(fileOf(new Uint8Array(50 * 1024 * 1024 + 1), "huge.txt")),
    ).rejects.toThrow("limited to 50 MB");
  });
});

describe("images", () => {
  afterEach(() => vi.unstubAllGlobals());

  test("re-encodes an image within the size limit when its bytes are too large for providers", async () => {
    // Fake canvas output: PNG keeps 3 bytes a pixel, JPEG half a byte.
    vi.stubGlobal("createImageBitmap", async () => ({ width: 1600, height: 1200, close() {} }));
    vi.stubGlobal(
      "OffscreenCanvas",
      class {
        constructor(
          readonly width: number,
          readonly height: number,
        ) {}
        getContext = () => ({ drawImage() {} });
        convertToBlob = async ({ type }: { type: string }) =>
          new Blob([
            new Uint8Array((this.width * this.height * (type === "image/jpeg" ? 1 : 6)) / 2),
          ]);
      },
    );

    const attachment = await processAttachment(fileOf(new Uint8Array(7_000_000), "photo.png"));

    const [modelImage] = attachment.images ?? [];
    expect(modelImage?.mimeType).toBe("image/jpeg");
    expect(modelImage!.base64.length).toBeLessThanOrEqual(3_500_000);
    expect(attachment.summary).toBe("1600×1200");
  });
});

describe("attachmentToContent", () => {
  test("caps the excerpt, points at read_attachment, and sends at most 6 images", async () => {
    const attachment = {
      ...(await processAttachment(fileOf("x".repeat(30_000), "long.txt"))),
      images: Array.from({ length: 8 }, (_, index) => ({
        mimeType: "image/png",
        base64: "AA==",
        label: `Image ${index + 1}`,
      })),
    };
    const content = attachmentToContent(attachment);
    const text = textOf(content);
    expect(text).toContain(
      `<attachment id="${attachment.id}" name="long.txt" kind="text" summary="1 line">`,
    );
    expect(text).toContain("x".repeat(24_000) + "\n</attachment>");
    expect(text).toContain(`Call read_attachment with id "${attachment.id}" and offset 24000`);
    expect(content.filter((part) => part.type === "image")).toHaveLength(6);
  });
});

describe("read_attachment", () => {
  async function setup(file: File) {
    const store = createAttachmentStore();
    const attachment = await processAttachment(file);
    store.add(attachment);
    const env = { attachments: store } as unknown as ToolEnv; // The tool only reads env.attachments.
    return { tool: createAttachmentTool(env), id: attachment.id };
  }

  test("pages through text with offset and limit", async () => {
    const { tool, id } = await setup(fileOf("abcdefghij", "letters.txt"));
    const result = await tool.execute("call", { id, offset: 2, limit: 3 });
    expect(textOf(result.content)).toBe(
      '"letters.txt", characters 2–5 of 10:\ncde\n[Continue with offset 5.]',
    );
  });

  test("renders any PDF page on demand", async () => {
    const pages = Array.from({ length: 8 }, (_, index) => `Page text ${index + 1}`);
    const { tool, id } = await setup(fileOf(buildPdf(pages), "long.pdf"));
    const result = await tool.execute("call", { id, pages: [8] });
    expect(textOf(result.content)).toBe('"long.pdf": Page 8');
    expect(result.content[1]).toMatchObject({ type: "image", mimeType: "image/png" });
    await expect(tool.execute("call", { id, pages: [9] })).rejects.toThrow("the PDF has 8 pages");
  });

  test("unknown id lists what is available", async () => {
    const { tool, id } = await setup(fileOf("hi", "a.txt"));
    await expect(tool.execute("call", { id: "nope" })).rejects.toThrow(`Available: ${id} (a.txt)`);
  });
});
