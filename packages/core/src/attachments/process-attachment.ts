import type { AttachmentKind, ProcessedAttachment } from "../contracts.ts";
import { bytesToBase64 } from "./base64.ts";
import type { ExtractedContent } from "./extracted-content.ts";

const MAX_ATTACHMENT_BYTES = 50 * 1024 * 1024;

interface FileFormat {
  kind: AttachmentKind;
  mimeType: string;
}

const formatsByExtension: Record<string, FileFormat> = {
  png: { kind: "image", mimeType: "image/png" },
  jpg: { kind: "image", mimeType: "image/jpeg" },
  jpeg: { kind: "image", mimeType: "image/jpeg" },
  gif: { kind: "image", mimeType: "image/gif" },
  webp: { kind: "image", mimeType: "image/webp" },
  bmp: { kind: "image", mimeType: "image/bmp" },
  pdf: { kind: "pdf", mimeType: "application/pdf" },
  docx: {
    kind: "docx",
    mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
  },
  pptx: {
    kind: "pptx",
    mimeType: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
  },
  xlsx: {
    kind: "spreadsheet",
    mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  },
  xlsm: { kind: "spreadsheet", mimeType: "application/vnd.ms-excel.sheet.macroEnabled.12" },
  xls: { kind: "spreadsheet", mimeType: "application/vnd.ms-excel" },
  ods: { kind: "spreadsheet", mimeType: "application/vnd.oasis.opendocument.spreadsheet" },
  csv: { kind: "spreadsheet", mimeType: "text/csv" },
  tsv: { kind: "spreadsheet", mimeType: "text/tab-separated-values" },
};

/**
 * Reads a user file into what the model and the tools need. Unknown extensions are accepted when
 * the bytes are UTF-8 text. Throws a user-readable Error for oversized or unsupported files.
 */
export async function processAttachment(file: File): Promise<ProcessedAttachment> {
  if (file.size > MAX_ATTACHMENT_BYTES) {
    const megabytes = (file.size / 1024 / 1024).toFixed(1);
    throw new Error(`"${file.name}" is ${megabytes} MB; attachments are limited to 50 MB.`);
  }
  const extension = file.name.slice(file.name.lastIndexOf(".") + 1).toLowerCase();
  const format = formatsByExtension[extension] ?? {
    kind: "text",
    mimeType: file.type || "text/plain",
  };
  const bytes = await file.arrayBuffer();
  const base64 = bytesToBase64(new Uint8Array(bytes));
  const content = await extractContent(format, file, bytes, base64);
  return {
    id: crypto.randomUUID(),
    name: file.name,
    mimeType: format.mimeType,
    kind: format.kind,
    size: file.size,
    base64,
    ...content,
  };
}

// Format libraries load on first use so the panel doesn't pay for pdf.js, mammoth, and SheetJS up front.
async function extractContent(
  format: FileFormat,
  file: File,
  bytes: ArrayBuffer,
  base64: string,
): Promise<ExtractedContent> {
  switch (format.kind) {
    case "image":
      return (await import("./image.ts")).processImage(file, format.mimeType, base64);
    case "pdf":
      return (await import("./pdf.ts")).processPdf(new Uint8Array(bytes));
    case "docx":
      return (await import("./docx.ts")).processDocx(bytes);
    case "spreadsheet":
      return (await import("./spreadsheet.ts")).processSpreadsheet(bytes);
    case "pptx":
      return (await import("./pptx.ts")).processPptx(bytes);
    case "text":
      return processText(file.name, bytes);
  }
}

function processText(name: string, bytes: ArrayBuffer): ExtractedContent {
  let text: string;
  try {
    text = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
  } catch (error) {
    throw new Error(unsupportedMessage(name), { cause: error });
  }
  if (text.includes("\0")) throw new Error(unsupportedMessage(name));
  const lines = text.split("\n").length;
  return { text, summary: lines === 1 ? "1 line" : `${lines.toLocaleString("en-US")} lines` };
}

function unsupportedMessage(name: string): string {
  return `Can't read "${name}". Attach images, PDFs, .docx, .xlsx/.csv, .pptx, or text files.`;
}
