import JSZip from "jszip";
import PptxGenJS from "pptxgenjs";
import { utils, write } from "xlsx";

/** A minimal valid PDF with one Helvetica text line per page. */
export function buildPdf(pageTexts: string[]): Uint8Array {
  const pageIds = pageTexts.map((_, index) => 3 + index * 2);
  const fontId = 3 + pageTexts.length * 2;
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    `<< /Type /Pages /Kids [${pageIds.map((id) => `${id} 0 R`).join(" ")}] /Count ${pageTexts.length} >>`,
    ...pageTexts.flatMap((text, index) => {
      const stream = `BT /F1 12 Tf 20 100 Td (${text}) Tj ET`;
      return [
        `<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents ${pageIds[index]! + 1} 0 R /Resources << /Font << /F1 ${fontId} 0 R >> >> >>`,
        `<< /Length ${stream.length} >>\nstream\n${stream}\nendstream`,
      ];
    }),
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
  ];
  let pdf = "%PDF-1.4\n";
  const offsets = objects.map((object, index) => {
    const offset = pdf.length;
    pdf += `${index + 1} 0 obj\n${object}\nendobj\n`;
    return offset;
  });
  const xrefOffset = pdf.length;
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  pdf += offsets.map((offset) => `${String(offset).padStart(10, "0")} 00000 n \n`).join("");
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xrefOffset}\n%%EOF`;
  return new TextEncoder().encode(pdf);
}

/** A bare-bones DOCX: content types, package rels, and one document part. */
export async function buildDocx(bodyXml: string): Promise<Uint8Array> {
  const zip = new JSZip();
  zip.file(
    "[Content_Types].xml",
    `<?xml version="1.0" encoding="UTF-8"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/></Types>`,
  );
  zip.file(
    "_rels/.rels",
    `<?xml version="1.0" encoding="UTF-8"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/></Relationships>`,
  );
  zip.file(
    "word/document.xml",
    `<?xml version="1.0" encoding="UTF-8"?><w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>${bodyXml}</w:body></w:document>`,
  );
  return zip.generateAsync({ type: "uint8array" });
}

export function buildXlsx(sheets: Record<string, unknown[][]>): Uint8Array {
  const workbook = utils.book_new();
  for (const [name, rows] of Object.entries(sheets))
    utils.book_append_sheet(workbook, utils.aoa_to_sheet(rows), name);
  return write(workbook, { type: "array", bookType: "xlsx" });
}

export async function buildPptx(
  slides: { title: string; body: string; notes?: string }[],
): Promise<ArrayBuffer> {
  const pptx = new PptxGenJS();
  pptx.defineSlideMaster({
    title: "TITLE_AND_BODY",
    objects: [
      {
        placeholder: {
          options: { name: "title", type: "title", x: 0.5, y: 0.3, w: 9, h: 1 },
          text: "",
        },
      },
    ],
  });
  for (const { title, body, notes } of slides) {
    const slide = pptx.addSlide({ masterName: "TITLE_AND_BODY" });
    slide.addText(title, { placeholder: "title" });
    slide.addText(body, { x: 0.5, y: 1.5, w: 9, h: 3 });
    if (notes) slide.addNotes(notes);
  }
  // pptxgenjs returns the type named by outputType but declares a union.
  return (await pptx.write({ outputType: "arraybuffer" })) as ArrayBuffer;
}

export function fileOf(content: string | ArrayBuffer | Uint8Array, name: string, type = ""): File {
  return new File([typeof content === "string" ? content : new Uint8Array(content)], name, {
    type,
  });
}
