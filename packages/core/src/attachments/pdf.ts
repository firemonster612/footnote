import { GlobalWorkerOptions, getDocument, type PDFDocumentProxy } from "pdfjs-dist";
import workerUrl from "pdfjs-dist/build/pdf.worker.min.mjs?url";
import type { AttachmentImage, ExtractedContent } from "./extracted-content.ts";

GlobalWorkerOptions.workerSrc = workerUrl;

const RENDERED_PAGE_COUNT = 6;
const RENDER_MAX_PX = 1200;

/** pdf.js types `canvasFactory` as Object; its DOM and Node factories both return this shape. */
interface CanvasFactory {
  create(
    width: number,
    height: number,
  ): { canvas: HTMLCanvasElement; context: CanvasRenderingContext2D };
  destroy(canvasAndContext: { canvas: HTMLCanvasElement; context: CanvasRenderingContext2D }): void;
}

/** Text of every page plus PNG renders of the first few. `bytes` is consumed (pdf.js may detach it). */
export async function processPdf(bytes: Uint8Array): Promise<ExtractedContent> {
  return withPdf(bytes, async (pdf) => {
    const pageNumbers = Array.from({ length: pdf.numPages }, (_, index) => index + 1);
    const pageTexts = await Promise.all(
      pageNumbers.map((pageNumber) => readPageText(pdf, pageNumber)),
    );
    const text = pageTexts.some((pageText) => pageText.length > 0)
      ? pageTexts.map((pageText, index) => `## Page ${index + 1}\n\n${pageText}`).join("\n\n")
      : "(No extractable text: the PDF is probably scanned. Look at the page images.)";
    const images = await renderPages(pdf, pageNumbers.slice(0, RENDERED_PAGE_COUNT));
    return { text, images, summary: pdf.numPages === 1 ? "1 page" : `${pdf.numPages} pages` };
  });
}

/** Renders the given 1-based pages as PNGs. */
export async function renderPdfPages(
  bytes: Uint8Array,
  pageNumbers: number[],
): Promise<AttachmentImage[]> {
  return withPdf(bytes, async (pdf) => {
    const outOfRange = pageNumbers.filter((pageNumber) => pageNumber > pdf.numPages);
    if (outOfRange.length > 0) {
      throw new Error(
        `Page ${outOfRange.join(", ")} is out of range: the PDF has ${pdf.numPages} pages.`,
      );
    }
    return renderPages(pdf, pageNumbers);
  });
}

async function withPdf<T>(
  bytes: Uint8Array,
  read: (pdf: PDFDocumentProxy) => Promise<T>,
): Promise<T> {
  const loading = getDocument({ data: bytes });
  try {
    return await read(await loading.promise);
  } finally {
    await loading.destroy();
  }
}

async function readPageText(pdf: PDFDocumentProxy, pageNumber: number): Promise<string> {
  const page = await pdf.getPage(pageNumber);
  const content = await page.getTextContent();
  return content.items
    .map((item) => ("str" in item ? item.str + (item.hasEOL ? "\n" : "") : ""))
    .join("")
    .trim();
}

async function renderPages(
  pdf: PDFDocumentProxy,
  pageNumbers: number[],
): Promise<AttachmentImage[]> {
  // Sequential: each render holds a full-size canvas.
  const images: AttachmentImage[] = [];
  for (const pageNumber of pageNumbers) {
    const page = await pdf.getPage(pageNumber);
    const unscaled = page.getViewport({ scale: 1 });
    const viewport = page.getViewport({
      scale: RENDER_MAX_PX / Math.max(unscaled.width, unscaled.height),
    });
    // See CanvasFactory: pdf.js leaves the factory untyped.
    const factory = pdf.canvasFactory as CanvasFactory;
    const target = factory.create(Math.ceil(viewport.width), Math.ceil(viewport.height));
    await page.render({ canvas: target.canvas, canvasContext: target.context, viewport }).promise;
    const dataUrl = target.canvas.toDataURL("image/png");
    factory.destroy(target);
    images.push({
      mimeType: "image/png",
      base64: dataUrl.slice(dataUrl.indexOf(",") + 1),
      label: `Page ${pageNumber}`,
    });
  }
  return images;
}
