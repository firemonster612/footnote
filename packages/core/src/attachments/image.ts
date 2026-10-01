import { bytesToBase64 } from "./base64.ts";
import type { ExtractedContent } from "./extracted-content.ts";

const MODEL_IMAGE_MAX_PX = 2000;
const MODEL_IMAGE_TYPES = new Set(["image/png", "image/jpeg", "image/gif", "image/webp"]);

/** Records dimensions and prepares a model-ready copy (≤2000px, PNG/JPEG/GIF/WebP); the original stays in `base64`. */
export async function processImage(
  file: Blob,
  mimeType: string,
  base64: string,
): Promise<ExtractedContent> {
  const bitmap = await createImageBitmap(file);
  try {
    const { width, height } = bitmap;
    const scale = Math.min(1, MODEL_IMAGE_MAX_PX / Math.max(width, height));
    const modelImage =
      scale === 1 && MODEL_IMAGE_TYPES.has(mimeType)
        ? { mimeType, base64 }
        : await encodeScaled(bitmap, scale, mimeType === "image/jpeg" ? "image/jpeg" : "image/png");
    return { summary: `${width}×${height}`, images: [{ ...modelImage, label: "Image" }] };
  } finally {
    bitmap.close();
  }
}

async function encodeScaled(
  bitmap: ImageBitmap,
  scale: number,
  mimeType: "image/png" | "image/jpeg",
): Promise<{ mimeType: string; base64: string }> {
  const canvas = new OffscreenCanvas(
    Math.round(bitmap.width * scale),
    Math.round(bitmap.height * scale),
  );
  const context = canvas.getContext("2d");
  if (!context) throw new Error("Could not create a 2D canvas to resize the image.");
  context.drawImage(bitmap, 0, 0, canvas.width, canvas.height);
  const blob = await canvas.convertToBlob({ type: mimeType, quality: 0.9 });
  return { mimeType, base64: bytesToBase64(new Uint8Array(await blob.arrayBuffer())) };
}
