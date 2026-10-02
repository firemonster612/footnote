import { bytesToBase64 } from "./base64.ts";
import type { ExtractedContent } from "./extracted-content.ts";

const MODEL_IMAGE_MAX_PX = 2000;
// Anthropic rejects images over 5 MB; staying well under it leaves room for other providers' limits.
const MODEL_IMAGE_MAX_BASE64_CHARS = 3_500_000;
const SHRINK_STEP = 0.75;
const MODEL_IMAGE_TYPES = new Set(["image/png", "image/jpeg", "image/gif", "image/webp"]);

/**
 * Records dimensions and prepares a model-ready copy (≤2000px, under ~3.5 MB of base64, PNG/JPEG/GIF/WebP); the
 * original stays in `base64`. A copy that's still too large is re-encoded as JPEG, shrinking until it fits.
 */
export async function processImage(
  file: Blob,
  mimeType: string,
  base64: string,
): Promise<ExtractedContent> {
  const bitmap = await createImageBitmap(file);
  try {
    const { width, height } = bitmap;
    let scale = Math.min(1, MODEL_IMAGE_MAX_PX / Math.max(width, height));
    let modelImage =
      scale === 1 && MODEL_IMAGE_TYPES.has(mimeType)
        ? { mimeType, base64 }
        : await encodeScaled(bitmap, scale, mimeType === "image/jpeg" ? "image/jpeg" : "image/png");
    while (modelImage.base64.length > MODEL_IMAGE_MAX_BASE64_CHARS) {
      modelImage = await encodeScaled(bitmap, scale, "image/jpeg");
      scale *= SHRINK_STEP;
    }
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
