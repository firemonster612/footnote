import { bytesToBase64 } from "../attachments/base64.ts";

const MAX_IMAGE_BYTES = 20 * 1024 * 1024;

/** Downloads an image for insertion. Throws when the URL isn't an image or exceeds 20 MB. */
export async function fetchImageAsBase64(
  url: string,
): Promise<{ mimeType: string; base64: string }> {
  const response = await fetch(url);
  if (!response.ok) throw new Error(`Fetching image ${url} failed: HTTP ${response.status}.`);
  const mimeType = (response.headers.get("content-type") ?? "").split(";")[0]!.trim().toLowerCase();
  if (!mimeType.startsWith("image/")) {
    throw new Error(`${url} is not an image (content type "${mimeType || "none"}").`);
  }
  if (Number(response.headers.get("content-length")) > MAX_IMAGE_BYTES)
    throw new Error(tooLarge(url));
  const bytes = new Uint8Array(await response.arrayBuffer());
  if (bytes.byteLength > MAX_IMAGE_BYTES) throw new Error(tooLarge(url));
  return { mimeType, base64: bytesToBase64(bytes) };
}

function tooLarge(url: string): string {
  return `Image at ${url} is larger than ${MAX_IMAGE_BYTES / 1024 / 1024} MB.`;
}
