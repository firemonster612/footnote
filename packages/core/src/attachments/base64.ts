// String.fromCharCode spreads its arguments onto the stack; chunking keeps large files under the argument limit.
const CHUNK_BYTES = 0x8000;

export function bytesToBase64(bytes: Uint8Array): string {
  let binary = "";
  for (let start = 0; start < bytes.length; start += CHUNK_BYTES) {
    binary += String.fromCharCode(...bytes.subarray(start, start + CHUNK_BYTES));
  }
  return btoa(binary);
}

export function base64ToBytes(base64: string): Uint8Array<ArrayBuffer> {
  return Uint8Array.from(atob(base64), (char) => char.charCodeAt(0));
}
