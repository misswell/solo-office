/** `btoa` instead of `Buffer` so embedded media also resolves inside the renderer. */
export function dataUrl(type: string, bytes: Uint8Array): string {
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  return `data:${type};base64,${btoa(binary)}`;
}

/** Image MIME types by part extension; readers, inserters and the writer all agree on these. */
export const IMAGE_MIME_BY_EXT: Record<string, string> = {
  png: "image/png",
  jpg: "image/jpeg",
  jpeg: "image/jpeg",
  gif: "image/gif",
  bmp: "image/bmp",
  tif: "image/tiff",
  tiff: "image/tiff",
  svg: "image/svg+xml",
  webp: "image/webp",
};

/** Extension to give an inserted image of a given MIME type. */
export const IMAGE_EXT_BY_MIME: Record<string, string> = {
  "image/png": "png",
  "image/jpeg": "jpg",
  "image/gif": "gif",
  "image/bmp": "bmp",
  "image/tiff": "tiff",
  "image/svg+xml": "svg",
  "image/webp": "webp",
};

/** Inverse of `dataUrl`, so a save writes back the very bytes the read put in. */
export function bytesOfDataUrl(url: string): Uint8Array {
  const comma = url.indexOf(",");
  const binary = comma < 0 ? "" : atob(url.slice(comma + 1));
  const out = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i += 1) out[i] = binary.charCodeAt(i);
  return out;
}
