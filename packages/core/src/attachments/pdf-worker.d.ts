// Vite (and WXT, which builds on it) resolves `?url` imports to the emitted asset's URL.
declare module "pdfjs-dist/build/pdf.worker.min.mjs?url" {
  const url: string;
  export default url;
}
