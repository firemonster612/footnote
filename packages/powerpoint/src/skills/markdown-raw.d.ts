// Vite and WXT inline `?raw` imports as strings.
declare module "*.md?raw" {
  const markdown: string;
  export default markdown;
}
