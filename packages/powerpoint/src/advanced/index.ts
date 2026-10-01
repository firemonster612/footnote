// Agent-side entry. The Office realm imports `./ops.ts` and `./helpers.ts` directly so PptxGenJS/JSZip stay out of it.
export { advancedOps } from "./ops.ts";
export { createAdvancedTools } from "./tools.ts";
export { codeHelpersReference, createFootnoteHelpers } from "./helpers.ts";
