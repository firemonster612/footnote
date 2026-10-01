import type { FootnoteTool, ToolEnv } from "@footnote/core/contracts";
import { createAdvancedTools } from "../advanced/index.ts";
import type { WriteGuard } from "../writeGuard.ts";
import { createCodeTool } from "./code.ts";
import { createReadTools } from "./read.ts";
import { createWriteTools } from "./write.ts";

export function createPowerPointTools(env: ToolEnv, guard: WriteGuard): FootnoteTool[] {
  return [
    ...createReadTools(env, guard),
    ...createWriteTools(env, guard),
    ...createAdvancedTools(env, guard),
    createCodeTool(env, guard),
  ];
}
