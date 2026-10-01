import type { FootnoteTool, ToolEnv } from "@footnote/core/contracts";
import { createAdvancedTools } from "../advanced/index.ts";
import type { WriteGuard } from "../writeGuard.ts";
import { createCodeTool } from "./code.ts";
import { createReadTools } from "./read.ts";
import type { SlidePositions } from "./schemas.ts";
import { createWriteTools } from "./write.ts";

export function createPowerPointTools(
  env: ToolEnv,
  guard: WriteGuard,
  positions: SlidePositions,
): FootnoteTool[] {
  return [
    ...createReadTools(env, guard, positions),
    ...createWriteTools(env, guard, positions),
    ...createAdvancedTools(env, guard, positions),
    createCodeTool(env, guard),
  ];
}
