import type { TSchema } from "typebox";
import type { FootnoteTool } from "../contracts.ts";

/** Widens a tool with typed parameters to the FootnoteTool element type tool lists use. */
export function defineTool<T extends TSchema>(tool: FootnoteTool<T>): FootnoteTool {
  // execute's params are contravariant, so a specific tool isn't assignable to FootnoteTool<TSchema>.
  // Sound because the agent validates arguments against `parameters` before calling execute.
  return tool as FootnoteTool;
}
