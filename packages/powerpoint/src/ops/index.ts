// Office-realm entry: bundled into the add-in frame by each shell. Must not import agent-side code.
import type { OpRegistry } from "@footnote/core/contracts";
import { advancedOps } from "../advanced/ops.ts";
import { readOps } from "./read.ts";
import { shapeOps } from "./shapes.ts";
import { slideOps } from "./slides.ts";
import { tableOps } from "./tables.ts";

export const powerpointOps: OpRegistry = {
  ...readOps,
  ...slideOps,
  ...shapeOps,
  ...tableOps,
  ...advancedOps,
};

export { createFootnoteHelpers } from "../advanced/helpers.ts";
