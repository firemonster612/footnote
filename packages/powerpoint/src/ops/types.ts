// Shapes exchanged between ops (Office realm) and tools (agent side). Structured-clone safe; type-only on the agent side.
// Units are points. Indexes are zero-based here; tools present one-based positions to the model.

export interface Bounds {
  left: number;
  top: number;
  width: number;
  height: number;
}

/** null means the text range mixes values for that property. */
export interface FontInfo {
  name: string | null;
  size: number | null;
  color: string | null;
  bold: boolean | null;
  italic: boolean | null;
  underline: string | null;
}

export interface ParagraphInfo {
  text: string;
  font: FontInfo;
}

export interface ShapeInfo extends Bounds {
  id: string;
  name: string;
  type: string;
  z?: number;
  rotation?: number;
  placeholder?: string;
  text?: string;
  /** Detail reads only. */
  font?: FontInfo;
  align?: string | null;
  bullets?: boolean | null;
  autoSize?: string;
  /** Detail reads only, when the shape's font is mixed: formatting per paragraph. */
  paragraphs?: ParagraphInfo[];
  fill?: { type: string; color?: string; transparency?: number };
  line?: { color: string; weight: number; dash: string };
  table?: { rows: number; columns: number; values: string[][] };
  children?: ShapeInfo[];
}

export interface SlideOutline {
  id: string;
  index: number;
  layout: string;
  title?: string;
  shapeCount: number;
  fingerprint: string;
}

export interface DeckInfo {
  title?: string;
  /** Present when PowerPointApi 1.10 is available. */
  slideWidth?: number;
  slideHeight?: number;
  slides: SlideOutline[];
}

export interface ThemeInfo {
  /** Theme color name → #RRGGBB. Sampled from master shapes when the theme API is unavailable. */
  colors: Record<string, string>;
  sampled: boolean;
  fonts: { heading?: string; body?: string };
}

export interface LayoutsInfo {
  masters: { id: string; name: string; layouts: { id: string; name: string }[] }[];
}

export interface SelectionInfo {
  slideIds: string[];
  shapeIds: string[];
  text?: { shapeId?: string; text: string; start: number; length: number };
}

export interface DeckState {
  deck: DeckInfo;
  selection: SelectionInfo;
  /** Shapes of the first selected slide. */
  selectedShapes?: ShapeInfo[];
  theme: ThemeInfo;
}

export interface SlideDetail {
  id: string;
  index: number;
  layout: string;
  fingerprint: string;
  shapes: ShapeInfo[];
}

export interface SlideState {
  index: number;
  fingerprint: string;
  /** Single-slide PPTX, when requested. */
  base64?: string;
}

/** A slide as it was before a turn first wrote it, for undo. */
export interface SlideSnapshot {
  index: number;
  /** Single-slide PPTX. */
  base64: string;
}

export interface WriteReceipt {
  /** IDs of the slides or shapes the op changed or created. */
  changed: string[];
  /** Read-back values of what changed. */
  verified: Record<string, unknown>;
  warnings: string[];
  /** Fingerprints of every slide the op touched, after the write. */
  fingerprints: Record<string, string>;
  createdSlideIds?: string[];
  deletedSlideIds?: string[];
  /** Undo snapshots the op exported before writing (the slides in `snapshotSlideIds`). */
  snapshots?: Record<string, SlideSnapshot>;
}

/** add_slide also returns the new slide's shapes (its placeholders) so the model can fill them without a read. */
export interface AddSlideReceipt extends WriteReceipt {
  shapes: Pick<ShapeInfo, "id" | "name" | "placeholder" | "left" | "top" | "width" | "height">[];
}

export type InsertFormatting = "KeepSourceFormatting" | "UseDestinationTheme";

/** Passed to every write op that edits existing slides; the op checks and exports inside its own PowerPoint.run. */
export interface WriteGuardArgs {
  /** Slide ID → fingerprint the model last read. A slide that changed since fails the write before anything changes. */
  expectedFingerprints?: Record<string, string>;
  /** Slides to export for undo before writing: the ones this turn hasn't snapshotted yet. */
  snapshotSlideIds?: string[];
}

// ---------------------------------------------------------------------------
// Write op arguments. Tool schemas (src/tools) must stay assignable to these.
// ---------------------------------------------------------------------------

export type HorizontalAlign = "Left" | "Center" | "Right" | "Justify";
export type ZOrder = "BringToFront" | "SendToBack" | "BringForward" | "SendBackward";

export interface FontStyle {
  name?: string;
  size?: number;
  /** "#RRGGBB" */
  color?: string;
  bold?: boolean;
  italic?: boolean;
  underline?: boolean;
}

export interface TextRun extends FontStyle {
  text: string;
}

export interface ShapeStyle {
  name?: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  rotation?: number;
  /** Replaces the shape's text. */
  text?: string;
  /** Replaces the shape's text with formatted runs (use instead of `text`). */
  runs?: TextRun[];
  /** Applies to all of the shape's text. */
  font?: FontStyle;
  paragraph?: { align?: HorizontalAlign; bullets?: boolean; indentLevel?: number };
  textFrame?: {
    autoSize?: "AutoSizeNone" | "AutoSizeTextToFitShape" | "AutoSizeShapeToFitText";
    verticalAlign?: "Top" | "Middle" | "Bottom";
    wordWrap?: boolean;
  };
  /** color "none" removes the fill. */
  fill?: { color: string; transparency?: number };
  line?: {
    color?: string;
    weight?: number;
    dash?: "Solid" | "Dash" | "RoundDot" | "LongDash" | "DashDot";
    visible?: boolean;
  };
  /** Links the whole shape, or the first occurrence of `text` inside it. */
  hyperlink?: { address: string; screenTip?: string; text?: string };
}

export interface ShapeUpdate extends ShapeStyle {
  shapeId: string;
  delete?: boolean;
  zOrder?: ZOrder;
}

export interface SlideShapeUpdates {
  slideId: string;
  updates: ShapeUpdate[];
}

export interface AddShapeArgs extends ShapeStyle {
  slideId: string;
  kind: "textbox" | "geometric" | "line" | "image";
  /** GeometricShapeType name, e.g. "Rectangle", "RoundRectangle", "Ellipse", "RightArrow". */
  geometry?: string;
  connector?: "Straight" | "Elbow" | "Curve";
  left: number;
  top: number;
  width: number;
  height: number;
  /** Image bytes (no data: prefix) for kind "image". */
  imageBase64?: string;
}

export interface TableCellEdit {
  row: number;
  column: number;
  text?: string;
  font?: FontStyle;
  /** "#RRGGBB" */
  fill?: string;
  align?: HorizontalAlign;
}

export interface EditTableArgs {
  slideId: string;
  /** Existing table shape. Omit when creating. */
  shapeId?: string;
  create?: {
    rows: number;
    columns: number;
    left: number;
    top: number;
    width: number;
    height: number;
    values?: string[][];
    /** TableStyle name, e.g. "MediumStyle2Accent1", "LightStyle1", "NoStyleTableGrid". */
    style?: string;
    name?: string;
  };
  insertRows?: { index: number; count: number };
  insertColumns?: { index: number; count: number };
  deleteRows?: number[];
  deleteColumns?: number[];
  cells?: TableCellEdit[];
  merge?: { row: number; column: number; rowCount: number; columnCount: number }[];
  columnWidths?: number[];
}
