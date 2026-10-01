import type { ShapeInfo } from "./types.ts";

/** Hash of shape IDs, bounds (to 0.1pt) and text, including group children. Changes when anything visible to the model moves or rewords. */
export function fingerprintShapes(shapes: ShapeInfo[]): string {
  return fnv1a(shapes.map(shapeKey).join("|"));
}

function shapeKey(shape: ShapeInfo): string {
  const bounds = [shape.left, shape.top, shape.width, shape.height]
    .map((value) => Math.round(value * 10))
    .join(",");
  const children = shape.children ? `[${shape.children.map(shapeKey).join("|")}]` : "";
  const table = shape.table ? JSON.stringify(shape.table.values) : "";
  return `${shape.id}:${bounds}:${shape.text ?? ""}${table}${children}`;
}

function fnv1a(input: string): string {
  let hash = 0x811c9dc5;
  for (let i = 0; i < input.length; i++) {
    hash ^= input.charCodeAt(i);
    hash = Math.imul(hash, 0x01000193);
  }
  return (hash >>> 0).toString(36);
}
