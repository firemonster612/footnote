import { read, utils, type WorkSheet } from "xlsx";
import type { ExtractedContent } from "./extracted-content.ts";

const MAX_ROWS_PER_SHEET = 500;
const MAX_COLUMNS = 50;
const MAX_CELL_CHARS = 200;

/** XLSX/XLS/ODS/CSV/TSV → one markdown table per sheet, first row as header. */
export function processSpreadsheet(bytes: ArrayBuffer): ExtractedContent {
  const workbook = read(bytes, { type: "array" });
  const text = workbook.SheetNames.map((sheetName) => {
    const sheet = workbook.Sheets[sheetName];
    return `## Sheet: ${sheetName}\n\n${sheet ? sheetToMarkdown(sheet) : "(empty)"}`;
  }).join("\n\n");
  const sheetCount = workbook.SheetNames.length;
  return { text, summary: sheetCount === 1 ? "1 sheet" : `${sheetCount} sheets` };
}

function sheetToMarkdown(sheet: WorkSheet): string {
  const rows = utils.sheet_to_json<string[]>(sheet, {
    header: 1,
    raw: false,
    defval: "",
    blankrows: false,
  });
  if (rows.length === 0) return "(empty)";

  const shownRows = rows.slice(0, MAX_ROWS_PER_SHEET + 1);
  const usedColumns = shownRows.reduce(
    (max, row) => Math.max(max, row.findLastIndex((cell) => cell !== "") + 1),
    0,
  );
  const columnCount = Math.min(usedColumns, MAX_COLUMNS);
  const lines = shownRows.map((row) => {
    const cells = Array.from({ length: columnCount }, (_, column) => formatCell(row[column] ?? ""));
    return `| ${cells.join(" | ")} |`;
  });
  lines.splice(1, 0, `|${" --- |".repeat(columnCount)}`);

  const notes: string[] = [];
  const dataRows = rows.length - 1;
  if (dataRows > MAX_ROWS_PER_SHEET)
    notes.push(`showing ${MAX_ROWS_PER_SHEET} of ${dataRows} data rows`);
  if (usedColumns > MAX_COLUMNS) notes.push(`showing ${MAX_COLUMNS} of ${usedColumns} columns`);
  if (notes.length > 0) lines.push("", `(Truncated: ${notes.join("; ")}.)`);
  return lines.join("\n");
}

function formatCell(value: string): string {
  const cell = value.replace(/\r?\n/g, " ").replaceAll("|", "\\|");
  return cell.length > MAX_CELL_CHARS ? `${cell.slice(0, MAX_CELL_CHARS)}…` : cell;
}
