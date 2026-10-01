// A small markdown subset for model output: the parser produces a tree of plain values,
// and the renderer maps it to React elements, so model text never becomes raw HTML.

export type Inline =
  | { type: "text"; text: string }
  | { type: "code"; text: string }
  | { type: "strong" | "em" | "del"; children: Inline[] }
  | { type: "link"; href: string; children: Inline[] };

export type Block =
  | { type: "paragraph"; children: Inline[] }
  | { type: "heading"; level: number; children: Inline[] }
  | { type: "code"; lang: string; text: string }
  | { type: "quote"; children: Block[] }
  | { type: "list"; ordered: boolean; start: number; items: Block[][] }
  | { type: "table"; header: Inline[][]; rows: Inline[][][] }
  | { type: "rule" };

const fencePattern = /^ {0,3}(`{3,}|~{3,})\s*([\w+-]*)/;
const headingPattern = /^ {0,3}(#{1,6})\s+(.*?)\s*#*\s*$/;
const rulePattern = /^ {0,3}([-*_])(\s*\1){2,}\s*$/;
const quotePattern = /^ {0,3}> ?/;
const listItemPattern = /^( {0,3})([-*+]|\d{1,9}[.)])\s+/;
const tableSeparatorPattern = /^\s*\|?\s*:?-+:?\s*(\|\s*:?-+:?\s*)*\|?\s*$/;

export function parseMarkdown(source: string): Block[] {
  return parseBlocks(source.replace(/\r\n?/g, "\n").split("\n"));
}

function parseBlocks(lines: string[]): Block[] {
  const blocks: Block[] = [];
  let i = 0;
  while (i < lines.length) {
    const line = lines[i]!;
    if (line.trim() === "") {
      i++;
      continue;
    }

    const fence = fencePattern.exec(line);
    if (fence) {
      const marker = fence[1]!;
      const body: string[] = [];
      i++;
      while (i < lines.length && !lines[i]!.trimStart().startsWith(marker)) body.push(lines[i++]!);
      i++; // closing fence; an unclosed fence (mid-stream) runs to the end
      blocks.push({ type: "code", lang: fence[2] ?? "", text: body.join("\n") });
      continue;
    }

    const heading = headingPattern.exec(line);
    if (heading) {
      blocks.push({
        type: "heading",
        level: heading[1]!.length,
        children: parseInline(heading[2]!),
      });
      i++;
      continue;
    }

    if (rulePattern.test(line)) {
      blocks.push({ type: "rule" });
      i++;
      continue;
    }

    if (quotePattern.test(line)) {
      const body: string[] = [];
      while (i < lines.length && quotePattern.test(lines[i]!))
        body.push(lines[i++]!.replace(quotePattern, ""));
      blocks.push({ type: "quote", children: parseBlocks(body) });
      continue;
    }

    if (listItemPattern.test(line)) {
      i = parseList(lines, i, blocks);
      continue;
    }

    if (line.includes("|") && tableSeparatorPattern.test(lines[i + 1] ?? "")) {
      const rows: Inline[][][] = [];
      const header = splitTableRow(line);
      i += 2;
      while (i < lines.length && lines[i]!.includes("|") && lines[i]!.trim() !== "")
        rows.push(splitTableRow(lines[i++]!));
      blocks.push({ type: "table", header, rows });
      continue;
    }

    const paragraph: string[] = [];
    while (i < lines.length && lines[i]!.trim() !== "" && !startsBlock(lines[i]!))
      paragraph.push(lines[i++]!.trim());
    blocks.push({ type: "paragraph", children: parseInline(paragraph.join(" ")) });
  }
  return blocks;
}

function startsBlock(line: string): boolean {
  return [fencePattern, headingPattern, rulePattern, quotePattern, listItemPattern].some(
    (pattern) => pattern.test(line),
  );
}

/** Parses one list starting at `start`, appends it to `blocks`, and returns the index after it. */
function parseList(lines: string[], start: number, blocks: Block[]): number {
  const first = listItemPattern.exec(lines[start]!)!;
  const ordered = /\d/.test(first[2]!);
  const items: Block[][] = [];
  let i = start;
  while (i < lines.length) {
    const item = listItemPattern.exec(lines[i]!);
    if (!item || /\d/.test(item[2]!) !== ordered) break;
    const contentIndent = item[0].length;
    const body = [lines[i]!.slice(contentIndent)];
    i++;
    // Continuation: indented lines (nested lists, wrapped text) and blank lines followed by indented lines.
    while (i < lines.length) {
      const next = lines[i]!;
      if (next.trim() === "") {
        if (!/^\s{2,}\S/.test(lines[i + 1] ?? "")) break;
        body.push("");
      } else if (/^\s{2,}/.test(next)) {
        body.push(next.replace(new RegExp(`^ {0,${contentIndent}}`), ""));
      } else if (!listItemPattern.test(next) && !startsBlock(next)) {
        body.push(next); // lazy continuation of the item's paragraph
      } else {
        break;
      }
      i++;
    }
    items.push(parseBlocks(body));
  }
  blocks.push({
    type: "list",
    ordered,
    start: ordered ? Number.parseInt(first[2]!, 10) : 1,
    items,
  });
  return i;
}

function splitTableRow(line: string): Inline[][] {
  return line
    .trim()
    .replace(/^\|/, "")
    .replace(/\|$/, "")
    .split(/(?<!\\)\|/)
    .map((cell) => parseInline(cell.trim().replace(/\\\|/g, "|")));
}

const inlinePattern =
  /(`+)(.+?)\1|\*\*(.+?)\*\*|__(.+?)__|~~(.+?)~~|\*(?=\S)(.+?)\*|(?<![\w])_(?=\S)(.+?)_(?![\w])|\[([^\]]+)\]\(\s*((?:[^()\s]|\([^()\s]*\))+)\s*\)/;

export function parseInline(text: string): Inline[] {
  const nodes: Inline[] = [];
  let rest = text;
  while (rest) {
    const match = inlinePattern.exec(rest);
    if (!match) {
      nodes.push({ type: "text", text: rest });
      break;
    }
    if (match.index > 0) nodes.push({ type: "text", text: rest.slice(0, match.index) });
    nodes.push(inlineNode(match));
    rest = rest.slice(match.index + match[0].length);
  }
  return nodes;
}

function inlineNode(match: RegExpExecArray): Inline {
  const [, , code, strong, strongAlt, del, em, emAlt, linkText, href] = match;
  if (code !== undefined) return { type: "code", text: code.trim() || code };
  if (strong !== undefined || strongAlt !== undefined)
    return { type: "strong", children: parseInline(strong ?? strongAlt!) };
  if (del !== undefined) return { type: "del", children: parseInline(del) };
  if (em !== undefined || emAlt !== undefined)
    return { type: "em", children: parseInline(em ?? emAlt!) };
  if (!isSafeHref(href!)) return { type: "text", text: linkText! };
  return { type: "link", href: href!, children: parseInline(linkText!) };
}

function isSafeHref(href: string): boolean {
  return /^(https?:|mailto:)/i.test(href);
}
