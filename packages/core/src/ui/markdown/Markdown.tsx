import { useMemo, type ReactNode } from "react";
import { CodeBlock } from "../components/CodeBlock.tsx";
import { parseMarkdown, type Block, type Inline } from "./parseMarkdown.ts";

export function Markdown({ text }: { text: string }) {
  const blocks = useMemo(() => parseMarkdown(text), [text]);
  return <div className="flex flex-col gap-2 break-words">{renderBlocks(blocks)}</div>;
}

function renderBlocks(blocks: Block[]): ReactNode[] {
  return blocks.map((block, index) => <BlockView key={index} block={block} />);
}

const headingClass = ["text-title font-semibold", "text-title font-semibold", "font-semibold"];

function BlockView({ block }: { block: Block }) {
  switch (block.type) {
    case "paragraph":
      return <p>{renderInline(block.children)}</p>;
    case "heading":
      return (
        <p
          role="heading"
          aria-level={block.level}
          className={headingClass[block.level - 1] ?? "font-semibold"}
        >
          {renderInline(block.children)}
        </p>
      );
    case "code":
      return <CodeBlock code={block.text} />;
    case "quote":
      return (
        <blockquote className="flex flex-col gap-2 border-l-2 border-input pl-3 text-muted-foreground">
          {renderBlocks(block.children)}
        </blockquote>
      );
    case "list": {
      const items = block.items.map((item, index) => (
        <li key={index} className="pl-0.5">
          <div className="flex flex-col gap-1">{renderBlocks(item)}</div>
        </li>
      ));
      return block.ordered ? (
        <ol start={block.start} className="list-decimal space-y-1 pl-5">
          {items}
        </ol>
      ) : (
        <ul className="list-disc space-y-1 pl-5">{items}</ul>
      );
    }
    case "table":
      return (
        <div className="overflow-x-auto">
          <table className="w-full border-collapse text-small">
            <thead>
              <tr>
                {block.header.map((cell, index) => (
                  <th
                    key={index}
                    className="border-b border-input px-2 py-1 text-left font-semibold"
                  >
                    {renderInline(cell)}
                  </th>
                ))}
              </tr>
            </thead>
            <tbody>
              {block.rows.map((row, rowIndex) => (
                <tr key={rowIndex}>
                  {row.map((cell, index) => (
                    <td key={index} className="border-b px-2 py-1 align-top">
                      {renderInline(cell)}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      );
    case "rule":
      return <hr />;
    default:
      return block satisfies never;
  }
}

function renderInline(nodes: Inline[]): ReactNode[] {
  return nodes.map((node, index) => {
    switch (node.type) {
      case "text":
        return node.text;
      case "code":
        return (
          <code key={index} className="rounded-sm bg-muted px-1 py-px font-mono text-small">
            {node.text}
          </code>
        );
      case "strong":
        return <strong key={index}>{renderInline(node.children)}</strong>;
      case "em":
        return <em key={index}>{renderInline(node.children)}</em>;
      case "del":
        return <del key={index}>{renderInline(node.children)}</del>;
      case "link":
        return (
          <a
            key={index}
            href={node.href}
            target="_blank"
            rel="noopener noreferrer"
            className="text-accent-text underline underline-offset-2"
          >
            {renderInline(node.children)}
          </a>
        );
      default:
        return node satisfies never;
    }
  });
}
