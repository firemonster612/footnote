import type { ToolCall, ToolResultMessage } from "@earendil-works/pi-ai";
import {
  Ban,
  ChevronRight,
  CircleCheck,
  CircleX,
  Hand,
  LoaderCircle,
  type LucideIcon,
} from "lucide-react";
import { useState } from "react";
import type { FootnoteTool } from "../../contracts.ts";
import { CodeBlock } from "../components/CodeBlock.tsx";
import { ImageThumbnail } from "../components/ImageThumbnail.tsx";
import type { ToolCallStatus } from "./transcript.ts";

const statusIcons: Record<ToolCallStatus, { icon: LucideIcon; className: string; label: string }> =
  {
    "awaiting-approval": {
      icon: Hand,
      className: "text-amber-600 dark:text-amber-400",
      label: "Waiting for approval",
    },
    running: { icon: LoaderCircle, className: "animate-spin text-neutral-500", label: "Running" },
    done: { icon: CircleCheck, className: "text-accent-fg", label: "Done" },
    error: { icon: CircleX, className: "text-red-600 dark:text-red-400", label: "Failed" },
    "not-run": { icon: Ban, className: "text-neutral-400", label: "Not run" },
  };

export function describeToolCall(call: ToolCall, tool: FootnoteTool | undefined): string {
  if (tool?.describeCall) {
    try {
      return tool.describeCall(call.arguments);
    } catch {
      // Arguments of a call that is still streaming can be partial; fall back to the label.
    }
  }
  return tool?.label ?? humanizeToolName(call.name);
}

function humanizeToolName(name: string): string {
  const words = name.replace(/[_-]+/g, " ").trim();
  return words.charAt(0).toUpperCase() + words.slice(1);
}

export function ToolCard({
  call,
  tool,
  result,
  status,
}: {
  call: ToolCall;
  tool: FootnoteTool | undefined;
  result: ToolResultMessage | undefined;
  status: ToolCallStatus;
}) {
  const [expanded, setExpanded] = useState(false);
  const { icon: StatusIcon, className, label } = statusIcons[status];
  const images = result?.content.filter((part) => part.type === "image") ?? [];
  const resultText = result?.content
    .flatMap((part) => (part.type === "text" ? [part.text] : []))
    .join("\n\n");
  const code = typeof call.arguments.code === "string" ? call.arguments.code : undefined;

  return (
    <div className="rounded-md border border-neutral-200 dark:border-neutral-700">
      <button
        type="button"
        onClick={() => setExpanded(!expanded)}
        aria-expanded={expanded}
        className="flex w-full items-center gap-2 px-2 py-1.5 text-left text-[12px] hover:bg-neutral-50 dark:hover:bg-neutral-800"
      >
        <StatusIcon size={14} className={`shrink-0 ${className}`} aria-label={label} />
        <span className="min-w-0 flex-1 truncate">{describeToolCall(call, tool)}</span>
        <ChevronRight
          size={14}
          className={`shrink-0 text-neutral-400 transition-transform ${expanded ? "rotate-90" : ""}`}
        />
      </button>
      {images.length > 0 && (
        <div className="flex flex-wrap gap-1.5 px-2 pb-2">
          {images.map((image, index) => (
            <ImageThumbnail key={index} image={image} alt={`${call.name} image ${index + 1}`} />
          ))}
        </div>
      )}
      {expanded && (
        <div className="flex flex-col gap-1.5 border-t border-neutral-200 p-2 dark:border-neutral-700">
          <span className="font-mono text-[11px] text-neutral-500">{call.name}</span>
          {code !== undefined ? (
            <CodeBlock code={code} />
          ) : (
            <CodeBlock code={JSON.stringify(call.arguments, null, 2)} />
          )}
          {resultText && (
            <pre
              className={`max-h-60 overflow-auto rounded-md p-2 font-mono text-[11px] leading-4 whitespace-pre-wrap ${
                result?.isError
                  ? "bg-red-50 text-red-800 dark:bg-red-950 dark:text-red-300"
                  : "bg-neutral-50 dark:bg-neutral-800/60"
              }`}
            >
              {resultText}
            </pre>
          )}
        </div>
      )}
    </div>
  );
}
