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
import type { FootnoteTool } from "../../contracts.ts";
import { CodeBlock } from "../components/CodeBlock.tsx";
import {
  Collapsible,
  CollapsibleContent,
  CollapsibleTrigger,
  disclosureChevron,
} from "../components/collapsible.tsx";
import { ImageThumbnail } from "../components/ImageThumbnail.tsx";
import { cn } from "../lib/utils.ts";
import type { ToolCallStatus } from "./transcript.ts";

const statusIcons: Record<ToolCallStatus, { icon: LucideIcon; className: string; label: string }> =
  {
    "awaiting-approval": {
      icon: Hand,
      className: "text-warning",
      label: "Waiting for approval",
    },
    running: {
      icon: LoaderCircle,
      className: "animate-spin text-muted-foreground",
      label: "Running",
    },
    done: { icon: CircleCheck, className: "text-accent-text", label: "Done" },
    error: { icon: CircleX, className: "text-danger", label: "Failed" },
    "not-run": { icon: Ban, className: "text-subtle-foreground", label: "Not run" },
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
  const { icon: StatusIcon, className, label } = statusIcons[status];
  const images = result?.content.filter((part) => part.type === "image") ?? [];
  const resultText = result?.content
    .flatMap((part) => (part.type === "text" ? [part.text] : []))
    .join("\n\n");
  const code = typeof call.arguments.code === "string" ? call.arguments.code : undefined;

  return (
    <Collapsible className="overflow-hidden rounded-md border">
      <CollapsibleTrigger className="group flex h-8 w-full cursor-default items-center gap-2 px-2 text-left text-small transition-colors duration-100 select-none hover:bg-highlight focus-visible:outline-offset-[-2px]">
        <StatusIcon size={14} className={cn("shrink-0", className)} aria-label={label} />
        <span className="min-w-0 flex-1 truncate">{describeToolCall(call, tool)}</span>
        <ChevronRight size={14} className={cn(disclosureChevron, "text-subtle-foreground")} />
      </CollapsibleTrigger>
      {images.length > 0 && (
        <div className="flex flex-wrap gap-1.5 px-2 pb-2">
          {images.map((image, index) => (
            <ImageThumbnail key={index} image={image} alt={`${call.name} image ${index + 1}`} />
          ))}
        </div>
      )}
      <CollapsibleContent>
        <div className="flex flex-col gap-1.5 border-t p-2">
          <span className="font-mono text-caption text-muted-foreground">{call.name}</span>
          {code !== undefined ? (
            <CodeBlock code={code} />
          ) : (
            <CodeBlock code={JSON.stringify(call.arguments, null, 2)} />
          )}
          {resultText && (
            <pre
              className={cn(
                "max-h-60 overflow-auto rounded-md p-2 font-mono text-caption leading-4 whitespace-pre-wrap",
                result?.isError ? "bg-danger-subtle text-danger" : "bg-muted",
              )}
            >
              {resultText}
            </pre>
          )}
        </div>
      </CollapsibleContent>
    </Collapsible>
  );
}
