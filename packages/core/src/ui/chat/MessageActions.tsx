import { Check, Copy, RotateCcw } from "lucide-react";
import { useEffect, useState } from "react";
import { Tooltip, TooltipContent, TooltipTrigger } from "../components/tooltip.tsx";
import { cn } from "../lib/utils.ts";

const CONFIRM_WINDOW_MS = 4_000;
const COPIED_FEEDBACK_MS = 1_500;

/** Copy (and, for requests, Revert) under a message. Hidden until the message is hovered or focused. */
export function MessageActions({
  text,
  align,
  onRevert,
  revertBlockedReason,
}: {
  text: string;
  align: "start" | "end";
  /** Present only for requests that can be reverted to. */
  onRevert?: () => void;
  /** Why Revert can't run right now (a response is streaming). */
  revertBlockedReason?: string;
}) {
  return (
    <div
      className={cn(
        "flex gap-0.5 opacity-0 transition-opacity duration-100 group-hover:opacity-100 group-focus-within:opacity-100",
        align === "end" ? "justify-end" : "justify-start",
      )}
    >
      <CopyAction text={text} />
      {onRevert && <RevertAction onRevert={onRevert} blockedReason={revertBlockedReason} />}
    </div>
  );
}

function ActionButton({
  label,
  tooltip,
  onClick,
  disabled,
  tone = "default",
  children,
}: {
  label: string;
  tooltip: string;
  onClick: () => void;
  disabled?: boolean;
  tone?: "default" | "danger";
  children: React.ReactNode;
}) {
  return (
    <Tooltip>
      <TooltipTrigger asChild>
        {/* aria-disabled instead of disabled so the tooltip still explains why. */}
        <button
          type="button"
          aria-label={label}
          aria-disabled={disabled || undefined}
          onClick={() => !disabled && onClick()}
          className={cn(
            "inline-flex h-6 items-center gap-1 rounded-sm px-1.5 text-caption text-subtle-foreground transition-colors hover:bg-highlight hover:text-foreground focus-visible:outline-2 focus-visible:outline-ring aria-disabled:opacity-50 aria-disabled:hover:bg-transparent",
            tone === "danger" && "text-danger hover:text-danger",
          )}
        >
          {children}
        </button>
      </TooltipTrigger>
      <TooltipContent>{tooltip}</TooltipContent>
    </Tooltip>
  );
}

function CopyAction({ text }: { text: string }) {
  const [copied, setCopied] = useState(false);
  useEffect(() => {
    if (!copied) return;
    const timer = setTimeout(() => setCopied(false), COPIED_FEEDBACK_MS);
    return () => clearTimeout(timer);
  }, [copied]);

  return (
    <ActionButton
      label={copied ? "Copied" : "Copy"}
      tooltip={copied ? "Copied" : "Copy text"}
      onClick={() => void copyText(text).then(() => setCopied(true))}
    >
      {copied ? <Check size={12} /> : <Copy size={12} />}
      {copied ? "Copied" : "Copy"}
    </ActionButton>
  );
}

/** Revert removes later messages and undoes their slide changes, so it takes a second click to confirm. */
function RevertAction({
  onRevert,
  blockedReason,
}: {
  onRevert: () => void;
  blockedReason?: string;
}) {
  const [armed, setArmed] = useState(false);
  useEffect(() => {
    if (!armed) return;
    const timer = setTimeout(() => setArmed(false), CONFIRM_WINDOW_MS);
    return () => clearTimeout(timer);
  }, [armed]);

  return (
    <ActionButton
      label={armed ? "Confirm revert" : "Revert"}
      tooltip={
        blockedReason ??
        (armed
          ? "Click again to remove this request and everything after it, and undo their slide changes"
          : "Go back to before this request and edit it")
      }
      disabled={blockedReason !== undefined}
      tone={armed ? "danger" : "default"}
      onClick={() => {
        if (!armed) return setArmed(true);
        setArmed(false);
        onRevert();
      }}
    >
      <RotateCcw size={12} />
      {armed ? "Confirm revert" : "Revert"}
    </ActionButton>
  );
}

async function copyText(text: string): Promise<void> {
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    // Clipboard API refuses when the panel isn't focused; the selection-based copy still works there.
    const area = document.createElement("textarea");
    area.value = text;
    area.style.position = "fixed";
    area.style.opacity = "0";
    document.body.append(area);
    area.select();
    document.execCommand("copy");
    area.remove();
  }
}
