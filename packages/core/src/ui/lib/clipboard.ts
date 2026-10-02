import { useEffect, useState } from "react";

const copiedFeedbackMs = 1_500;

/** Copies text, returning whether it worked. */
export async function copyText(text: string): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    // Clipboard API refuses when the panel isn't focused; the selection-based copy still works there.
    const area = document.createElement("textarea");
    area.value = text;
    area.style.position = "fixed";
    area.style.opacity = "0";
    document.body.append(area);
    area.select();
    const copied = document.execCommand("copy");
    area.remove();
    return copied;
  }
}

/** `copied` turns true for a moment after a successful `copy`. */
export function useCopy(): { copied: boolean; copy: (text: string) => void } {
  const [copied, setCopied] = useState(false);
  useEffect(() => {
    if (!copied) return;
    const timer = setTimeout(() => setCopied(false), copiedFeedbackMs);
    return () => clearTimeout(timer);
  }, [copied]);
  return { copied, copy: (text) => void copyText(text).then(setCopied) };
}
