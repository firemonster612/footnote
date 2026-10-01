import { Check, Copy } from "lucide-react";
import { useState } from "react";

const copiedResetMs = 1500;

export function CodeBlock({ code }: { code: string }) {
  const [copied, setCopied] = useState(false);

  async function copy() {
    await navigator.clipboard.writeText(code);
    setCopied(true);
    setTimeout(() => setCopied(false), copiedResetMs);
  }

  return (
    <div className="group relative">
      <pre className="max-h-80 overflow-auto rounded-md bg-neutral-100 p-2 pr-8 font-mono text-[12px] leading-[18px] dark:bg-neutral-800">
        <code>{code}</code>
      </pre>
      <button
        type="button"
        onClick={copy}
        aria-label={copied ? "Copied" : "Copy code"}
        className="absolute top-1.5 right-1.5 rounded p-1 text-neutral-500 opacity-0 group-hover:opacity-100 hover:bg-neutral-200 focus-visible:opacity-100 dark:hover:bg-neutral-700"
      >
        {copied ? <Check size={14} /> : <Copy size={14} />}
      </button>
    </div>
  );
}
