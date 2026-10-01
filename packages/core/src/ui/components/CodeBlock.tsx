import { Check, Copy } from "lucide-react";
import { useState } from "react";
import { IconButton } from "./icon-button.tsx";

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
      <pre className="max-h-80 overflow-auto rounded-md bg-muted p-2 pr-8 font-mono text-small leading-[18px]">
        <code>{code}</code>
      </pre>
      <IconButton
        icon={copied ? Check : Copy}
        label={copied ? "Copied" : "Copy code"}
        onClick={copy}
        className="absolute top-1 right-1 size-6 bg-muted opacity-0 group-hover:opacity-100 focus-visible:opacity-100 [&_svg]:size-3.5"
      />
    </div>
  );
}
