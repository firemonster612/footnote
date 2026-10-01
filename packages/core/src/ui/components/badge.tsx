import type { ComponentProps } from "react";
import { cn } from "../lib/utils.ts";

/** A compact outlined chip, e.g. an attachment name. */
export function Badge({ className, ...props }: ComponentProps<"span">) {
  return (
    <span
      data-slot="badge"
      className={cn(
        "inline-flex h-6 max-w-full min-w-0 items-center gap-1 rounded-sm border bg-background px-1.5 text-caption text-muted-foreground [&_svg]:shrink-0",
        className,
      )}
      {...props}
    />
  );
}
