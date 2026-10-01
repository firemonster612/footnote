import type { ComponentProps } from "react";
import { cn } from "../lib/utils.ts";

/** Border, focus halo and hover shared by every text-entry control. */
export const fieldClass =
  "w-full min-w-0 rounded-md border border-input bg-background text-body shadow-control transition-[border-color,box-shadow] duration-100 placeholder:text-subtle-foreground hover:border-input-hover focus-visible:border-ring focus-visible:ring-2 focus-visible:ring-ring/20 focus-visible:outline-none disabled:opacity-50";

export function Input({ className, ...props }: ComponentProps<"input">) {
  return <input data-slot="input" className={cn(fieldClass, "h-8 px-2.5", className)} {...props} />;
}
