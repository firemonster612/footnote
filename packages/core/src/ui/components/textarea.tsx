import type { ComponentProps } from "react";
import { cn } from "../lib/utils.ts";
import { fieldClass } from "./input.tsx";

export function Textarea({ className, ...props }: ComponentProps<"textarea">) {
  return (
    <textarea
      data-slot="textarea"
      className={cn(fieldClass, "min-h-16 resize-y px-2.5 py-1.5", className)}
      {...props}
    />
  );
}
