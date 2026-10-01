import { Label as LabelPrimitive } from "radix-ui";
import type { ComponentProps } from "react";
import { cn } from "../lib/utils.ts";

export function Label({ className, ...props }: ComponentProps<typeof LabelPrimitive.Root>) {
  return (
    <LabelPrimitive.Root
      data-slot="label"
      className={cn("text-small font-medium text-foreground select-none", className)}
      {...props}
    />
  );
}
