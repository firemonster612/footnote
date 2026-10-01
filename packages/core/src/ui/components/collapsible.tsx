import { Collapsible as CollapsiblePrimitive } from "radix-ui";
import type { ComponentProps } from "react";
import { cn } from "../lib/utils.ts";

export const Collapsible = CollapsiblePrimitive.Root;
export const CollapsibleTrigger = CollapsiblePrimitive.Trigger;

/** Chevron for a trigger carrying `group`: points right, turns down while open. */
export const disclosureChevron =
  "shrink-0 transition-transform duration-150 group-data-[state=open]:rotate-90";

export function CollapsibleContent({
  className,
  ...props
}: ComponentProps<typeof CollapsiblePrimitive.Content>) {
  return (
    <CollapsiblePrimitive.Content
      data-slot="collapsible-content"
      className={cn(
        "overflow-hidden data-[state=closed]:animate-collapsible-up data-[state=open]:animate-collapsible-down",
        className,
      )}
      {...props}
    />
  );
}
