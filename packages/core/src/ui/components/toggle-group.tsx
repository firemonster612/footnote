import { ToggleGroup as ToggleGroupPrimitive } from "radix-ui";
import type { ComponentProps } from "react";
import { cn } from "../lib/utils.ts";
import { HighlightGroup, highlightItem } from "./highlight-group.tsx";

/** A segmented control: the selected segment is a raised thumb, hover slides beneath it. */
export function ToggleGroup({
  className,
  children,
  ...props
}: ComponentProps<typeof ToggleGroupPrimitive.Root>) {
  return (
    <ToggleGroupPrimitive.Root
      data-slot="toggle-group"
      className={cn("inline-flex shrink-0 rounded-md bg-muted p-0.5", className)}
      {...props}
    >
      <HighlightGroup className="flex gap-0.5" indicatorClassName="bg-pressed">
        {children}
      </HighlightGroup>
    </ToggleGroupPrimitive.Root>
  );
}

export function ToggleGroupItem({
  className,
  ...props
}: ComponentProps<typeof ToggleGroupPrimitive.Item>) {
  return (
    <ToggleGroupPrimitive.Item
      data-slot="toggle-group-item"
      {...highlightItem}
      className={cn(
        "inline-flex h-6 min-w-6 cursor-default items-center justify-center gap-1 rounded-sm px-1 text-muted-foreground transition-colors duration-100 select-none hover:text-foreground aria-checked:bg-background aria-checked:text-foreground aria-checked:shadow-thumb [&_svg]:shrink-0",
        className,
      )}
      {...props}
    />
  );
}
