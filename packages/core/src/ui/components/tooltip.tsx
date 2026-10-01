import { Tooltip as TooltipPrimitive } from "radix-ui";
import type { ComponentProps } from "react";
import { cn } from "../lib/utils.ts";

export function TooltipProvider(props: ComponentProps<typeof TooltipPrimitive.Provider>) {
  return <TooltipPrimitive.Provider delayDuration={500} skipDelayDuration={300} {...props} />;
}

export const Tooltip = TooltipPrimitive.Root;
export const TooltipTrigger = TooltipPrimitive.Trigger;

export function TooltipContent({
  className,
  sideOffset = 6,
  collisionPadding = 6,
  ...props
}: ComponentProps<typeof TooltipPrimitive.Content>) {
  return (
    <TooltipPrimitive.Portal>
      <TooltipPrimitive.Content
        data-slot="tooltip-content"
        sideOffset={sideOffset}
        collisionPadding={collisionPadding}
        className={cn(
          "z-50 max-w-60 rounded-sm bg-foreground px-2 py-1 text-caption text-background shadow-menu data-[state=closed]:animate-menu-out data-[state=delayed-open]:animate-menu-in",
          className,
        )}
        {...props}
      />
    </TooltipPrimitive.Portal>
  );
}
