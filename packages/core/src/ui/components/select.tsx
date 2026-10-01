import { cva, type VariantProps } from "class-variance-authority";
import { Check, ChevronDown } from "lucide-react";
import { Select as SelectPrimitive } from "radix-ui";
import type { ComponentProps } from "react";
import { cn } from "../lib/utils.ts";
import { HighlightGroup, highlightItem } from "./highlight-group.tsx";
import { fieldClass } from "./input.tsx";

export const Select = SelectPrimitive.Root;
export const SelectValue = SelectPrimitive.Value;

const triggerVariants = cva(
  "flex min-w-0 cursor-default items-center justify-between gap-1 text-left whitespace-nowrap select-none disabled:pointer-events-none disabled:opacity-45 data-[placeholder]:text-subtle-foreground [&>span]:truncate",
  {
    variants: {
      variant: {
        field: cn(fieldClass, "h-8 pr-2 pl-2.5 data-[state=open]:border-ring"),
        ghost:
          "h-7 rounded-md px-2 text-small text-muted-foreground transition-colors duration-100 hover:bg-highlight hover:text-foreground data-[state=open]:bg-highlight data-[state=open]:text-foreground",
      },
    },
    defaultVariants: { variant: "field" },
  },
);

export function SelectTrigger({
  className,
  variant,
  children,
  ...props
}: ComponentProps<typeof SelectPrimitive.Trigger> & VariantProps<typeof triggerVariants>) {
  return (
    <SelectPrimitive.Trigger
      data-slot="select-trigger"
      className={cn(triggerVariants({ variant }), className)}
      {...props}
    >
      {children}
      <SelectPrimitive.Icon asChild>
        <ChevronDown size={14} className="shrink-0 text-subtle-foreground" />
      </SelectPrimitive.Icon>
    </SelectPrimitive.Trigger>
  );
}

export function SelectContent({
  className,
  children,
  ...props
}: ComponentProps<typeof SelectPrimitive.Content>) {
  return (
    <SelectPrimitive.Portal>
      <SelectPrimitive.Content
        data-slot="select-content"
        position="popper"
        sideOffset={4}
        collisionPadding={8}
        className={cn(
          "z-50 max-h-[min(20rem,var(--radix-select-content-available-height))] max-w-(--radix-select-content-available-width) min-w-(--radix-select-trigger-width) origin-(--radix-select-content-transform-origin) overflow-hidden rounded-lg bg-popover text-foreground shadow-menu data-[state=closed]:animate-menu-out data-[state=open]:animate-menu-in",
          className,
        )}
        {...props}
      >
        <SelectPrimitive.Viewport className="p-1">
          <HighlightGroup className="flex flex-col gap-px">{children}</HighlightGroup>
        </SelectPrimitive.Viewport>
      </SelectPrimitive.Content>
    </SelectPrimitive.Portal>
  );
}

/** A menu row. The sliding highlight marks hover and keyboard focus; the checked row is teal with a check. */
export function SelectItem({
  className,
  children,
  ...props
}: ComponentProps<typeof SelectPrimitive.Item>) {
  return (
    <SelectPrimitive.Item
      data-slot="select-item"
      {...highlightItem}
      className={cn(
        "flex h-7 min-w-0 cursor-default items-center gap-2 rounded-sm pr-2 pl-2 text-body outline-none select-none data-[disabled]:opacity-45 data-[state=checked]:font-medium data-[state=checked]:text-accent-text",
        className,
      )}
      {...props}
    >
      <SelectPrimitive.ItemText className="min-w-0 truncate">{children}</SelectPrimitive.ItemText>
      <SelectPrimitive.ItemIndicator className="ml-auto flex shrink-0">
        <Check size={14} />
      </SelectPrimitive.ItemIndicator>
    </SelectPrimitive.Item>
  );
}
