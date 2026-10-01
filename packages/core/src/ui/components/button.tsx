import { cva, type VariantProps } from "class-variance-authority";
import type { ComponentProps } from "react";
import { cn } from "../lib/utils.ts";

export const buttonVariants = cva(
  "inline-flex shrink-0 cursor-default items-center justify-center gap-1.5 rounded-md font-medium whitespace-nowrap transition-[color,background-color,border-color] duration-100 select-none disabled:pointer-events-none disabled:opacity-45 [&_svg]:pointer-events-none [&_svg]:shrink-0",
  {
    variants: {
      variant: {
        primary:
          "bg-accent text-accent-foreground shadow-control hover:bg-accent-hover active:bg-accent-pressed",
        secondary:
          "border border-input bg-background shadow-control hover:border-input-hover hover:bg-highlight active:bg-pressed",
        ghost: "text-muted-foreground hover:bg-highlight hover:text-foreground active:bg-pressed",
        danger:
          "border border-input bg-background text-danger shadow-control hover:border-danger/50 hover:bg-danger-subtle active:bg-danger-subtle",
      },
      size: {
        sm: "h-7 px-2.5 text-small",
        md: "h-8 px-3 text-small",
        icon: "size-7",
      },
    },
    defaultVariants: { variant: "secondary", size: "sm" },
  },
);

export function Button({
  className,
  variant,
  size,
  type = "button",
  ...props
}: ComponentProps<"button"> & VariantProps<typeof buttonVariants>) {
  return (
    <button
      type={type}
      data-slot="button"
      className={cn(buttonVariants({ variant, size }), className)}
      {...props}
    />
  );
}
