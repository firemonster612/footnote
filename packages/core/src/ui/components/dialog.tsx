import { X } from "lucide-react";
import { Dialog as DialogPrimitive } from "radix-ui";
import type { ComponentProps } from "react";
import { cn } from "../lib/utils.ts";

export const Dialog = DialogPrimitive.Root;
export const DialogTrigger = DialogPrimitive.Trigger;
export const DialogTitle = DialogPrimitive.Title;

/** Fills the panel over a dimmed backdrop, with a close button in the corner. */
export function DialogContent({
  className,
  children,
  ...props
}: ComponentProps<typeof DialogPrimitive.Content>) {
  return (
    <DialogPrimitive.Portal>
      <DialogPrimitive.Overlay className="fixed inset-0 z-50 bg-black/70 data-[state=closed]:animate-menu-out data-[state=open]:animate-menu-in" />
      <DialogPrimitive.Content
        data-slot="dialog-content"
        className={cn(
          "fixed inset-0 z-50 flex items-center justify-center p-3 outline-none data-[state=closed]:animate-menu-out data-[state=open]:animate-menu-in",
          className,
        )}
        {...props}
      >
        {children}
        <DialogPrimitive.Close
          aria-label="Close"
          className="absolute top-2 right-2 inline-flex size-7 items-center justify-center rounded-md text-white/85 transition-colors hover:bg-white/15 hover:text-white"
        >
          <X size={18} />
        </DialogPrimitive.Close>
      </DialogPrimitive.Content>
    </DialogPrimitive.Portal>
  );
}
