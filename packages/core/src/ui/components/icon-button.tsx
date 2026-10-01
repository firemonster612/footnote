import type { LucideIcon } from "lucide-react";
import type { ComponentProps } from "react";
import { Button } from "./button.tsx";
import { Tooltip, TooltipContent, TooltipTrigger } from "./tooltip.tsx";

/** A square ghost button whose label is its accessible name and its tooltip. */
export function IconButton({
  icon: Icon,
  label,
  ...props
}: ComponentProps<typeof Button> & { icon: LucideIcon; label: string }) {
  return (
    <Tooltip>
      <TooltipTrigger asChild>
        <Button variant="ghost" size="icon" aria-label={label} {...props}>
          <Icon size={16} />
        </Button>
      </TooltipTrigger>
      <TooltipContent>{label}</TooltipContent>
    </Tooltip>
  );
}
