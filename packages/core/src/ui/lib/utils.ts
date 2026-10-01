import { clsx, type ClassValue } from "clsx";
import { extendTailwindMerge } from "tailwind-merge";

// Teach tailwind-merge Footnote's type scale (styles.css) so `text-small` isn't mistaken for a color.
const twMerge = extendTailwindMerge({
  extend: { classGroups: { "font-size": [{ text: ["caption", "small", "body", "title"] }] } },
});

export function cn(...inputs: ClassValue[]): string {
  return twMerge(clsx(inputs));
}
