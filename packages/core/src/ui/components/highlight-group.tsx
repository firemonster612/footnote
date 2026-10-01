import { useRef, type ComponentProps } from "react";
import { cn } from "../lib/utils.ts";

/** Spread onto each element the enclosing HighlightGroup's indicator should track. */
export const highlightItem = { "data-highlight-item": "" };

/**
 * A set of items sharing one hover/keyboard highlight that slides between them.
 *
 * The indicator follows the item under the pointer and the item holding focus (Radix menu items
 * take focus as they become `data-highlighted`, so keyboard navigation drives it too). It appears
 * in place, slides from item to item, and fades only once the pointer has left the group and focus
 * isn't visibly inside it. Positions come from offsets, not bounding rects, so an opening menu's
 * scale animation can't skew them. `prefers-reduced-motion` zeroes the transition (styles.css).
 */
export function HighlightGroup({
  className,
  indicatorClassName,
  children,
  ...props
}: ComponentProps<"div"> & { indicatorClassName?: string }) {
  const groupRef = useRef<HTMLDivElement>(null);
  const indicatorRef = useRef<HTMLSpanElement>(null);
  const pointerInside = useRef(false);

  function moveTo(target: EventTarget) {
    const group = groupRef.current;
    const indicator = indicatorRef.current;
    if (!group || !indicator || !(target instanceof Element)) return;
    const item = target.closest<HTMLElement>("[data-highlight-item]");
    if (!item || item.closest("[data-highlight-group]") !== group) return;
    if (item.matches("[data-disabled], :disabled")) return;

    let x = 0;
    let y = 0;
    for (let node: Element | null = item; node instanceof HTMLElement && node !== group;) {
      x += node.offsetLeft;
      y += node.offsetTop;
      node = node.offsetParent;
    }
    // From hidden, appear in place; between items, slide.
    const sliding = indicator.dataset.visible === "true";
    indicator.style.transitionProperty = sliding ? "transform, width, height, opacity" : "opacity";
    indicator.style.transform = `translate(${x}px, ${y}px)`;
    indicator.style.width = `${item.offsetWidth}px`;
    indicator.style.height = `${item.offsetHeight}px`;
    indicator.dataset.visible = "true";
  }

  function hideUnlessFocusVisible() {
    const group = groupRef.current;
    const indicator = indicatorRef.current;
    if (!group || !indicator || pointerInside.current) return;
    if (group.querySelector(":focus-visible")) return;
    indicator.dataset.visible = "false";
  }

  return (
    <div
      {...props}
      ref={groupRef}
      data-highlight-group=""
      className={cn("relative isolate", className)}
      onPointerOver={(event) => {
        pointerInside.current = true;
        moveTo(event.target);
      }}
      onPointerLeave={() => {
        pointerInside.current = false;
        hideUnlessFocusVisible();
      }}
      onFocus={(event) => moveTo(event.target)}
      onBlur={(event) => {
        if (
          !(event.relatedTarget instanceof Node && groupRef.current?.contains(event.relatedTarget))
        )
          hideUnlessFocusVisible();
      }}
    >
      <span
        ref={indicatorRef}
        aria-hidden="true"
        data-visible="false"
        className={cn(
          "pointer-events-none absolute top-0 left-0 -z-10 rounded-sm bg-highlight opacity-0 duration-150 ease-[cubic-bezier(0.2,0,0,1)] data-[visible=true]:opacity-100",
          indicatorClassName,
        )}
      />
      {children}
    </div>
  );
}
