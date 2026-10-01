import { useId, type ReactNode } from "react";

/** A labelled form row; passes the generated input id to its child. */
export function Field({ label, children }: { label: string; children: (id: string) => ReactNode }) {
  const id = useId();
  return (
    <div className="flex flex-col gap-1">
      <label
        htmlFor={id}
        className="text-[12px] font-medium text-neutral-700 dark:text-neutral-300"
      >
        {label}
      </label>
      {children(id)}
    </div>
  );
}
