import { useId, type ReactNode } from "react";
import { Label } from "../components/label.tsx";

/** A labelled form row; passes the generated input id to its child. */
export function Field({ label, children }: { label: string; children: (id: string) => ReactNode }) {
  const id = useId();
  return (
    <div className="flex flex-col gap-1">
      <Label htmlFor={id}>{label}</Label>
      {children(id)}
    </div>
  );
}
