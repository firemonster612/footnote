import { X } from "lucide-react";
import { useEffect, useState } from "react";

export interface ImageSource {
  mimeType: string;
  data: string;
}

/** A thumbnail that opens a full-panel view on click; Escape or a click closes it. */
export function ImageThumbnail({ image, alt }: { image: ImageSource; alt: string }) {
  const [open, setOpen] = useState(false);
  const src = `data:${image.mimeType};base64,${image.data}`;

  useEffect(() => {
    if (!open) return;
    const closeOnEscape = (event: KeyboardEvent) => event.key === "Escape" && setOpen(false);
    window.addEventListener("keydown", closeOnEscape);
    return () => window.removeEventListener("keydown", closeOnEscape);
  }, [open]);

  return (
    <>
      <button
        type="button"
        onClick={() => setOpen(true)}
        aria-label={`Enlarge ${alt}`}
        className="overflow-hidden rounded border border-neutral-200 hover:border-neutral-400 dark:border-neutral-700"
      >
        <img src={src} alt={alt} className="h-20 max-w-40 object-contain" />
      </button>
      {open && (
        <div
          role="dialog"
          aria-modal="true"
          aria-label={alt}
          onClick={() => setOpen(false)}
          className="fixed inset-0 z-50 flex items-center justify-center bg-black/70 p-3"
        >
          <img src={src} alt={alt} className="max-h-full max-w-full rounded shadow-lg" />
          <button
            type="button"
            autoFocus
            aria-label="Close"
            className="absolute top-2 right-2 rounded p-1 text-white hover:bg-white/15"
          >
            <X size={18} />
          </button>
        </div>
      )}
    </>
  );
}
