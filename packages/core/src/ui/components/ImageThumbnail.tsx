import { useState } from "react";
import { Dialog, DialogContent, DialogTitle, DialogTrigger } from "./dialog.tsx";

export interface ImageSource {
  mimeType: string;
  data: string;
}

/** A thumbnail that opens a full-panel view on click; Escape or a click closes it. */
export function ImageThumbnail({ image, alt }: { image: ImageSource; alt: string }) {
  const [open, setOpen] = useState(false);
  const src = `data:${image.mimeType};base64,${image.data}`;

  return (
    <Dialog open={open} onOpenChange={setOpen}>
      <DialogTrigger
        aria-label={`Enlarge ${alt}`}
        className="overflow-hidden rounded-md border transition-colors hover:border-input-hover"
      >
        <img src={src} alt={alt} className="h-20 max-w-40 object-contain" />
      </DialogTrigger>
      <DialogContent aria-describedby={undefined} onClick={() => setOpen(false)}>
        <DialogTitle className="sr-only">{alt}</DialogTitle>
        <img src={src} alt={alt} className="max-h-full max-w-full rounded-md shadow-menu" />
      </DialogContent>
    </Dialog>
  );
}
