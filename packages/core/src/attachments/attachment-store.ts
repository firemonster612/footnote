import type { AttachmentMeta, AttachmentStore, ProcessedAttachment } from "../contracts.ts";

export interface WritableAttachmentStore extends AttachmentStore {
  add(attachment: ProcessedAttachment): void;
  remove(id: string): void;
  clear(): void;
}

export function createAttachmentStore(): WritableAttachmentStore {
  const attachments = new Map<string, ProcessedAttachment>();
  return {
    list: () => [...attachments.values()].map(toMeta),
    get: (id) => attachments.get(id),
    add: (attachment) => void attachments.set(attachment.id, attachment),
    remove: (id) => void attachments.delete(id),
    clear: () => attachments.clear(),
  };
}

function toMeta({ id, name, mimeType, kind, size, summary }: ProcessedAttachment): AttachmentMeta {
  return { id, name, mimeType, kind, size, summary };
}
