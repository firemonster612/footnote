import { ShieldCheck } from "lucide-react";
import { useState } from "react";
import type { ApprovalDecision, ApprovalRequest } from "../../contracts.ts";
import { CodeBlock } from "../components/CodeBlock.tsx";
import { Button, TextInput } from "../components/controls.tsx";

export function ApprovalCard({
  request,
  onDecide,
}: {
  request: ApprovalRequest;
  onDecide: (decision: ApprovalDecision) => void;
}) {
  const [comment, setComment] = useState("");
  const withComment = (decision: ApprovalDecision): ApprovalDecision =>
    comment.trim() ? { ...decision, comment: comment.trim() } : decision;

  return (
    <section
      aria-label="Approval needed"
      className="flex flex-col gap-2 rounded-md border border-amber-300 bg-amber-50/60 p-2 dark:border-amber-700/70 dark:bg-amber-950/30"
    >
      <div className="flex items-start gap-2">
        <ShieldCheck size={15} className="mt-0.5 shrink-0 text-amber-700 dark:text-amber-400" />
        <span className="min-w-0 flex-1 font-medium">{request.summary}</span>
      </div>
      {request.code !== undefined && <CodeBlock code={request.code} />}
      <TextInput
        value={comment}
        onChange={(event) => setComment(event.target.value)}
        placeholder="Comment for the model (optional)"
        aria-label="Comment for the model"
        className="h-7 text-[12px]"
      />
      <div className="flex flex-wrap gap-1.5">
        <Button
          variant="primary"
          onClick={() => onDecide(withComment({ allow: true, scope: "once" }))}
        >
          Allow
        </Button>
        {request.access === "write" && (
          <Button onClick={() => onDecide(withComment({ allow: true, scope: "chat" }))}>
            Allow writes for this chat
          </Button>
        )}
        <Button variant="danger" onClick={() => onDecide(withComment({ allow: false }))}>
          Deny
        </Button>
      </div>
    </section>
  );
}
