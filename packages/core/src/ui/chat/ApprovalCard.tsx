import { ShieldCheck } from "lucide-react";
import { useState } from "react";
import type { ApprovalDecision, ApprovalRequest } from "../../contracts.ts";
import { CodeBlock } from "../components/CodeBlock.tsx";
import { Button } from "../components/button.tsx";
import { Input } from "../components/input.tsx";

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
      className="flex flex-col gap-2 rounded-lg border border-warning-border bg-warning-subtle p-2.5"
    >
      <div className="flex items-start gap-2">
        <ShieldCheck size={15} className="mt-0.5 shrink-0 text-warning" />
        <span className="min-w-0 flex-1 font-medium">{request.summary}</span>
      </div>
      {request.code !== undefined && <CodeBlock code={request.code} />}
      <Input
        value={comment}
        onChange={(event) => setComment(event.target.value)}
        placeholder="Comment for the model (optional)"
        aria-label="Comment for the model"
        className="h-7 text-small"
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
