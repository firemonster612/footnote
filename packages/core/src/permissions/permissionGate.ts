import type { AgentToolCall, BeforeToolCallResult } from "@earendil-works/pi-agent-core";
import type {
  ApprovalDecision,
  ApprovalRequest,
  FootnoteTool,
  PermissionMode,
  ToolAccess,
} from "../contracts.ts";

export function needsApproval(
  access: ToolAccess,
  mode: PermissionMode,
  writesAllowed: boolean,
): boolean {
  if (mode === "full") return false;
  switch (access) {
    case "read":
    case "web":
      return false;
    case "write":
      return !writesAllowed;
    case "code":
      return true;
  }
}

export function denialReason(toolName: string, comment: string | undefined): string {
  const feedback = comment?.trim()
    ? ` Their comment: "${comment.trim()}"`
    : " They left no comment.";
  return `The user denied this ${toolName} call.${feedback} Do not retry the same call unchanged: follow their comment, change your approach, or ask them what they want.`;
}

export interface PermissionGateState {
  mode: PermissionMode;
  /** Set when the user picks "Allow writes for this chat". */
  writesAllowed: boolean;
}

export interface PermissionGate {
  state(): PermissionGateState;
  setMode(mode: PermissionMode): void;
  pendingApprovals(): ApprovalRequest[];
  /** For Pi's beforeToolCall: waits for the user's decision when the tool needs approval. */
  check(
    tool: FootnoteTool,
    toolCall: AgentToolCall,
    args: unknown,
    signal?: AbortSignal,
  ): Promise<BeforeToolCallResult | undefined>;
  resolveApproval(approvalId: string, decision: ApprovalDecision): void;
}

/** `onChange` runs whenever the mode, the chat-wide write allowance, or the pending approvals change. */
export function createPermissionGate(
  initial: PermissionGateState,
  onChange: () => void,
): PermissionGate {
  const state = { ...initial };
  const pending = new Map<
    string,
    { request: ApprovalRequest; settle: (decision: ApprovalDecision | "cancelled") => void }
  >();

  function ask(
    request: ApprovalRequest,
    signal?: AbortSignal,
  ): Promise<ApprovalDecision | "cancelled"> {
    if (signal?.aborted) return Promise.resolve("cancelled");
    return new Promise((resolve) => {
      const onAbort = () => settle("cancelled");
      const settle = (decision: ApprovalDecision | "cancelled") => {
        signal?.removeEventListener("abort", onAbort);
        pending.delete(request.id);
        resolve(decision);
        onChange();
      };
      signal?.addEventListener("abort", onAbort, { once: true });
      pending.set(request.id, { request, settle });
      onChange();
    });
  }

  return {
    state: () => ({ ...state }),
    setMode(mode) {
      state.mode = mode;
      onChange();
    },
    pendingApprovals: () => [...pending.values()].map((entry) => entry.request),

    async check(tool, toolCall, args, signal) {
      if (!needsApproval(tool.access, state.mode, state.writesAllowed)) return undefined;
      const code = tool.access === "code" ? codeOf(args) : undefined;
      const decision = await ask(
        {
          id: crypto.randomUUID(),
          toolCallId: toolCall.id,
          toolName: tool.name,
          access: tool.access,
          summary: tool.describeCall?.(args) ?? tool.label,
          ...(code !== undefined && { code }),
          args,
        },
        signal,
      );
      // Cancellation only happens when the run's signal aborts, and Pi reports aborted calls itself.
      if (decision === "cancelled") return { block: true };
      if (!decision.allow)
        return { block: true, reason: denialReason(tool.name, decision.comment) };
      if (decision.scope === "chat" && tool.access === "write") {
        state.writesAllowed = true;
        onChange();
      }
      return undefined;
    },

    resolveApproval(approvalId, decision) {
      pending.get(approvalId)?.settle(decision);
    },
  };
}

function codeOf(args: unknown): string | undefined {
  if (typeof args !== "object" || args === null || !("code" in args)) return undefined;
  return typeof args.code === "string" ? args.code : undefined;
}
