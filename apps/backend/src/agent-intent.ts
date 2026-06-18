import type { AgentRunInput, AgentRunMode } from "@components-kit/open-workbook-protocol";

export const AGENT_INTENT_ACTIONS = [
  "read_values",
  "read_schema",
  "find_target",
  "write_values",
  "write_formulas",
  "format_range",
  "clear_values",
  "append_table_rows",
  "sort_table",
  "filter_range",
  "autofit",
  "copy_template_sheet",
  "calculate",
  "save"
] as const;

export type AgentIntentAction = typeof AGENT_INTENT_ACTIONS[number];

export type AgentIntentSource = "caller_structured" | "deterministic_fallback" | "mixed";

export interface NormalizedAgentIntent {
  source: AgentIntentSource;
  action?: AgentIntentAction;
  confidence?: number;
  reason?: string;
  targetHints?: string[];
  accepted: boolean;
  rejectedReason?: string;
}

const ACTIONS = new Set<string>(AGENT_INTENT_ACTIONS);

export function normalizeAgentIntent(input: AgentRunInput): NormalizedAgentIntent {
  const raw = input.intent as { action?: unknown; confidence?: unknown; reason?: unknown; targetHints?: unknown } | undefined;
  if (!raw) {
    return { source: "deterministic_fallback", accepted: true };
  }
  if (typeof raw.action !== "string" || !ACTIONS.has(raw.action)) {
    return {
      source: "mixed",
      accepted: false,
      rejectedReason: "unsupported_intent_action"
    };
  }
  const confidence = typeof raw.confidence === "number" && Number.isFinite(raw.confidence)
    ? Math.max(0, Math.min(1, raw.confidence))
    : undefined;
  const targetHints = Array.isArray(raw.targetHints)
    ? raw.targetHints.filter((hint): hint is string => typeof hint === "string" && hint.trim().length > 0).slice(0, 8)
    : undefined;
  return {
    source: "caller_structured",
    action: raw.action as AgentIntentAction,
    ...(confidence !== undefined ? { confidence } : {}),
    ...(typeof raw.reason === "string" && raw.reason.trim() ? { reason: raw.reason.trim() } : {}),
    ...(targetHints && targetHints.length > 0 ? { targetHints } : {}),
    accepted: true
  };
}

export function modeForIntentAction(action: AgentIntentAction): AgentRunMode {
  if (action === "read_values") return "answer";
  if (action === "read_schema") return "answer";
  if (action === "find_target") return "find";
  return "preview_update";
}

export function isAgentIntentAction(value: unknown): value is AgentIntentAction {
  return typeof value === "string" && ACTIONS.has(value);
}
