import { AGENT_INTENT_ACTIONS, type AgentIntentAction, type AgentRunInput, type AgentRunMode } from "@components-kit/open-workbook-protocol";

export { AGENT_INTENT_ACTIONS, type AgentIntentAction };

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
const ACTION_ALIASES: Record<string, AgentIntentAction> = {
  apply_style_from_template: "copy_style_from_template",
  apply_styles_from_template: "copy_style_from_template",
  copy_style: "copy_style_from_template",
  copy_styles: "copy_style_from_template",
  apply_filter: "filter_range",
  apply_filters: "filter_range",
  add_filter: "filter_range",
  add_filters: "filter_range",
  autofilter: "filter_range",
  auto_filter: "filter_range",
  add_dropdown: "write_data_validation",
  set_dropdown: "write_data_validation",
  data_validation: "write_data_validation",
  add_data_validation: "write_data_validation",
  conditional_format: "write_conditional_formatting",
  conditional_formatting: "write_conditional_formatting",
  add_conditional_formatting: "write_conditional_formatting",
  formula_format: "write_conditional_formatting",
  swap_columns: "reorder_range_columns"
};

export function normalizeAgentIntent(input: AgentRunInput): NormalizedAgentIntent {
  const raw = input.intent as { action?: unknown; confidence?: unknown; reason?: unknown; targetHints?: unknown } | undefined;
  if (!raw) {
    return { source: "deterministic_fallback", accepted: true };
  }
  const action = typeof raw.action === "string" ? ACTION_ALIASES[raw.action] ?? raw.action : undefined;
  if (typeof action !== "string" || !ACTIONS.has(action)) {
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
    source: action === raw.action ? "caller_structured" : "mixed",
    action: action as AgentIntentAction,
    ...(confidence !== undefined ? { confidence } : {}),
    ...(typeof raw.reason === "string" && raw.reason.trim() ? { reason: raw.reason.trim() } : raw.action !== action ? { reason: `Normalized alias ${String(raw.action)} to ${action}.` } : {}),
    ...(targetHints && targetHints.length > 0 ? { targetHints } : {}),
    accepted: true
  };
}

export function modeForIntentAction(action: AgentIntentAction): AgentRunMode {
  if (action === "read_values") return "answer";
  if (action === "read_schema") return "answer";
  if (action === "list_open_workbooks") return "answer";
  if (action === "get_workbook_info") return "answer";
  if (action === "refresh_workbook_snapshot") return "answer";
  if (action === "get_workbook_snapshot") return "answer";
  if (action === "detect_external_changes") return "answer";
  if (action === "export_local_config") return "answer";
  if (action === "read_embedded_local_config") return "answer";
  if (action === "read_formula_patterns") return "answer";
  if (action === "get_formula_dependency_graph") return "answer";
  if (action === "trace_formula_precedents") return "answer";
  if (action === "trace_formula_dependents") return "answer";
  if (action === "validate_formula_range") return "validate";
  if (action === "find_formula_errors") return "answer";
  if (action === "explain_formula") return "answer";
  if (action === "read_named_item") return "answer";
  if (action === "read_region") return "answer";
  if (action === "find_similar_rows") return "answer";
  if (action === "read_style_summary") return "answer";
  if (action === "format_diagnostics") return "answer";
  if (action === "read_style_fingerprint") return "answer";
  if (action === "compare_style_fingerprint") return "answer";
  if (isRangeReadIntentAction(action)) return "answer";
  if (action === "get_theme") return "answer";
  if (action === "apply_theme") return "answer";
  if (action === "repair_filters_from_template") return "answer";
  if (action === "repair_print_layout") return "answer";
  if (action === "repair_named_ranges") return "answer";
  if (action === "repair_formula_errors") return "answer";
  if (action === "repair_merged_cells") return "answer";
  if (action === "detect_header_row") return "answer";
  if (action === "detect_outliers") return "answer";
  if (action === "fuzzy_match") return "answer";
  if (action === "detect_templates") return "answer";
  if (action === "read_template") return "answer";
  if (action === "list_templates") return "answer";
  if (action === "infer_template_regions") return "answer";
  if (isWorkflowIntentAction(action)) return "answer";
  if (action === "validate_sheet_against_template") return "validate";
  if (action === "list_snapshots") return "answer";
  if (action === "read_snapshot") return "answer";
  if (action === "compare_snapshots") return "answer";
  if (action === "list_backups") return "answer";
  if (action === "read_backup") return "answer";
  if (action === "verify_backup") return "answer";
  if (action.startsWith("validate_")) return "validate";
  if (action === "find_target") return "find";
  return "preview_update";
}

export function isAgentIntentAction(value: unknown): value is AgentIntentAction {
  return typeof value === "string" && ACTIONS.has(value);
}

function isRangeReadIntentAction(action: AgentIntentAction): boolean {
  return action === "read_range_compact"
    || action === "get_range_summary"
    || action === "read_hyperlinks"
    || action === "read_comments"
    || action === "read_notes"
    || action === "read_merged_cells"
    || action === "read_data_validation"
    || action === "read_conditional_formatting"
    || action === "search_range"
    || action === "find_blank_cells"
    || action === "find_range_errors";
}

function isWorkflowIntentAction(action: AgentIntentAction): boolean {
  return action === "prepare_session"
    || action === "create_formula_sheet"
    || action === "create_template_report"
    || action === "create_pivot_chart_summary"
    || action === "preview_risky_edit"
    || action === "inspect_analyze"
    || action === "rollback_validate";
}
