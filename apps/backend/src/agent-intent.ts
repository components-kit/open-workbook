import type { AgentRunInput, AgentRunMode } from "@components-kit/open-workbook-protocol";

export const AGENT_INTENT_ACTIONS = [
  "read_values",
  "read_schema",
  "list_open_workbooks",
  "get_workbook_info",
  "refresh_workbook_snapshot",
  "get_workbook_snapshot",
  "detect_external_changes",
  "restore_workbook_backup",
  "export_local_config",
  "import_local_config",
  "embed_local_config",
  "read_embedded_local_config",
  "import_embedded_local_config",
  "close_workbook",
  "prepare_session",
  "create_formula_sheet",
  "create_template_report",
  "create_pivot_chart_summary",
  "preview_risky_edit",
  "inspect_analyze",
  "rollback_validate",
  "read_formula_patterns",
  "get_formula_dependency_graph",
  "trace_formula_precedents",
  "trace_formula_dependents",
  "validate_formula_range",
  "validate_formula_against_template",
  "find_formula_errors",
  "explain_formula",
  "copy_formula_patterns",
  "fill_formula_down",
  "fill_formula_right",
  "repair_formula_patterns",
  "convert_formulas_to_values",
  "recalculate_formulas",
  "read_named_item",
  "create_name",
  "update_name",
  "delete_name",
  "read_region",
  "register_region",
  "clear_region_values",
  "write_region_values",
  "fill_region",
  "find_target",
  "write_values",
  "write_formulas",
  "write_number_formats",
  "format_range",
  "read_range_compact",
  "get_range_summary",
  "read_hyperlinks",
  "read_comments",
  "read_notes",
  "read_merged_cells",
  "read_data_validation",
  "read_conditional_formatting",
  "search_range",
  "find_blank_cells",
  "find_range_errors",
  "write_styles_many",
  "read_style_fingerprint",
  "compare_style_fingerprint",
  "get_theme",
  "apply_theme",
  "copy_style_from_template",
  "repair_style_consistency",
  "repair_style_from_template",
  "repair_formulas_from_template",
  "repair_filters_from_template",
  "repair_table_structure",
  "repair_print_layout",
  "repair_named_ranges",
  "repair_formula_errors",
  "repair_merged_cells",
  "detect_header_row",
  "normalize_headers",
  "trim_whitespace",
  "remove_duplicates",
  "parse_dates",
  "parse_numbers",
  "standardize_currency",
  "fill_missing_values",
  "split_column",
  "merge_columns",
  "detect_outliers",
  "fuzzy_match",
  "clear_range",
  "clear_values",
  "clear_values_raw",
  "clear_formats",
  "copy_range",
  "move_range",
  "insert_rows",
  "delete_rows",
  "insert_columns",
  "delete_columns",
  "merge_range",
  "unmerge_range",
  "append_table_rows",
  "update_table_rows",
  "create_table",
  "resize_table",
  "reorder_table_columns",
  "clear_table_data",
  "clear_table_filters",
  "sort_table",
  "filter_range",
  "set_table_total_row",
  "set_table_style",
  "copy_table_structure",
  "validate_table_against_template",
  "create_sheet",
  "copy_sheet",
  "rename_sheet",
  "delete_sheet",
  "hide_sheet",
  "unhide_sheet",
  "protect_sheet",
  "unprotect_sheet",
  "clear_sheet",
  "set_sheet_tab_color",
  "autofit",
  "autofit_rows",
  "copy_template_sheet",
  "detect_templates",
  "register_template",
  "unregister_template",
  "read_template",
  "list_templates",
  "infer_template_regions",
  "clear_template_data_regions",
  "fill_template_regions",
  "validate_sheet_against_template",
  "repair_sheet_from_template",
  "create_snapshot",
  "create_backup",
  "list_snapshots",
  "read_snapshot",
  "compare_snapshots",
  "refresh_snapshot",
  "invalidate_snapshot",
  "delete_snapshot",
  "list_backups",
  "read_backup",
  "verify_backup",
  "create_file_backup",
  "restore_file_backup",
  "prune_backups",
  "pin_backup",
  "unpin_backup",
  "delete_backup",
  "validate_compact",
  "validate_workbook",
  "validate_sheet",
  "validate_template_consistency",
  "validate_formulas",
  "validate_styles",
  "validate_tables",
  "validate_filters",
  "validate_print_layout",
  "validate_no_broken_references",
  "validate_no_formula_errors",
  "validate_no_unintended_changes",
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
