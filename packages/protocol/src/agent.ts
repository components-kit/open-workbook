import type { AgentId, WorkbookId } from "./ids.js";

export type WorkbookContextId = string & { readonly __brand: "WorkbookContextId" };
export type AgentOperationId = string & { readonly __brand: "AgentOperationId" };

export type AgentRunMode =
  | "auto"
  | "status"
  | "prepare"
  | "find"
  | "answer"
  | "preview_update"
  | "apply_update"
  | "validate"
  | "rollback";

export type AgentRunStatus =
  | "SUCCESS"
  | "PREVIEW_READY"
  | "NEEDS_INPUT"
  | "AMBIGUOUS_TARGET"
  | "NOT_FOUND"
  | "STALE_CONTEXT"
  | "VALIDATION_FAILED"
  | "CONFLICT"
  | "ERROR";

export type AgentNextAction =
  | "answer_now"
  | "call_apply_update"
  | "ask_user"
  | "call_with_target"
  | "fetch_resource"
  | "retry_after_refresh"
  | "manual_review";

export interface AgentRunTarget {
  workbookId?: WorkbookId | string;
  workbookName?: string;
  candidateId?: string;
  sheetName?: string;
  tableName?: string;
  range?: string;
  row?: number;
  column?: string;
  entity?: string;
}

export type AgentIntentAction =
  | "read_values"
  | "read_schema"
  | "list_open_workbooks"
  | "get_workbook_info"
  | "refresh_workbook_snapshot"
  | "get_workbook_snapshot"
  | "detect_external_changes"
  | "restore_workbook_backup"
  | "export_local_config"
  | "import_local_config"
  | "embed_local_config"
  | "read_embedded_local_config"
  | "import_embedded_local_config"
  | "close_workbook"
  | "prepare_session"
  | "create_formula_sheet"
  | "create_template_report"
  | "create_pivot_chart_summary"
  | "preview_risky_edit"
  | "inspect_analyze"
  | "rollback_validate"
  | "read_formula_patterns"
  | "get_formula_dependency_graph"
  | "trace_formula_precedents"
  | "trace_formula_dependents"
  | "validate_formula_range"
  | "validate_formula_against_template"
  | "find_formula_errors"
  | "explain_formula"
  | "copy_formula_patterns"
  | "fill_formula_down"
  | "fill_formula_right"
  | "repair_formula_patterns"
  | "convert_formulas_to_values"
  | "recalculate_formulas"
  | "read_named_item"
  | "create_name"
  | "update_name"
  | "delete_name"
  | "read_region"
  | "register_region"
  | "clear_region_values"
  | "write_region_values"
  | "fill_region"
  | "find_target"
  | "write_values"
  | "write_formulas"
  | "write_number_formats"
  | "format_range"
  | "read_range_compact"
  | "get_range_summary"
  | "read_hyperlinks"
  | "read_comments"
  | "read_notes"
  | "read_merged_cells"
  | "read_data_validation"
  | "read_conditional_formatting"
  | "search_range"
  | "find_blank_cells"
  | "find_range_errors"
  | "write_styles_many"
  | "read_style_fingerprint"
  | "compare_style_fingerprint"
  | "get_theme"
  | "apply_theme"
  | "copy_style_from_template"
  | "repair_style_consistency"
  | "repair_style_from_template"
  | "repair_formulas_from_template"
  | "repair_filters_from_template"
  | "repair_table_structure"
  | "repair_print_layout"
  | "repair_named_ranges"
  | "repair_formula_errors"
  | "repair_merged_cells"
  | "detect_header_row"
  | "normalize_headers"
  | "trim_whitespace"
  | "remove_duplicates"
  | "parse_dates"
  | "parse_numbers"
  | "standardize_currency"
  | "fill_missing_values"
  | "split_column"
  | "merge_columns"
  | "detect_outliers"
  | "fuzzy_match"
  | "clear_range"
  | "clear_values"
  | "clear_values_raw"
  | "clear_formats"
  | "copy_range"
  | "move_range"
  | "insert_rows"
  | "delete_rows"
  | "insert_columns"
  | "delete_columns"
  | "merge_range"
  | "unmerge_range"
  | "append_table_rows"
  | "update_table_rows"
  | "create_table"
  | "resize_table"
  | "reorder_table_columns"
  | "clear_table_data"
  | "clear_table_filters"
  | "sort_table"
  | "filter_range"
  | "set_table_total_row"
  | "set_table_style"
  | "copy_table_structure"
  | "validate_table_against_template"
  | "create_sheet"
  | "copy_sheet"
  | "rename_sheet"
  | "delete_sheet"
  | "hide_sheet"
  | "unhide_sheet"
  | "protect_sheet"
  | "unprotect_sheet"
  | "clear_sheet"
  | "set_sheet_tab_color"
  | "autofit"
  | "autofit_rows"
  | "copy_template_sheet"
  | "detect_templates"
  | "register_template"
  | "unregister_template"
  | "read_template"
  | "list_templates"
  | "infer_template_regions"
  | "clear_template_data_regions"
  | "fill_template_regions"
  | "validate_sheet_against_template"
  | "repair_sheet_from_template"
  | "create_snapshot"
  | "create_backup"
  | "list_snapshots"
  | "read_snapshot"
  | "compare_snapshots"
  | "refresh_snapshot"
  | "invalidate_snapshot"
  | "delete_snapshot"
  | "list_backups"
  | "read_backup"
  | "verify_backup"
  | "create_file_backup"
  | "restore_file_backup"
  | "prune_backups"
  | "pin_backup"
  | "unpin_backup"
  | "delete_backup"
  | "validate_compact"
  | "validate_workbook"
  | "validate_sheet"
  | "validate_template_consistency"
  | "validate_formulas"
  | "validate_styles"
  | "validate_tables"
  | "validate_filters"
  | "validate_print_layout"
  | "validate_no_broken_references"
  | "validate_no_formula_errors"
  | "validate_no_unintended_changes"
  | "calculate"
  | "save";

export interface AgentRunIntent {
  action: AgentIntentAction;
  confidence?: number;
  reason?: string;
  targetHints?: string[];
}

export interface AgentRunInput {
  request: string;
  mode?: AgentRunMode;
  workbookContextId?: WorkbookContextId | string;
  operationId?: AgentOperationId | string;
  transactionId?: string;
  confirmationToken?: string;
  intent?: AgentRunIntent;
  target?: AgentRunTarget;
  values?: Record<string, unknown> & {
    patches?: Array<{
      target: AgentRunTarget;
      values?: unknown[][];
      rows?: unknown[][];
      reason?: string;
    }>;
  };
  responseMode?: "brief" | "standard" | "verbose";
  budget?: {
    maxPayloadBytes?: number;
    maxEstimatedTokens?: number;
    maxExamples?: number;
  };
}

export interface AgentRunExecutionContext {
  agentId?: AgentId | string;
  agentName?: string;
  clientType?: "mcp" | "daemon" | "cli" | "unknown";
}

export interface AgentCandidate {
  id: string;
  kind: "workbook" | "sheet" | "table" | "column" | "row" | "range" | "region" | "pivot" | "chart";
  label: string;
  sheetName?: string;
  tableName?: string;
  range?: string;
  confidence: number;
  reason?: string;
  nextRequestHint?: string;
}

export interface AgentProofReference {
  sheetName: string;
  range: string;
  label?: string;
  resourceUri?: string;
}

export interface AgentResourceLink {
  uri: string;
  name: string;
  description: string;
  mimeType: "application/json";
}

export interface AgentChangePreview {
  sheetName: string;
  range?: string;
  cell?: string;
  columnName?: string;
  before?: unknown;
  after?: unknown;
}

export interface AgentRunOutput {
  status: AgentRunStatus;
  mode: AgentRunMode;
  workbookContextId?: WorkbookContextId | string;
  operationId?: AgentOperationId | string;
  confirmationToken?: string;
  summary: string;
  answer?: unknown;
  metrics?: Record<string, unknown>;
  changes?: AgentChangePreview[];
  candidates?: AgentCandidate[];
  proof: AgentProofReference[];
  resourceLinks: AgentResourceLink[];
  nextAction: AgentNextAction;
  warnings: string[];
  telemetry: {
    internalCallCount: number;
    payloadBytes: number;
    estimatedTokens: number;
    elapsedMs: number;
    cacheHit: boolean;
    autoApplied?: boolean;
    safetyDecision?: string;
    previewOperationId?: AgentOperationId | string;
    validationStatus?: "passed" | "failed" | "not_run";
    metadataCacheStatus?: "hit" | "miss" | "not_applicable";
    internalReadCount?: number;
    fullReadCellCount?: number;
    candidateCount?: number;
    resourceLinkCount?: number;
    estimatedTokensSaved?: number;
    routeMode?: AgentRunMode;
    routeMatchedRule?: string;
    routeConfidence?: number;
    routeReasons?: string[];
    operationRisk?: string;
    actionHandlerId?: string;
    autoApplyBlockedReason?: string;
    targetFingerprintStatus?: "matched" | "changed" | "not_applicable";
    targetHintCount?: number;
    targetHintUsed?: boolean;
    intentSource?: "caller_structured" | "deterministic_fallback" | "mixed";
    intentAction?: AgentIntentAction;
    intentAccepted?: boolean;
    intentRejectedReason?: string;
  };
}
