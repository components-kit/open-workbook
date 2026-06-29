import type { AgentId, TransactionId, WorkbookId } from "./ids.js";

export type WorkbookContextId = string & { readonly __brand: "WorkbookContextId" };
export type AgentOperationId = string & { readonly __brand: "AgentOperationId" };

export const AGENT_RUN_MODES = [
  "auto",
  "status",
  "prepare",
  "find",
  "answer",
  "preview_update",
  "apply_update",
  "operation_status",
  "cancel_operation",
  "validate",
  "rollback"
] as const;

export type AgentRunMode = typeof AGENT_RUN_MODES[number];

export const AGENT_RUN_STATUSES = [
  "SUCCESS",
  "IN_PROGRESS",
  "PREVIEW_READY",
  "NEEDS_INPUT",
  "AMBIGUOUS_TARGET",
  "NEEDS_WORKFLOW_REDIRECT",
  "NOT_FOUND",
  "STALE_CONTEXT",
  "VALIDATION_FAILED",
  "CONFLICT",
  "ERROR"
] as const;

export type AgentRunStatus = typeof AGENT_RUN_STATUSES[number];

export type AgentNextAction =
  | "answer_now"
  | "call_apply_update"
  | "ask_user"
  | "call_preview_update"
  | "call_with_target"
  | "fetch_resource"
  | "retry_after_refresh"
  | "manual_review";

export type AgentTaskOutcome =
  | "final_answer"
  | "preview_ready"
  | "apply_complete"
  | "needs_user_input"
  | "cannot_complete";

export interface AgentRequiredFollowup {
  mode?: AgentRunMode;
  nextAction?: AgentNextAction;
  operationId?: AgentOperationId | string;
  confirmationToken?: string;
  instruction: string;
}

export const AGENT_DETAIL_LEVELS = [
  "workbook_summary",
  "semantic_index",
  "sheet_summary",
  "style_overview",
  "workbook_design_overview",
  "table_sample",
  "full_table"
] as const;

export type AgentDetailLevel = typeof AGENT_DETAIL_LEVELS[number];

export interface AgentRunTarget {
  workbookId?: WorkbookId | string;
  workbookName?: string;
  candidateId?: string;
  sheetName?: string;
  tableName?: string;
  range?: string;
  address?: string;
  row?: number;
  column?: string;
  entity?: string;
}

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
  "get_permissions",
  "set_permissions",
  "allow_destructive_actions",
  "close_workbook",
  "prepare_session",
  "create_formula_sheet",
  "create_template_report",
  "create_pivot_chart_summary",
  "preview_risky_edit",
  "inspect_analyze",
  "rollback_validate",
  "read_formulas",
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
  "find_similar_rows",
  "analyze_reference_sheet",
  "find_style_references",
  "style_overview",
  "workbook_design_overview",
  "grouped_header",
  "improve_visual_readability",
  "transform_values",
  "derive_values",
  "settle_reconciliation",
  "transform_sheets",
  "write_values",
  "write_formulas",
  "write_number_formats",
  "format_range",
  "write_data_validation",
  "write_conditional_formatting",
  "clear_style_dimensions",
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
  "replace_range_with_styled_table",
  "read_style_summary",
  "format_diagnostics",
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
  "reorder_range_columns",
  "insert_rows",
  "delete_rows",
  "insert_columns",
  "delete_columns",
  "hide_columns",
  "unhide_columns",
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
  "apply_table_view",
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
  "freeze_panes",
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

export const AGENT_CONTEXT_STRATEGIES = ["auto", "overview", "focused", "analysis", "audit", "mutation"] as const;
export type AgentContextStrategy = typeof AGENT_CONTEXT_STRATEGIES[number];

export const AGENT_CONTEXT_SCOPES = ["active_selection", "active_region", "active_sheet", "target", "workbook"] as const;
export type AgentContextScope = typeof AGENT_CONTEXT_SCOPES[number];
export const AGENT_CONTEXT_LEVELS = [0, 1, 2, 3, 4, 5] as const;
export type AgentContextLevel = typeof AGENT_CONTEXT_LEVELS[number];

export const AGENT_CONTEXT_FACETS = [
  "metadata",
  "schema",
  "regions",
  "tables",
  "values",
  "field_context",
  "formulas",
  "formats",
  "validation",
  "filters",
  "names",
  "comments",
  "merged_cells",
  "conditional_formatting",
  "charts",
  "pivots",
  "hidden"
] as const;
export type AgentContextFacet = typeof AGENT_CONTEXT_FACETS[number];

export interface AgentContextPolicy {
  level?: AgentContextLevel;
  strategy?: AgentContextStrategy;
  scope?: AgentContextScope;
  include?: AgentContextFacet[];
}

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
  continuation?: AgentContinuation;
  intent?: AgentRunIntent;
  target?: AgentRunTarget;
  values?: Record<string, unknown> & {
    patches?: Array<{
      target: AgentRunTarget;
      values?: unknown[][];
      rows?: unknown[][];
      formulas?: Array<Array<string | null>>;
      style?: Record<string, unknown>;
      numberFormat?: string | string[][];
      numberFormats?: string | string[][];
      validation?: Record<string, unknown>;
      conditionalFormatting?: Record<string, unknown>;
      note?: string;
      comment?: string;
      options?: string[] | Record<string, unknown>;
      allowedValues?: string[];
      reason?: string;
    }>;
  };
  autoApply?: boolean;
  context?: AgentContextPolicy;
  detailLevel?: AgentDetailLevel;
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
  semanticRole?: AgentSemanticRole;
  aliases?: string[];
  confidence: number;
  reason?: string;
  nextRequestHint?: string;
}

export type AgentSemanticRole =
  | "workbook"
  | "data_table"
  | "transaction_sheet"
  | "summary_sheet"
  | "template_sheet"
  | "lookup_sheet"
  | "form_region"
  | "formula_region"
  | "style_reference"
  | "named_region"
  | "selection"
  | "unknown";

export interface AgentSemanticIndexEntry {
  id: string;
  label: string;
  role: AgentSemanticRole;
  sourceKind: AgentCandidate["kind"] | "selection";
  sheetName?: string;
  tableName?: string;
  range?: string;
  aliases: string[];
  confidence: number;
  evidence: string[];
  supportedActions: AgentIntentAction[];
  nextRequestHints?: string[];
}

export interface AgentSemanticWorkbookIndex {
  kind: "semantic_workbook_index";
  source: "cached_metadata";
  workbook: {
    workbookId?: WorkbookId | string;
    name: string;
    activeSheet?: string;
    sheetCount: number;
  };
  detailLevel: "structure" | "sampled";
  entryCount: number;
  entries: AgentSemanticIndexEntry[];
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

export interface AgentContinuation {
  workbookContextId?: WorkbookContextId | string;
  operationId?: AgentOperationId | string;
  transactionId?: TransactionId | string;
  resultUri?: string;
  fullResultUri?: string;
  freshness?: {
    workbookId?: WorkbookId | string;
    workbookContentVersion?: number;
    workbookStructureHash?: string;
    contextUpdatedAt?: number;
  };
  nextRequest?: string;
  responseMode?: "brief" | "standard" | "verbose";
}

export interface AgentChangePreview {
  sheetName: string;
  range?: string;
  cell?: string;
  columnName?: string;
  before?: unknown;
  after?: unknown;
}

export interface AgentContextUsed {
  strategy: AgentContextStrategy;
  scope: AgentContextScope;
  levelRequested?: AgentContextLevel;
  levelUsed: AgentContextLevel;
  levelReason: string;
  stagesUsed: string[];
  included: string[];
  rangesRead?: string[];
  rowsRead?: number;
  estimatedTokens?: number;
  truncated?: boolean;
  confidence?: number;
  source?: "cache" | "live" | "mixed" | "none";
  continuation?: {
    available: boolean;
    suggestedNext?: string[];
  };
}

export interface AgentRunOutput {
  status: AgentRunStatus;
  mode: AgentRunMode;
  workbookContextId?: WorkbookContextId | string;
  operationId?: AgentOperationId | string;
  transactionId?: TransactionId | string;
  confirmationToken?: string;
  summary: string;
  answer?: unknown;
  metrics?: Record<string, unknown>;
  changes?: AgentChangePreview[];
  candidates?: AgentCandidate[];
  proof: AgentProofReference[];
  resourceLinks: AgentResourceLink[];
  invalidatedContextIds?: Array<WorkbookContextId | string>;
  invalidatedResourceUris?: string[];
  continuation?: AgentContinuation;
  contextUsed?: AgentContextUsed;
  nextAction: AgentNextAction;
  taskOutcome?: AgentTaskOutcome;
  finalAnswer?: string;
  agentInstruction?: string;
  maxRecommendedFollowupCalls?: number;
  requiredFollowup?: AgentRequiredFollowup;
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
    metadataFreshnessReason?: string;
    metadataDetailLevel?: "structure" | "sampled";
    internalReadCount?: number;
    fullReadCellCount?: number;
    fullReadUsed?: boolean;
    safetyFingerprintOnly?: boolean;
    workflowRoute?: string;
    workflowConfidence?: number;
    workflowReasons?: string[];
    semanticIndexStatus?: "built" | "not_applicable";
    semanticEntryCount?: number;
    semanticCandidateUsed?: boolean;
    metadataPolicy?: "structure_only" | "sampled_allowed" | "sampled_required";
    readPolicy?: "metadata_only" | "targeted_read" | "preview_only" | "apply_only" | "not_applicable";
    contextDecision?: {
      strategy: AgentContextStrategy;
      scope: AgentContextScope;
      include: AgentContextFacet[];
      level?: AgentContextLevel;
      source: "caller" | "inferred";
      reason: string;
    };
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
    workflowKind?: string;
    groupedOperationCount?: number;
    styleCopyCount?: number;
    clearFormatCount?: number;
    fragmentationRedirectCount?: number;
    detectedFamily?: string;
    suggestedWorkflowKind?: string;
    targetFingerprintStatus?: "matched" | "changed" | "not_applicable";
    targetHintCount?: number;
    targetHintUsed?: boolean;
    intentSource?: "caller_structured" | "deterministic_fallback" | "mixed";
    intentAction?: AgentIntentAction;
    intentAccepted?: boolean;
    intentRejectedReason?: string;
  };
}
