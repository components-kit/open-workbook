import type { WorkbookId } from "./ids.js";

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
  | "find_target"
  | "write_values"
  | "write_formulas"
  | "format_range"
  | "clear_values"
  | "append_table_rows"
  | "sort_table"
  | "filter_range"
  | "autofit"
  | "copy_template_sheet"
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
    autoApplyBlockedReason?: string;
    targetFingerprintStatus?: "matched" | "changed" | "not_applicable";
    intentSource?: "caller_structured" | "deterministic_fallback" | "mixed";
    intentAction?: AgentIntentAction;
    intentAccepted?: boolean;
    intentRejectedReason?: string;
  };
}
