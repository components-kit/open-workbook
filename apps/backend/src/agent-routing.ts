import type { AgentRunMode } from "@components-kit/open-workbook-protocol";
import { modeForIntentAction, type NormalizedAgentIntent } from "./agent-intent.js";

export interface IntentRoute {
  mode: AgentRunMode;
  matchedRule: string;
  confidence: number;
  reasons: string[];
  workflowRoute: AgentWorkflowRoute;
  workflowConfidence: number;
  workflowReasons: string[];
  metadataPolicy: "structure_only" | "sampled_allowed" | "sampled_required";
  readPolicy: "metadata_only" | "targeted_read" | "preview_only" | "apply_only" | "not_applicable";
}

export type AgentWorkflowRoute =
  | "workbook.summary"
  | "semantic_index.find"
  | "sheet.summary"
  | "schema.inspect"
  | "style.inspect"
  | "format.diagnostics"
  | "range.read"
  | "table.sample"
  | "mutation.preview"
  | "mutation.apply"
  | "validation.run"
  | "safety.rollback"
  | "status"
  | "operation.status";

interface IntentRouteRule {
  mode: AgentRunMode;
  matchedRule: string;
  confidence: number;
  pattern: RegExp;
  reason: string;
}

const ROUTE_RULES: IntentRouteRule[] = [
  {
    mode: "preview_update",
    matchedRule: "mutation.keyword",
    confidence: 0.82,
    pattern: /\b(add|update|set|change|replace|write|fill|append|edit|fix|repair|create|insert|delete|remove|clear|rename|format|style|duplicate|copy|template)\b/i,
    reason: "Request contains workbook mutation keywords."
  }
];

export function routeAgentRequest(request: string, requestedMode: AgentRunMode = "auto", intent?: NormalizedAgentIntent): IntentRoute {
  const workflow = routeAgentWorkflow(request, requestedMode, intent);
  if (requestedMode !== "auto") {
    return {
      mode: requestedMode,
      matchedRule: "mode.explicit",
      confidence: 1,
      reasons: [`Caller explicitly requested ${requestedMode}.`],
      ...workflow
    };
  }

  if (intent?.accepted && intent.action) {
    return {
      mode: modeForIntentAction(intent.action),
      matchedRule: "caller_intent.action",
      confidence: intent.confidence ?? 0.9,
      reasons: [intent.reason ?? `Caller provided structured intent action ${intent.action}.`],
      ...workflow
    };
  }

  const matched = ROUTE_RULES.find((rule) => rule.pattern.test(request));
  if (matched) {
    return {
      mode: matched.mode,
      matchedRule: matched.matchedRule,
      confidence: matched.confidence,
      reasons: [matched.reason],
      ...workflow
    };
  }

  return {
    mode: "answer",
    matchedRule: "default.answer",
    confidence: 0.65,
    reasons: ["No mutation route matched; defaulting to answer mode."],
    ...workflow
  };
}

function routeAgentWorkflow(
  request: string,
  requestedMode: AgentRunMode,
  intent?: NormalizedAgentIntent
): Pick<IntentRoute, "workflowRoute" | "workflowConfidence" | "workflowReasons" | "metadataPolicy" | "readPolicy"> {
  if (requestedMode === "status") return workflow("status", 1, "Caller requested runtime status.", "structure_only", "not_applicable");
  if (requestedMode === "apply_update") return workflow("mutation.apply", 1, "Caller requested applying a previewed operation.", "structure_only", "apply_only");
  if (requestedMode === "validate") return workflow("validation.run", 1, "Caller requested validation.", "sampled_allowed", "targeted_read");
  if (requestedMode === "rollback") return workflow("safety.rollback", 1, "Caller requested rollback.", "structure_only", "not_applicable");
  if (requestedMode === "operation_status" || requestedMode === "cancel_operation") return workflow("operation.status", 1, "Caller requested operation state.", "structure_only", "not_applicable");
  if (requestedMode === "prepare") return workflow("workbook.summary", 0.95, "Prepare builds reusable workbook structure context.", "structure_only", "metadata_only");
  if (requestedMode === "find") return workflow("semantic_index.find", 0.95, "Find searches cached workbook metadata and semantic index.", "sampled_allowed", "metadata_only");
  if (requestedMode === "preview_update") return workflow("mutation.preview", 1, "Caller requested mutation preview.", "sampled_allowed", "preview_only");

  const action = intent?.accepted ? intent.action : undefined;
  if (action === "read_style_summary" || action === "read_style_fingerprint" || action === "compare_style_fingerprint") {
    return workflow("style.inspect", intent?.confidence ?? 0.9, "Structured style inspection action.", "sampled_allowed", "targeted_read");
  }
  if (action === "format_diagnostics") {
    return workflow("format.diagnostics", intent?.confidence ?? 0.9, "Structured formatting diagnostics action.", "sampled_allowed", "targeted_read");
  }
  if (action === "read_schema") return workflow("schema.inspect", intent?.confidence ?? 0.9, "Structured schema inspection action.", "structure_only", "metadata_only");
  if (action === "find_target") return workflow("semantic_index.find", intent?.confidence ?? 0.9, "Structured target finding action.", "sampled_allowed", "metadata_only");
  if (action === "read_values") return workflow("range.read", intent?.confidence ?? 0.85, "Structured value read action.", "sampled_allowed", "targeted_read");
  if (action && modeForIntentAction(action) === "preview_update") return workflow("mutation.preview", intent?.confidence ?? 0.9, "Structured mutation action.", "sampled_allowed", "preview_only");

  const lower = request.toLowerCase();
  if (/\b(semantic index|workbook index|find|where is|locate|which sheet|which table)\b/.test(lower)) {
    return workflow("semantic_index.find", 0.84, "Request asks for workbook target discovery.", "sampled_allowed", "metadata_only");
  }
  if (/\b(what is this workbook|about this workbook|look at this workbook|look into|check this workbook|review this workbook|overview|summar(y|ize)|how many sheets|list sheets|list tables|named ranges?|\.xlsx)\b/.test(lower)) {
    return workflow("workbook.summary", 0.86, "Request asks for workbook overview.", "structure_only", "metadata_only");
  }
  if (/\b(this|current|active)?\s*sheet\b/.test(lower) && /\b(look|overview|summar(y|ize)|inspect|what)\b/.test(lower)) {
    return workflow("sheet.summary", 0.82, "Request asks for sheet overview.", "structure_only", "metadata_only");
  }
  if (/\b(style|styling|formatting|border|fill|font|alignment|number format)\b/.test(lower) && /\b(current|inspect|what|show|read|look)\b/.test(lower)) {
    return workflow("style.inspect", 0.82, "Request asks for current styling.", "sampled_allowed", "targeted_read");
  }
  if (/\b(formatting error|wrong format|date format|number format|diagnos(e|is)|why.*format)\b/.test(lower)) {
    return workflow("format.diagnostics", 0.84, "Request asks for formatting diagnostics.", "sampled_allowed", "targeted_read");
  }
  if (/\b(sample|first rows|table sample)\b/.test(lower)) {
    return workflow("table.sample", 0.8, "Request asks for bounded table sample.", "sampled_required", "targeted_read");
  }
  if (ROUTE_RULES.some((rule) => rule.pattern.test(request))) {
    return workflow("mutation.preview", 0.82, "Request contains workbook mutation language.", "sampled_allowed", "preview_only");
  }
  return workflow("range.read", 0.65, "Default answer route may use targeted reads after metadata.", "sampled_allowed", "targeted_read");
}

function workflow(
  workflowRoute: AgentWorkflowRoute,
  workflowConfidence: number,
  reason: string,
  metadataPolicy: IntentRoute["metadataPolicy"],
  readPolicy: IntentRoute["readPolicy"]
) {
  return {
    workflowRoute,
    workflowConfidence,
    workflowReasons: [reason],
    metadataPolicy,
    readPolicy
  };
}
