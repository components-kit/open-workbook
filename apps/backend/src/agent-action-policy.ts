import type { ExcelOperation } from "@components-kit/open-workbook-protocol";
import type { PendingAgentAction } from "./agent-operation-store.js";

export type AgentOperationRisk =
  | "read_only"
  | "safe_format"
  | "small_value_write"
  | "table_append"
  | "formula_write"
  | "broad_range_write"
  | "structure_change"
  | "destructive";

export type AgentCacheRisk = "none" | "low" | "medium" | "high";
export type AgentSafetyRisk = "low" | "medium" | "high";
export type AgentCacheAction = "none" | "update_cache" | "partial_invalidate" | "rebuild_context";

export interface AgentUpdateRisk {
  cacheRisk: AgentCacheRisk;
  safetyRisk: AgentSafetyRisk;
  cacheAction: AgentCacheAction;
  reason: string;
  invalidatedFacets: string[];
  preservedFacets: string[];
  requiresRefreshBeforeNextMutation: boolean;
}

export interface AgentActionDefinition {
  kind: string;
  risk: AgentOperationRisk;
  previewRequired: boolean;
  confirmationRequired: boolean;
}

export const AGENT_ACTION_REGISTRY: AgentActionDefinition[] = [
  { kind: "range.write_values", risk: "small_value_write", previewRequired: true, confirmationRequired: true },
  { kind: "range.write_values_many", risk: "broad_range_write", previewRequired: true, confirmationRequired: true },
  { kind: "range.write_formulas", risk: "formula_write", previewRequired: true, confirmationRequired: true },
  { kind: "range.write_number_formats", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "range.write_number_formats_many", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "range.write_styles", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "range.write_styles_many", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "range.write_data_validation", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "range.write_conditional_formatting", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "range.clear_style_dimensions", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "range.clear_style_dimensions_many", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "range.clear", risk: "destructive", previewRequired: true, confirmationRequired: true },
  { kind: "range.clear_many", risk: "destructive", previewRequired: true, confirmationRequired: true },
  { kind: "range.clear_values", risk: "destructive", previewRequired: true, confirmationRequired: true },
  { kind: "range.clear_formats", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "range.clear_formats_many", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "range.clear_values_keep_format", risk: "destructive", previewRequired: true, confirmationRequired: true },
  { kind: "range.insert_rows", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "range.delete_rows", risk: "destructive", previewRequired: true, confirmationRequired: true },
  { kind: "range.insert_columns", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "range.delete_columns", risk: "destructive", previewRequired: true, confirmationRequired: true },
  { kind: "range.merge", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "range.unmerge", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "range.autofit_columns", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "range.autofit_rows", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "range.autofit_many", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "range.copy", risk: "broad_range_write", previewRequired: true, confirmationRequired: true },
  { kind: "range.move", risk: "destructive", previewRequired: true, confirmationRequired: true },
  { kind: "range.reorder_columns", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "range.apply_autofilter", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "range.clear_autofilter", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "range.hide_columns", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "range.unhide_columns", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "sheet.copy", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "sheet.copy_clean_data_regions", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "sheet.create", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "sheet.rename", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "sheet.delete", risk: "destructive", previewRequired: true, confirmationRequired: true },
  { kind: "sheet.hide", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "sheet.unhide", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "sheet.protect", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "sheet.unprotect", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "sheet.clear", risk: "destructive", previewRequired: true, confirmationRequired: true },
  { kind: "sheet.set_tab_color", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "sheet.freeze_panes", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "workbook.calculate", risk: "read_only", previewRequired: true, confirmationRequired: true },
  { kind: "workbook.save", risk: "destructive", previewRequired: true, confirmationRequired: true },
  { kind: "workbook.snapshot", risk: "read_only", previewRequired: true, confirmationRequired: true },
  { kind: "workbook.create_backup", risk: "read_only", previewRequired: true, confirmationRequired: true },
  { kind: "snapshot.refresh", risk: "read_only", previewRequired: true, confirmationRequired: true },
  { kind: "snapshot.invalidate", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "snapshot.delete", risk: "destructive", previewRequired: true, confirmationRequired: true },
  { kind: "backup.create_file", risk: "read_only", previewRequired: true, confirmationRequired: true },
  { kind: "backup.restore_file", risk: "destructive", previewRequired: true, confirmationRequired: true },
  { kind: "backup.prune", risk: "destructive", previewRequired: true, confirmationRequired: true },
  { kind: "backup.pin", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "backup.unpin", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "backup.delete", risk: "destructive", previewRequired: true, confirmationRequired: true },
  { kind: "workbook.restore_backup", risk: "destructive", previewRequired: true, confirmationRequired: true },
  { kind: "workbook.import_local_config", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "workbook.embed_local_config", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "workbook.import_embedded_local_config", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "workbook.close", risk: "destructive", previewRequired: true, confirmationRequired: true },
  { kind: "formula.copy_patterns", risk: "formula_write", previewRequired: true, confirmationRequired: true },
  { kind: "formula.fill_pattern", risk: "formula_write", previewRequired: true, confirmationRequired: true },
  { kind: "formula.repair_patterns", risk: "formula_write", previewRequired: true, confirmationRequired: true },
  { kind: "formula.convert_to_values", risk: "destructive", previewRequired: true, confirmationRequired: true },
  { kind: "names.create", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "names.update", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "names.delete", risk: "destructive", previewRequired: true, confirmationRequired: true },
  { kind: "region.register", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "region.clear_values", risk: "destructive", previewRequired: true, confirmationRequired: true },
  { kind: "region.write_values", risk: "broad_range_write", previewRequired: true, confirmationRequired: true },
  { kind: "region.fill", risk: "broad_range_write", previewRequired: true, confirmationRequired: true },
  { kind: "table.append_rows", risk: "table_append", previewRequired: true, confirmationRequired: true },
  { kind: "table.update_rows", risk: "broad_range_write", previewRequired: true, confirmationRequired: true },
  { kind: "table.create", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "table.resize", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "table.reorder_columns", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "table.clear_data_keep_formulas", risk: "destructive", previewRequired: true, confirmationRequired: true },
  { kind: "table.clear_filters", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "table.apply_filters", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "table.sort", risk: "broad_range_write", previewRequired: true, confirmationRequired: true },
  { kind: "table.apply_view", risk: "broad_range_write", previewRequired: true, confirmationRequired: true },
  { kind: "table.set_total_row", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "table.set_style", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "table.copy_structure", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "template.register", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "template.unregister", risk: "destructive", previewRequired: true, confirmationRequired: true },
  { kind: "template.repair_sheet", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "style.copy_dimensions", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "style.copy_dimensions_many", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "workflow.replace_styled_table", risk: "destructive", previewRequired: true, confirmationRequired: true },
  { kind: "style.repair_consistency", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "visual_readability.apply", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "clean.transform", risk: "broad_range_write", previewRequired: true, confirmationRequired: true },
  { kind: "clean.transform_many", risk: "broad_range_write", previewRequired: true, confirmationRequired: true }
];

const RISK_RANK: Record<AgentOperationRisk, number> = {
  read_only: 0,
  safe_format: 1,
  small_value_write: 2,
  table_append: 3,
  formula_write: 4,
  broad_range_write: 5,
  structure_change: 6,
  destructive: 7
};

export function classifyAgentActionRisk(action: PendingAgentAction): AgentOperationRisk {
  const risks = action.kind === "batch"
    ? action.operations.map(riskForOperation)
    : action.kind === "workflow.replace_styled_table"
      ? [...action.operations.map(riskForOperation), ...action.styleCopies.map(() => riskForOperationKind("style.copy_dimensions"))]
      : action.kind === "style.copy_dimensions_many"
        ? action.requests.map(() => riskForOperationKind("style.copy_dimensions"))
    : [riskForOperationKind(action.kind)];
  return highestRisk(risks);
}

export function assessAgentUpdateRisk(action: PendingAgentAction): AgentUpdateRisk {
  const operationRisk = classifyAgentActionRisk(action);
  switch (operationRisk) {
    case "read_only":
      return updateRisk("none", "low", "none", "Read-only operation does not change workbook context.", [], ALL_FACETS, false);
    case "safe_format":
      return updateRisk("low", "low", "update_cache", "Formatting/filter-style operation preserves values and schema.", ["formats"], ["schema", "headers", "fieldContext", "validation", "values"], false);
    case "small_value_write":
      return updateRisk("low", "low", "update_cache", "Small known value write can update cached values and mark dependent results stale.", ["values", "aggregates", "formulaResults"], ["schema", "headers", "fieldContext", "validation"], false);
    case "table_append":
      return updateRisk("medium", "medium", "partial_invalidate", "Append row changes table dimensions and aggregate values.", ["tableDimensions", "values", "aggregates", "rowPositions"], ["schema", "headers", "fieldContext", "validation"], false);
    case "formula_write":
      return updateRisk("medium", "medium", "partial_invalidate", "Formula writes can change formula regions and formula results.", ["formulas", "formulaResults", "aggregates", "values"], ["headers", "fieldContext", "validation"], false);
    case "broad_range_write":
      return updateRisk("low", "medium", "update_cache", "Known range values can update cached values while dependent aggregate/formula facets become stale.", ["values", "aggregates", "formulaResults"], ["schema", "headers", "fieldContext", "validation"], false);
    case "structure_change":
      return updateRisk("high", "medium", "rebuild_context", "Structure changes can invalidate schema, headers, field mapping, and regions.", ["schema", "headers", "tableDimensions", "regions", "fieldContext", "validation", "rowPositions"], ["metadata"], true);
    case "destructive":
      return updateRisk("high", "high", "rebuild_context", "Destructive operation can remove data or structure and requires context rebuild.", ["metadata", "schema", "headers", "tableDimensions", "regions", "fieldContext", "validation", "values", "aggregates", "rowPositions"], [], true);
  }
}

function riskForOperation(operation: ExcelOperation): AgentOperationRisk {
  if (operation.kind === "range.write_values_many") {
    const cellCount = operation.entries.reduce((total, entry) => total + matrixCellCount(entry.values), 0);
    return cellCount <= 4 ? "small_value_write" : "broad_range_write";
  }
  return riskForOperationKind(operation.kind);
}

export function riskForOperationKind(
  kind: ExcelOperation["kind"]
    | "table.append_rows"
    | "table.update_rows"
    | "table.create"
    | "table.resize"
    | "table.reorder_columns"
    | "table.clear_data_keep_formulas"
    | "table.clear_filters"
    | "table.apply_filters"
    | "table.sort"
    | "table.apply_view"
    | "table.set_total_row"
    | "table.set_style"
    | "table.copy_structure"
    | "template.register"
    | "template.unregister"
    | "template.repair_sheet"
    | "style.copy_dimensions"
    | "style.copy_dimensions_many"
    | "workflow.replace_styled_table"
    | "style.repair_consistency"
    | "visual_readability.apply"
    | "clean.transform"
    | "clean.transform_many"
    | "workbook.snapshot"
    | "workbook.create_backup"
    | "snapshot.refresh"
    | "snapshot.invalidate"
    | "snapshot.delete"
    | "backup.create_file"
    | "backup.restore_file"
    | "backup.prune"
    | "backup.pin"
    | "backup.unpin"
    | "backup.delete"
    | "workbook.restore_backup"
    | "workbook.import_local_config"
    | "workbook.embed_local_config"
    | "workbook.import_embedded_local_config"
    | "workbook.close"
    | "formula.copy_patterns"
    | "formula.fill_pattern"
    | "formula.repair_patterns"
    | "formula.convert_to_values"
    | "names.create"
    | "names.update"
    | "names.delete"
    | "region.register"
    | "region.clear_values"
    | "region.write_values"
    | "region.fill"
): AgentOperationRisk {
  const registered = AGENT_ACTION_REGISTRY.find((action) => action.kind === kind);
  if (registered) {
    return registered.risk;
  }
  if (kind.startsWith("range.read")) return "read_only";
  if (kind.includes("style") || kind.includes("format") || kind.includes("autofit") || kind.includes("filter")) return "safe_format";
  if (kind.includes("formula")) return "formula_write";
  if (kind.startsWith("sheet.") || kind.includes("insert") || kind.includes("delete") || kind.includes("merge")) return "structure_change";
  if (kind.includes("clear") || kind.includes("save") || kind.includes("restore")) return "destructive";
  return "broad_range_write";
}

function highestRisk(risks: AgentOperationRisk[]): AgentOperationRisk {
  return risks.reduce<AgentOperationRisk>((highest, risk) => RISK_RANK[risk] > RISK_RANK[highest] ? risk : highest, "read_only");
}

const ALL_FACETS = ["metadata", "schema", "headers", "tableDimensions", "regions", "fieldContext", "validation", "formats", "formulas", "formulaResults", "filters", "values", "aggregates", "rowPositions", "selection", "names"];

function updateRisk(
  cacheRisk: AgentCacheRisk,
  safetyRisk: AgentSafetyRisk,
  cacheAction: AgentCacheAction,
  reason: string,
  invalidatedFacets: string[],
  preservedFacets: string[],
  requiresRefreshBeforeNextMutation: boolean
): AgentUpdateRisk {
  return {
    cacheRisk,
    safetyRisk,
    cacheAction,
    reason,
    invalidatedFacets,
    preservedFacets,
    requiresRefreshBeforeNextMutation
  };
}

function matrixCellCount(matrix: unknown[][]): number {
  return matrix.reduce((total, row) => total + row.length, 0);
}
