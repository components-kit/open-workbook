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

export interface AgentActionDefinition {
  kind: string;
  risk: AgentOperationRisk;
  previewRequired: boolean;
  confirmationRequired: boolean;
}

export const AGENT_ACTION_REGISTRY: AgentActionDefinition[] = [
  { kind: "range.write_values", risk: "small_value_write", previewRequired: true, confirmationRequired: true },
  { kind: "range.write_formulas", risk: "formula_write", previewRequired: true, confirmationRequired: true },
  { kind: "range.write_styles", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "range.clear_values_keep_format", risk: "destructive", previewRequired: true, confirmationRequired: true },
  { kind: "range.autofit_columns", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "range.apply_autofilter", risk: "safe_format", previewRequired: true, confirmationRequired: true },
  { kind: "sheet.copy", risk: "structure_change", previewRequired: true, confirmationRequired: true },
  { kind: "workbook.calculate", risk: "read_only", previewRequired: true, confirmationRequired: true },
  { kind: "workbook.save", risk: "destructive", previewRequired: true, confirmationRequired: true },
  { kind: "table.append_rows", risk: "table_append", previewRequired: true, confirmationRequired: true },
  { kind: "table.sort", risk: "broad_range_write", previewRequired: true, confirmationRequired: true }
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
  const kinds = action.kind === "batch"
    ? action.operations.map((operation) => operation.kind)
    : [action.kind];
  return highestRisk(kinds.map(riskForOperationKind));
}

export function riskForOperationKind(kind: ExcelOperation["kind"] | "table.append_rows" | "table.sort"): AgentOperationRisk {
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
