import type { AgentRunInput } from "@components-kit/open-workbook-protocol";
import type { AgentIntentAction } from "./agent-intent.js";

export type AgentActionHandlerId =
  | "save_workbook"
  | "calculate_workbook"
  | "copy_template_sheet"
  | "first_open_reviewed"
  | "sort_table"
  | "filter_range"
  | "autofit_columns"
  | "clear_values"
  | "format_range";

export interface AgentActionHandlerDefinition {
  id: AgentActionHandlerId;
  capabilityName: string;
  intentAction?: AgentIntentAction;
  requiresResolvedTarget: boolean;
  riskKind:
    | "read_only"
    | "safe_format"
    | "broad_range_write"
    | "structure_change"
    | "destructive";
  matches: (input: AgentRunInput, request: string) => boolean;
}

export const AGENT_ACTION_HANDLERS: AgentActionHandlerDefinition[] = [
  {
    id: "save_workbook",
    capabilityName: "excel.workbook.save",
    intentAction: "save",
    requiresResolvedTarget: false,
    riskKind: "destructive",
    matches: (_input, request) => /\b(save)\b/.test(request)
  },
  {
    id: "calculate_workbook",
    capabilityName: "excel.workbook.calculate",
    intentAction: "calculate",
    requiresResolvedTarget: false,
    riskKind: "read_only",
    matches: (_input, request) => /\b(recalculate|calculate|calculation)\b/.test(request)
  },
  {
    id: "copy_template_sheet",
    capabilityName: "excel.template.create_sheet_from_template",
    intentAction: "copy_template_sheet",
    requiresResolvedTarget: false,
    riskKind: "structure_change",
    matches: (_input, request) => /\btemplate\b/.test(request) || (/\b(duplicate|copy)\b/.test(request) && /\bsheet\b/.test(request))
  },
  {
    id: "first_open_reviewed",
    capabilityName: "excel.range.write_values",
    requiresResolvedTarget: false,
    riskKind: "broad_range_write",
    matches: (input) => !input.values && /\bfirst\s+open\b/i.test(input.request) && /\breviewed\b/i.test(input.request)
  },
  {
    id: "sort_table",
    capabilityName: "excel.table.sort",
    intentAction: "sort_table",
    requiresResolvedTarget: false,
    riskKind: "broad_range_write",
    matches: (_input, request) => /\b(sort)\b/.test(request)
  },
  {
    id: "filter_range",
    capabilityName: "excel.table.apply_filters",
    intentAction: "filter_range",
    requiresResolvedTarget: true,
    riskKind: "safe_format",
    matches: (_input, request) => /\b(filter|filters)\b/.test(request)
  },
  {
    id: "autofit_columns",
    capabilityName: "excel.range.autofit_columns",
    intentAction: "autofit",
    requiresResolvedTarget: true,
    riskKind: "safe_format",
    matches: (_input, request) => /\b(autofit|auto\s*fit)\b/.test(request)
  },
  {
    id: "clear_values",
    capabilityName: "excel.range.clear_values_keep_format",
    intentAction: "clear_values",
    requiresResolvedTarget: true,
    riskKind: "destructive",
    matches: (_input, request) => /\b(clear|remove|delete|wipe)\b/.test(request) && /\b(data|values?|contents?|test data|input data)\b/.test(request)
  },
  {
    id: "format_range",
    capabilityName: "excel.range.write_styles",
    intentAction: "format_range",
    requiresResolvedTarget: true,
    riskKind: "safe_format",
    matches: (_input, request) => /\b(style|format|formatting|header\s+row)\b/.test(request)
  }
];

export function findAgentActionHandler(input: AgentRunInput, action: AgentIntentAction | undefined, requiresResolvedTarget: boolean): AgentActionHandlerDefinition | undefined {
  const request = input.request.toLowerCase();
  return AGENT_ACTION_HANDLERS.find((handler) =>
    handler.requiresResolvedTarget === requiresResolvedTarget &&
    ((action !== undefined && handler.intentAction === action) || handler.matches(input, request))
  );
}
