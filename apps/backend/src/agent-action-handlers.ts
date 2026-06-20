import type { AgentRunInput } from "@components-kit/open-workbook-protocol";
import type { AgentIntentAction } from "./agent-intent.js";

export type AgentActionHandlerId =
  | "save_workbook"
  | "calculate_workbook"
  | "create_snapshot"
  | "create_backup"
  | "refresh_snapshot"
  | "invalidate_snapshot"
  | "delete_snapshot"
  | "create_file_backup"
  | "restore_file_backup"
  | "prune_backups"
  | "pin_backup"
  | "unpin_backup"
  | "delete_backup"
  | "restore_workbook_backup"
  | "import_local_config"
  | "embed_local_config"
  | "import_embedded_local_config"
  | "close_workbook"
  | "copy_formula_patterns"
  | "fill_formula_down"
  | "fill_formula_right"
  | "repair_formula_patterns"
  | "convert_formulas_to_values"
  | "recalculate_formulas"
  | "create_name"
  | "update_name"
  | "delete_name"
  | "register_region"
  | "clear_region_values"
  | "write_region_values"
  | "fill_region"
  | "copy_template_sheet"
  | "register_template"
  | "unregister_template"
  | "clear_template_data_regions"
  | "fill_template_regions"
  | "repair_sheet_from_template"
  | "first_open_reviewed"
  | "write_formulas"
  | "write_number_formats"
  | "clear_style_dimensions"
  | "copy_style_from_template"
  | "repair_style_consistency"
  | "repair_style_from_template"
  | "repair_formulas_from_template"
  | "repair_table_structure"
  | "normalize_headers"
  | "trim_whitespace"
  | "remove_duplicates"
  | "parse_dates"
  | "parse_numbers"
  | "standardize_currency"
  | "fill_missing_values"
  | "split_column"
  | "merge_columns"
  | "sort_table"
  | "filter_range"
  | "apply_table_view"
  | "autofit_columns"
  | "autofit_rows"
  | "clear_range"
  | "clear_values"
  | "clear_values_raw"
  | "clear_formats"
  | "copy_range"
  | "move_range"
  | "write_styles_many"
  | "insert_rows"
  | "delete_rows"
  | "insert_columns"
  | "delete_columns"
  | "merge_range"
  | "unmerge_range"
  | "format_range"
  | "append_table_rows"
  | "update_table_rows"
  | "create_table"
  | "resize_table"
  | "reorder_table_columns"
  | "clear_table_data"
  | "clear_table_filters"
  | "set_table_total_row"
  | "set_table_style"
  | "copy_table_structure"
  | "create_sheet"
  | "copy_sheet"
  | "rename_sheet"
  | "delete_sheet"
  | "hide_sheet"
  | "unhide_sheet"
  | "protect_sheet"
  | "unprotect_sheet"
  | "clear_sheet"
  | "set_sheet_tab_color";

export interface AgentActionHandlerDefinition {
  id: AgentActionHandlerId;
  capabilityName: string;
  intentAction?: AgentIntentAction;
  requiresResolvedTarget: boolean;
  riskKind:
    | "read_only"
    | "safe_format"
    | "formula_write"
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
    id: "create_snapshot",
    capabilityName: "excel.workbook.snapshot",
    intentAction: "create_snapshot",
    requiresResolvedTarget: false,
    riskKind: "read_only",
    matches: (_input, request) => /\b(snapshot)\b/.test(request)
  },
  {
    id: "create_backup",
    capabilityName: "excel.workbook.create_backup",
    intentAction: "create_backup",
    requiresResolvedTarget: false,
    riskKind: "read_only",
    matches: (_input, request) => /\bbackup\b/.test(request) && /\b(create|capture|make|save)\b/.test(request)
  },
  {
    id: "refresh_snapshot",
    capabilityName: "excel.snapshot.refresh",
    intentAction: "refresh_snapshot",
    requiresResolvedTarget: false,
    riskKind: "read_only",
    matches: (_input, request) => /\bsnapshot\b/.test(request) && /\b(refresh|update|recapture)\b/.test(request)
  },
  {
    id: "invalidate_snapshot",
    capabilityName: "excel.snapshot.invalidate",
    intentAction: "invalidate_snapshot",
    requiresResolvedTarget: false,
    riskKind: "safe_format",
    matches: (_input, request) => /\bsnapshot\b/.test(request) && /\binvalidate\b/.test(request)
  },
  {
    id: "delete_snapshot",
    capabilityName: "excel.snapshot.delete",
    intentAction: "delete_snapshot",
    requiresResolvedTarget: false,
    riskKind: "destructive",
    matches: (_input, request) => /\bsnapshot\b/.test(request) && /\b(delete|remove)\b/.test(request)
  },
  {
    id: "create_file_backup",
    capabilityName: "excel.backup.create_file",
    intentAction: "create_file_backup",
    requiresResolvedTarget: false,
    riskKind: "read_only",
    matches: (_input, request) => /\b(file|full workbook)\b/.test(request) && /\bbackup\b/.test(request) && /\b(create|capture|make|save)\b/.test(request)
  },
  {
    id: "restore_file_backup",
    capabilityName: "excel.backup.restore_file",
    intentAction: "restore_file_backup",
    requiresResolvedTarget: false,
    riskKind: "destructive",
    matches: (_input, request) => /\b(file )?backup\b/.test(request) && /\b(restore|recover|open)\b/.test(request)
  },
  {
    id: "prune_backups",
    capabilityName: "excel.backup.prune",
    intentAction: "prune_backups",
    requiresResolvedTarget: false,
    riskKind: "destructive",
    matches: (_input, request) => /\bbackups?\b/.test(request) && /\b(prune|retention|cleanup|clean up)\b/.test(request)
  },
  {
    id: "pin_backup",
    capabilityName: "excel.backup.pin",
    intentAction: "pin_backup",
    requiresResolvedTarget: false,
    riskKind: "safe_format",
    matches: (_input, request) => /\bbackup\b/.test(request) && /\bpin\b/.test(request)
  },
  {
    id: "unpin_backup",
    capabilityName: "excel.backup.unpin",
    intentAction: "unpin_backup",
    requiresResolvedTarget: false,
    riskKind: "safe_format",
    matches: (_input, request) => /\bbackup\b/.test(request) && /\bunpin\b/.test(request)
  },
  {
    id: "delete_backup",
    capabilityName: "excel.backup.delete",
    intentAction: "delete_backup",
    requiresResolvedTarget: false,
    riskKind: "destructive",
    matches: (_input, request) => /\bbackup\b/.test(request) && /\b(delete|remove)\b/.test(request)
  },
  {
    id: "restore_workbook_backup",
    capabilityName: "excel.workbook.restore_backup",
    intentAction: "restore_workbook_backup",
    requiresResolvedTarget: false,
    riskKind: "destructive",
    matches: (_input, request) => /\b(restore|rollback)\b/.test(request) && /\bworkbook\b/.test(request) && /\bbackup\b/.test(request)
  },
  {
    id: "import_local_config",
    capabilityName: "excel.workbook.import_local_config",
    intentAction: "import_local_config",
    requiresResolvedTarget: false,
    riskKind: "structure_change",
    matches: (_input, request) => /\b(import|load)\b/.test(request) && /\blocal config\b/.test(request) && !/\bembedded\b/.test(request)
  },
  {
    id: "embed_local_config",
    capabilityName: "excel.workbook.embed_local_config",
    intentAction: "embed_local_config",
    requiresResolvedTarget: false,
    riskKind: "structure_change",
    matches: (_input, request) => /\b(embed|write)\b/.test(request) && /\blocal config\b/.test(request)
  },
  {
    id: "import_embedded_local_config",
    capabilityName: "excel.workbook.import_embedded_local_config",
    intentAction: "import_embedded_local_config",
    requiresResolvedTarget: false,
    riskKind: "structure_change",
    matches: (_input, request) => /\bimport\b/.test(request) && /\bembedded local config\b/.test(request)
  },
  {
    id: "close_workbook",
    capabilityName: "excel.workbook.close",
    intentAction: "close_workbook",
    requiresResolvedTarget: false,
    riskKind: "destructive",
    matches: (_input, request) => /\bclose\b/.test(request) && /\bworkbook\b/.test(request)
  },
  {
    id: "copy_formula_patterns",
    capabilityName: "excel.formula.copy_patterns",
    intentAction: "copy_formula_patterns",
    requiresResolvedTarget: false,
    riskKind: "formula_write",
    matches: (_input, request) => /\b(copy|replicate)\b/.test(request) && /\bformula patterns?\b/.test(request)
  },
  {
    id: "fill_formula_down",
    capabilityName: "excel.formula.fill_down",
    intentAction: "fill_formula_down",
    requiresResolvedTarget: false,
    riskKind: "formula_write",
    matches: (_input, request) => /\bfill\b/.test(request) && /\bformula/.test(request) && /\bdown\b/.test(request)
  },
  {
    id: "fill_formula_right",
    capabilityName: "excel.formula.fill_right",
    intentAction: "fill_formula_right",
    requiresResolvedTarget: false,
    riskKind: "formula_write",
    matches: (_input, request) => /\bfill\b/.test(request) && /\bformula/.test(request) && /\bright\b/.test(request)
  },
  {
    id: "repair_formula_patterns",
    capabilityName: "excel.formula.repair_patterns",
    intentAction: "repair_formula_patterns",
    requiresResolvedTarget: false,
    riskKind: "formula_write",
    matches: (_input, request) => /\b(repair|fix)\b/.test(request) && /\bformula patterns?\b/.test(request)
  },
  {
    id: "convert_formulas_to_values",
    capabilityName: "excel.formula.convert_to_values",
    intentAction: "convert_formulas_to_values",
    requiresResolvedTarget: true,
    riskKind: "destructive",
    matches: (_input, request) => /\b(convert|replace)\b/.test(request) && /\bformulas?\b/.test(request) && /\bvalues?\b/.test(request)
  },
  {
    id: "recalculate_formulas",
    capabilityName: "excel.formula.recalculate",
    intentAction: "recalculate_formulas",
    requiresResolvedTarget: false,
    riskKind: "read_only",
    matches: (_input, request) => /\b(recalculate|calculate)\b/.test(request) && /\bformulas?\b/.test(request)
  },
  {
    id: "create_name",
    capabilityName: "excel.names.create",
    intentAction: "create_name",
    requiresResolvedTarget: false,
    riskKind: "structure_change",
    matches: (_input, request) => /\b(create|add)\b/.test(request) && /\b(named range|name|named item)\b/.test(request)
  },
  {
    id: "update_name",
    capabilityName: "excel.names.update",
    intentAction: "update_name",
    requiresResolvedTarget: false,
    riskKind: "structure_change",
    matches: (_input, request) => /\b(update|change|retarget)\b/.test(request) && /\b(named range|name|named item)\b/.test(request)
  },
  {
    id: "delete_name",
    capabilityName: "excel.names.delete",
    intentAction: "delete_name",
    requiresResolvedTarget: false,
    riskKind: "destructive",
    matches: (_input, request) => /\b(delete|remove)\b/.test(request) && /\b(named range|name|named item)\b/.test(request)
  },
  {
    id: "register_region",
    capabilityName: "excel.region.register",
    intentAction: "register_region",
    requiresResolvedTarget: false,
    riskKind: "structure_change",
    matches: (_input, request) => /\b(register|create|add)\b/.test(request) && /\bregion\b/.test(request)
  },
  {
    id: "clear_region_values",
    capabilityName: "excel.region.clear_values",
    intentAction: "clear_region_values",
    requiresResolvedTarget: false,
    riskKind: "destructive",
    matches: (_input, request) => /\b(clear|remove|wipe)\b/.test(request) && /\bregion\b/.test(request) && /\b(values?|contents?|data)\b/.test(request)
  },
  {
    id: "write_region_values",
    capabilityName: "excel.region.write_values",
    intentAction: "write_region_values",
    requiresResolvedTarget: false,
    riskKind: "broad_range_write",
    matches: (_input, request) => /\b(write|update|set)\b/.test(request) && /\bregion\b/.test(request)
  },
  {
    id: "fill_region",
    capabilityName: "excel.region.fill",
    intentAction: "fill_region",
    requiresResolvedTarget: false,
    riskKind: "broad_range_write",
    matches: (_input, request) => /\b(fill|populate)\b/.test(request) && /\bregion\b/.test(request)
  },
  {
    id: "copy_template_sheet",
    capabilityName: "excel.template.create_sheet_from_template",
    intentAction: "copy_template_sheet",
    requiresResolvedTarget: false,
    riskKind: "structure_change",
    matches: (input, request) =>
      input.intent?.action !== "copy_sheet" &&
      !/\b(style|format|formatting)\b/.test(request) &&
      ((/\b(create|new|copy|duplicate)\b/.test(request) && /\btemplate\b/.test(request)) ||
        (/\b(copy|duplicate)\b/.test(request) && /\bsheet\b/.test(request)))
  },
  {
    id: "register_template",
    capabilityName: "excel.template.register",
    intentAction: "register_template",
    requiresResolvedTarget: false,
    riskKind: "structure_change",
    matches: (_input, request) => /\b(register|capture|save)\b/.test(request) && /\btemplate\b/.test(request)
  },
  {
    id: "unregister_template",
    capabilityName: "excel.template.unregister",
    intentAction: "unregister_template",
    requiresResolvedTarget: false,
    riskKind: "destructive",
    matches: (_input, request) => /\b(unregister|delete|remove)\b/.test(request) && /\btemplate\b/.test(request)
  },
  {
    id: "repair_sheet_from_template",
    capabilityName: "excel.template.repair_sheet_from_template",
    intentAction: "repair_sheet_from_template",
    requiresResolvedTarget: false,
    riskKind: "structure_change",
    matches: (_input, request) => /\b(repair|fix)\b/.test(request) && /\btemplate\b/.test(request) && !/\b(style|format|formatting|formula)\b/.test(request)
  },
  {
    id: "clear_template_data_regions",
    capabilityName: "excel.template.clear_data_regions",
    intentAction: "clear_template_data_regions",
    requiresResolvedTarget: false,
    riskKind: "destructive",
    matches: (_input, request) => /\b(clear|empty)\b/.test(request) && /\b(data regions?|template regions?)\b/.test(request)
  },
  {
    id: "fill_template_regions",
    capabilityName: "excel.template.fill_regions",
    intentAction: "fill_template_regions",
    requiresResolvedTarget: false,
    riskKind: "broad_range_write",
    matches: (_input, request) => /\b(fill|populate)\b/.test(request) && /\b(data regions?|template regions?)\b/.test(request)
  },
  {
    id: "copy_style_from_template",
    capabilityName: "excel.style.copy_from_template",
    intentAction: "copy_style_from_template",
    requiresResolvedTarget: false,
    riskKind: "safe_format",
    matches: (_input, request) => /\b(copy|match|apply)\b/.test(request)
      && /\b(style|format|formatting|look)\b/.test(request)
      && /\b(template|same|source|from|to|target|like|as)\b/.test(request)
  },
  {
    id: "repair_style_consistency",
    capabilityName: "excel.style.repair_consistency",
    intentAction: "repair_style_consistency",
    requiresResolvedTarget: false,
    riskKind: "safe_format",
    matches: (_input, request) => /\b(repair|fix)\b/.test(request) && /\b(style|format|formatting)\b/.test(request)
  },
  {
    id: "repair_style_from_template",
    capabilityName: "excel.repair.style_from_template",
    intentAction: "repair_style_from_template",
    requiresResolvedTarget: false,
    riskKind: "safe_format",
    matches: (_input, request) => /\b(repair|fix)\b/.test(request) && /\b(style|format|formatting)\b/.test(request) && /\btemplate\b/.test(request)
  },
  {
    id: "repair_formulas_from_template",
    capabilityName: "excel.repair.formulas_from_template",
    intentAction: "repair_formulas_from_template",
    requiresResolvedTarget: false,
    riskKind: "formula_write",
    matches: (_input, request) => /\b(repair|fix)\b/.test(request) && /\bformulas?\b/.test(request) && /\btemplate\b/.test(request)
  },
  {
    id: "create_sheet",
    capabilityName: "excel.sheet.create",
    intentAction: "create_sheet",
    requiresResolvedTarget: false,
    riskKind: "structure_change",
    matches: (_input, request) => /\b(create|add|new)\b/.test(request) && /\bsheet\b/.test(request)
  },
  {
    id: "copy_sheet",
    capabilityName: "excel.sheet.copy",
    intentAction: "copy_sheet",
    requiresResolvedTarget: false,
    riskKind: "structure_change",
    matches: (_input, request) => /\b(copy|duplicate)\b/.test(request) && /\bsheet\b/.test(request) && !/\btemplate\b/.test(request)
  },
  {
    id: "rename_sheet",
    capabilityName: "excel.sheet.rename",
    intentAction: "rename_sheet",
    requiresResolvedTarget: false,
    riskKind: "structure_change",
    matches: (_input, request) => /\b(rename)\b/.test(request) && /\bsheet\b/.test(request)
  },
  {
    id: "delete_sheet",
    capabilityName: "excel.sheet.delete",
    intentAction: "delete_sheet",
    requiresResolvedTarget: false,
    riskKind: "destructive",
    matches: (_input, request) => /\b(delete|remove)\b/.test(request) && /\bsheet\b/.test(request)
  },
  {
    id: "hide_sheet",
    capabilityName: "excel.sheet.hide",
    intentAction: "hide_sheet",
    requiresResolvedTarget: false,
    riskKind: "structure_change",
    matches: (_input, request) => /\b(hide)\b/.test(request) && /\bsheet\b/.test(request)
  },
  {
    id: "unhide_sheet",
    capabilityName: "excel.sheet.unhide",
    intentAction: "unhide_sheet",
    requiresResolvedTarget: false,
    riskKind: "structure_change",
    matches: (_input, request) => /\b(unhide|show)\b/.test(request) && /\bsheet\b/.test(request)
  },
  {
    id: "protect_sheet",
    capabilityName: "excel.sheet.protect",
    intentAction: "protect_sheet",
    requiresResolvedTarget: false,
    riskKind: "structure_change",
    matches: (_input, request) => /\bprotect\b/.test(request) && /\bsheet\b/.test(request) && !/\bunprotect\b/.test(request)
  },
  {
    id: "unprotect_sheet",
    capabilityName: "excel.sheet.unprotect",
    intentAction: "unprotect_sheet",
    requiresResolvedTarget: false,
    riskKind: "structure_change",
    matches: (_input, request) => /\b(unprotect|unlock)\b/.test(request) && /\bsheet\b/.test(request)
  },
  {
    id: "clear_sheet",
    capabilityName: "excel.sheet.clear",
    intentAction: "clear_sheet",
    requiresResolvedTarget: false,
    riskKind: "destructive",
    matches: (_input, request) => /\b(clear|wipe)\b/.test(request) && /\bsheet\b/.test(request)
  },
  {
    id: "set_sheet_tab_color",
    capabilityName: "excel.sheet.set_tab_color",
    intentAction: "set_sheet_tab_color",
    requiresResolvedTarget: false,
    riskKind: "safe_format",
    matches: (_input, request) => /\b(tab color|sheet color|color tab)\b/.test(request)
  },
  {
    id: "first_open_reviewed",
    capabilityName: "excel.range.write_values",
    requiresResolvedTarget: false,
    riskKind: "broad_range_write",
    matches: (input) => !input.values && /\bfirst\s+open\b/i.test(input.request) && /\breviewed\b/i.test(input.request)
  },
  {
    id: "write_formulas",
    capabilityName: "excel.range.write_formulas",
    intentAction: "write_formulas",
    requiresResolvedTarget: true,
    riskKind: "destructive",
    matches: (input, request) => /\b(write|copy|fill|update|fix)\b/.test(request) && /\b(formula|formulas)\b/.test(request) && hasFormulaLikeValue(input.values)
  },
  {
    id: "write_number_formats",
    capabilityName: "excel.range.write_number_formats",
    intentAction: "write_number_formats",
    requiresResolvedTarget: true,
    riskKind: "safe_format",
    matches: (_input, request) => /\b(number\s+format|number\s+formats|currency|percent|percentage|decimal|date format)\b/.test(request)
  },
  {
    id: "clear_style_dimensions",
    capabilityName: "excel.range.clear_style_dimensions",
    intentAction: "clear_style_dimensions",
    requiresResolvedTarget: true,
    riskKind: "safe_format",
    matches: (_input, request) => /\b(clear|remove|delete|wipe)\b/.test(request) && /\b(borders?|fills?|fonts?|alignment|number\s*formats?|row heights?|column widths?)\b/.test(request) && !/\b(all formats?|all formatting|everything)\b/.test(request)
  },
  {
    id: "apply_table_view",
    capabilityName: "excel.table.apply_view",
    intentAction: "apply_table_view",
    requiresResolvedTarget: true,
    riskKind: "broad_range_write",
    matches: (_input, request) => /\b(filter|filters)\b/.test(request) && /\b(sort)\b/.test(request)
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
    id: "append_table_rows",
    capabilityName: "excel.table.append_rows",
    intentAction: "append_table_rows",
    requiresResolvedTarget: true,
    riskKind: "broad_range_write",
    matches: (_input, request) => /\b(append|add|insert)\b/.test(request) && /\b(rows?|records?|table)\b/.test(request) && !/\bborders?\b/.test(request)
  },
  {
    id: "update_table_rows",
    capabilityName: "excel.table.update_rows",
    intentAction: "update_table_rows",
    requiresResolvedTarget: true,
    riskKind: "broad_range_write",
    matches: (_input, request) => /\b(update|change|edit)\b/.test(request) && /\b(rows?|records?|table)\b/.test(request)
  },
  {
    id: "create_table",
    capabilityName: "excel.table.create",
    intentAction: "create_table",
    requiresResolvedTarget: false,
    riskKind: "structure_change",
    matches: (_input, request) => /\b(create|add|make)\b/.test(request) && /\btable\b/.test(request)
  },
  {
    id: "resize_table",
    capabilityName: "excel.table.resize",
    intentAction: "resize_table",
    requiresResolvedTarget: true,
    riskKind: "structure_change",
    matches: (_input, request) => /\b(resize|expand|shrink)\b/.test(request) && /\btable\b/.test(request)
  },
  {
    id: "reorder_table_columns",
    capabilityName: "excel.table.reorder_columns",
    intentAction: "reorder_table_columns",
    requiresResolvedTarget: true,
    riskKind: "structure_change",
    matches: (_input, request) => /\b(reorder|move|rearrange|swap)\b/.test(request) && /\b(columns?|table columns?)\b/.test(request)
  },
  {
    id: "clear_table_data",
    capabilityName: "excel.table.clear_data_keep_formulas",
    intentAction: "clear_table_data",
    requiresResolvedTarget: true,
    riskKind: "destructive",
    matches: (_input, request) => /\b(clear|remove|wipe)\b/.test(request) && /\b(table data|table rows|data rows)\b/.test(request)
  },
  {
    id: "clear_table_filters",
    capabilityName: "excel.table.clear_filters",
    intentAction: "clear_table_filters",
    requiresResolvedTarget: true,
    riskKind: "safe_format",
    matches: (_input, request) => /\b(clear|remove)\b/.test(request) && /\b(filters?|table filters?)\b/.test(request)
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
    id: "set_table_total_row",
    capabilityName: "excel.table.set_total_row",
    intentAction: "set_table_total_row",
    requiresResolvedTarget: true,
    riskKind: "safe_format",
    matches: (_input, request) => /\b(total row|totals row|show totals|hide totals)\b/.test(request)
  },
  {
    id: "set_table_style",
    capabilityName: "excel.table.set_style",
    intentAction: "set_table_style",
    requiresResolvedTarget: true,
    riskKind: "safe_format",
    matches: (_input, request) => /\b(table style|style table|format table)\b/.test(request)
  },
  {
    id: "copy_table_structure",
    capabilityName: "excel.table.copy_structure",
    intentAction: "copy_table_structure",
    requiresResolvedTarget: true,
    riskKind: "structure_change",
    matches: (_input, request) => /\b(copy|duplicate)\b/.test(request) && /\btable structure\b/.test(request)
  },
  {
    id: "repair_table_structure",
    capabilityName: "excel.repair.table_structure",
    intentAction: "repair_table_structure",
    requiresResolvedTarget: true,
    riskKind: "structure_change",
    matches: (_input, request) => /\b(repair|fix)\b/.test(request) && /\btable structure\b/.test(request)
  },
  {
    id: "autofit_columns",
    capabilityName: "excel.range.autofit_columns",
    intentAction: "autofit",
    requiresResolvedTarget: true,
    riskKind: "safe_format",
    matches: (_input, request) => /\b(autofit|auto\s*fit)\b/.test(request) && !/\b(rows?|height)\b/.test(request)
  },
  {
    id: "autofit_rows",
    capabilityName: "excel.range.autofit_rows",
    intentAction: "autofit_rows",
    requiresResolvedTarget: true,
    riskKind: "safe_format",
    matches: (_input, request) => /\b(autofit|auto\s*fit)\b/.test(request) && /\b(rows?|height)\b/.test(request)
  },
  {
    id: "clear_range",
    capabilityName: "excel.range.clear",
    intentAction: "clear_range",
    requiresResolvedTarget: true,
    riskKind: "destructive",
    matches: (_input, request) => /\b(clear|remove|delete|wipe)\b/.test(request) && /\b(all|everything|range|cells?)\b/.test(request)
  },
  {
    id: "normalize_headers",
    capabilityName: "excel.clean.normalize_headers",
    intentAction: "normalize_headers",
    requiresResolvedTarget: true,
    riskKind: "broad_range_write",
    matches: (_input, request) => /\b(normalize|clean|standardize)\b/.test(request) && /\bheaders?\b/.test(request)
  },
  {
    id: "trim_whitespace",
    capabilityName: "excel.clean.trim_whitespace",
    intentAction: "trim_whitespace",
    requiresResolvedTarget: true,
    riskKind: "broad_range_write",
    matches: (_input, request) => /\b(trim|clean)\b/.test(request) && /\b(whitespace|spaces?)\b/.test(request)
  },
  {
    id: "remove_duplicates",
    capabilityName: "excel.clean.remove_duplicates",
    intentAction: "remove_duplicates",
    requiresResolvedTarget: true,
    riskKind: "destructive",
    matches: (_input, request) => /\b(remove|delete|dedupe|de-duplicate)\b/.test(request) && /\bduplicates?\b/.test(request)
  },
  {
    id: "parse_dates",
    capabilityName: "excel.clean.parse_dates",
    intentAction: "parse_dates",
    requiresResolvedTarget: true,
    riskKind: "broad_range_write",
    matches: (_input, request) => /\b(parse|convert|normalize|clean)\b/.test(request) && /\bdates?\b/.test(request)
  },
  {
    id: "parse_numbers",
    capabilityName: "excel.clean.parse_numbers",
    intentAction: "parse_numbers",
    requiresResolvedTarget: true,
    riskKind: "broad_range_write",
    matches: (_input, request) => /\b(parse|convert|normalize|clean)\b/.test(request) && /\bnumbers?\b/.test(request)
  },
  {
    id: "standardize_currency",
    capabilityName: "excel.clean.standardize_currency",
    intentAction: "standardize_currency",
    requiresResolvedTarget: true,
    riskKind: "broad_range_write",
    matches: (_input, request) => /\b(standardize|parse|convert|clean)\b/.test(request) && /\bcurrenc(y|ies)\b/.test(request)
  },
  {
    id: "fill_missing_values",
    capabilityName: "excel.clean.fill_missing_values",
    intentAction: "fill_missing_values",
    requiresResolvedTarget: true,
    riskKind: "broad_range_write",
    matches: (_input, request) => /\b(fill|replace)\b/.test(request) && /\b(missing|blank|empty)\b/.test(request)
  },
  {
    id: "split_column",
    capabilityName: "excel.clean.split_column",
    intentAction: "split_column",
    requiresResolvedTarget: true,
    riskKind: "broad_range_write",
    matches: (_input, request) => /\bsplit\b/.test(request) && /\bcolumn\b/.test(request)
  },
  {
    id: "merge_columns",
    capabilityName: "excel.clean.merge_columns",
    intentAction: "merge_columns",
    requiresResolvedTarget: true,
    riskKind: "broad_range_write",
    matches: (_input, request) => /\b(merge|combine)\b/.test(request) && /\bcolumns?\b/.test(request)
  },
  {
    id: "clear_formats",
    capabilityName: "excel.range.clear_formats",
    intentAction: "clear_formats",
    requiresResolvedTarget: true,
    riskKind: "safe_format",
    matches: (_input, request) => /\b(clear|remove|delete|wipe)\b/.test(request) && /\b(formats?|formatting|styles?)\b/.test(request)
  },
  {
    id: "clear_values",
    capabilityName: "excel.range.clear_values_keep_format",
    intentAction: "clear_values",
    requiresResolvedTarget: true,
    riskKind: "destructive",
    matches: (_input, request) => /\b(clear|remove|delete|wipe)\b/.test(request) && /\b(data|values?|contents?|test data|input data)\b/.test(request) && !/\b(formats?|formatting|styles?)\b/.test(request)
  },
  {
    id: "clear_values_raw",
    capabilityName: "excel.range.clear_values",
    intentAction: "clear_values_raw",
    requiresResolvedTarget: true,
    riskKind: "destructive",
    matches: (_input, request) => /\b(clear|remove|delete|wipe)\b/.test(request) && /\braw values?|contents?\b/.test(request)
  },
  {
    id: "copy_range",
    capabilityName: "excel.range.copy",
    intentAction: "copy_range",
    requiresResolvedTarget: false,
    riskKind: "broad_range_write",
    matches: (_input, request) => /\b(copy|duplicate)\b/.test(request) && /\b(range|cells?)\b/.test(request)
  },
  {
    id: "move_range",
    capabilityName: "excel.range.move",
    intentAction: "move_range",
    requiresResolvedTarget: false,
    riskKind: "destructive",
    matches: (_input, request) => /\b(move|relocate)\b/.test(request) && /\b(range|cells?)\b/.test(request)
  },
  {
    id: "write_styles_many",
    capabilityName: "excel.range.write_styles_many",
    intentAction: "write_styles_many",
    requiresResolvedTarget: false,
    riskKind: "safe_format",
    matches: (_input, request) => /\b(format|style)\b/.test(request) && /\b(multiple|many|several|ranges?)\b/.test(request)
  },
  {
    id: "insert_rows",
    capabilityName: "excel.range.insert_rows",
    intentAction: "insert_rows",
    requiresResolvedTarget: true,
    riskKind: "structure_change",
    matches: (_input, request) => /\binsert\b/.test(request) && /\brows?\b/.test(request)
  },
  {
    id: "delete_rows",
    capabilityName: "excel.range.delete_rows",
    intentAction: "delete_rows",
    requiresResolvedTarget: true,
    riskKind: "destructive",
    matches: (_input, request) => /\b(delete|remove)\b/.test(request) && /\brows?\b/.test(request)
  },
  {
    id: "insert_columns",
    capabilityName: "excel.range.insert_columns",
    intentAction: "insert_columns",
    requiresResolvedTarget: true,
    riskKind: "structure_change",
    matches: (_input, request) => /\binsert\b/.test(request) && /\bcolumns?\b/.test(request)
  },
  {
    id: "delete_columns",
    capabilityName: "excel.range.delete_columns",
    intentAction: "delete_columns",
    requiresResolvedTarget: true,
    riskKind: "destructive",
    matches: (_input, request) => /\b(delete|remove)\b/.test(request) && /\bcolumns?\b/.test(request)
  },
  {
    id: "merge_range",
    capabilityName: "excel.range.merge",
    intentAction: "merge_range",
    requiresResolvedTarget: true,
    riskKind: "structure_change",
    matches: (_input, request) => /\bmerge\b/.test(request) && /\b(range|cells?)\b/.test(request)
  },
  {
    id: "unmerge_range",
    capabilityName: "excel.range.unmerge",
    intentAction: "unmerge_range",
    requiresResolvedTarget: true,
    riskKind: "structure_change",
    matches: (_input, request) => /\bunmerge\b/.test(request)
  },
  {
    id: "format_range",
    capabilityName: "excel.range.write_styles",
    intentAction: "format_range",
    requiresResolvedTarget: true,
    riskKind: "safe_format",
    matches: (_input, request) => /\b(style|format|formatting|header\s+row|borders?)\b/.test(request)
  }
];

export function findAgentActionHandler(input: AgentRunInput, action: AgentIntentAction | undefined, requiresResolvedTarget: boolean): AgentActionHandlerDefinition | undefined {
  const request = input.request.toLowerCase();
  const scopeHandlers = AGENT_ACTION_HANDLERS.filter((handler) => handler.requiresResolvedTarget === requiresResolvedTarget);
  if (action !== undefined) {
    return scopeHandlers.find((handler) => handler.intentAction === action);
  }
  return scopeHandlers.find((handler) => handler.matches(input, request));
}

function hasFormulaLikeValue(values: AgentRunInput["values"]): boolean {
  if (!values) {
    return false;
  }
  const matrix = Array.isArray(values.values)
    ? values.values
    : Array.isArray(values.rows)
      ? values.rows
      : [Object.values(values)];
  return matrix.flat().some((value) => typeof value === "string" && value.trim().startsWith("="));
}
