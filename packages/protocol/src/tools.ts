import type { DestructiveLevel } from "./operations.js";

export type CatalogStatus = "stable" | "preview" | "planned" | "unsupported";
export type ToolNamespace =
  | "runtime"
  | "workbook"
  | "backup"
  | "sheet"
  | "range"
  | "batch"
  | "template"
  | "style"
  | "formula"
  | "table"
  | "filter"
  | "sort"
  | "pivot"
  | "chart"
  | "names"
  | "region"
  | "plan"
  | "task"
  | "collab"
  | "transaction"
  | "lock"
  | "conflict"
  | "diff"
  | "events"
  | "snapshot"
  | "validate"
  | "repair"
  | "clean"
  | "permissions";

export interface ToolContract {
  name: string;
  title: string;
  namespace: ToolNamespace;
  status: CatalogStatus;
  mutatesWorkbook: boolean;
  destructiveLevel: DestructiveLevel;
  requiresConfirmation: boolean;
  requiredCapabilities: string[];
  description: string;
}

export interface ToolCatalogOptions {
  includePreview?: boolean;
  includePlanned?: boolean;
  includeUnsupported?: boolean;
}

const STABLE_TOOLS = new Set([
  "excel.runtime.get_status",
  "excel.runtime.get_capabilities",
  "excel.runtime.get_active_context",
  "excel.workbook.list_open_workbooks",
  "excel.workbook.get_workbook_info",
  "excel.workbook.get_workbook_map",
  "excel.workbook.snapshot",
  "excel.workbook.refresh_snapshot",
  "excel.workbook.get_snapshot",
  "excel.workbook.detect_external_changes",
  "excel.workbook.calculate",
  "excel.workbook.save",
  "excel.workbook.save_as",
  "excel.workbook.create_backup",
  "excel.workbook.restore_backup",
  "excel.workbook.export_copy",
  "excel.backup.create_file",
  "excel.backup.list",
  "excel.backup.get",
  "excel.backup.verify",
  "excel.backup.restore_file",
  "excel.backup.delete",
  "excel.backup.prune",
  "excel.backup.pin",
  "excel.backup.unpin",
  "excel.workbook.export_local_config",
  "excel.workbook.import_local_config",
  "excel.workbook.embed_local_config",
  "excel.workbook.read_embedded_local_config",
  "excel.workbook.import_embedded_local_config",
  "excel.workbook.close",
  "excel.sheet.list",
  "excel.sheet.get_info",
  "excel.sheet.create",
  "excel.sheet.copy",
  "excel.sheet.rename",
  "excel.sheet.delete",
  "excel.sheet.hide",
  "excel.sheet.unhide",
  "excel.sheet.protect",
  "excel.sheet.unprotect",
  "excel.sheet.clear",
  "excel.sheet.get_used_range",
  "excel.sheet.set_tab_color",
  "excel.range.read_values",
  "excel.range.read_formulas",
  "excel.range.read_number_formats",
  "excel.range.read_display_text",
  "excel.range.read_styles",
  "excel.range.read_hyperlinks",
  "excel.range.read_comments",
  "excel.range.read_notes",
  "excel.range.read_merged_cells",
  "excel.range.read_data_validation",
  "excel.range.read_conditional_formatting",
  "excel.range.read_full",
  "excel.range.search",
  "excel.range.find_blank_cells",
  "excel.range.find_errors",
  "excel.range.write_values",
  "excel.range.write_formulas",
  "excel.range.write_number_formats",
  "excel.range.write_styles",
  "excel.range.clear",
  "excel.range.clear_values",
  "excel.range.clear_formats",
  "excel.range.clear_values_keep_format",
  "excel.range.copy",
  "excel.range.move",
  "excel.range.insert_rows",
  "excel.range.delete_rows",
  "excel.range.insert_columns",
  "excel.range.delete_columns",
  "excel.range.autofit_columns",
  "excel.range.autofit_rows",
  "excel.range.merge",
  "excel.range.unmerge",
  "excel.batch.validate",
  "excel.batch.dry_run",
  "excel.batch.apply",
  "excel.plan.create",
  "excel.plan.preview",
  "excel.plan.refresh_preview",
  "excel.plan.rebase",
  "excel.plan.apply",
  "excel.plan.rollback",
  "excel.task.create",
  "excel.task.claim",
  "excel.task.update",
  "excel.task.set_progress",
  "excel.task.add_blocker",
  "excel.task.resolve_blocker",
  "excel.task.evaluate_schedule",
  "excel.task.resume_ready",
  "excel.task.complete",
  "excel.task.fail",
  "excel.task.cancel",
  "excel.task.list",
  "excel.task.get",
  "excel.collab.get_status",
  "excel.collab.list_agents",
  "excel.collab.list_tasks",
  "excel.collab.list_locks",
  "excel.collab.list_transactions",
  "excel.collab.get_conflicts",
  "excel.collab.get_recent_events",
  "excel.lock.get_policy",
  "excel.lock.set_policy",
  "excel.lock.acquire",
  "excel.lock.renew",
  "excel.lock.release",
  "excel.conflict.get_guidance",
  "excel.conflict.explain",
  "excel.conflict.get_telemetry",
  "excel.conflict.clear_telemetry",
  "excel.transaction.get",
  "excel.transaction.list",
  "excel.transaction.preview_rollback",
  "excel.transaction.rollback",
  "excel.transaction.preview_rollback_chain",
  "excel.transaction.rollback_chain",
  "excel.diff.create",
  "excel.diff.summarize",
  "excel.diff.get_details",
  "excel.diff.export_json",
  "excel.diff.export_html",
  "excel.events.subscribe",
  "excel.events.unsubscribe",
  "excel.events.get_recent",
  "excel.events.clear",
  "excel.events.set_debounce",
  "excel.snapshot.create",
  "excel.snapshot.refresh",
  "excel.snapshot.get",
  "excel.snapshot.compare",
  "excel.snapshot.invalidate",
  "excel.snapshot.list",
  "excel.snapshot.delete",
  "excel.template.detect_templates",
  "excel.template.register",
  "excel.template.unregister",
  "excel.template.get",
  "excel.template.list",
  "excel.template.infer_regions",
  "excel.template.create_sheet_from_template",
  "excel.template.clear_data_regions",
  "excel.template.fill_regions",
  "excel.template.validate_sheet_against_template",
  "excel.template.repair_sheet_from_template",
  "excel.style.get_fingerprint",
  "excel.style.compare_fingerprint",
  "excel.style.copy_from_template",
  "excel.style.apply_style",
  "excel.style.validate_consistency",
  "excel.style.repair_consistency",
  "excel.style.get_theme",
  "excel.style.apply_theme",
  "excel.style.copy_column_widths",
  "excel.style.copy_row_heights",
  "excel.style.copy_borders",
  "excel.style.copy_fills",
  "excel.style.copy_fonts",
  "excel.style.copy_alignment",
  "excel.style.copy_number_formats",
  "excel.style.copy_conditional_formatting",
  "excel.style.copy_data_validation",
  "excel.style.copy_freeze_panes",
  "excel.style.copy_print_settings",
  "excel.style.copy_page_layout",
  "excel.style.copy_hidden_rows_columns",
  "excel.formula.read_patterns",
  "excel.formula.copy_patterns",
  "excel.formula.fill_down",
  "excel.formula.fill_right",
  "excel.formula.validate",
  "excel.formula.validate_against_template",
  "excel.formula.repair_patterns",
  "excel.formula.find_errors",
  "excel.formula.find_circular_references",
  "excel.formula.get_dependency_graph",
  "excel.formula.trace_precedents",
  "excel.formula.trace_dependents",
  "excel.formula.convert_to_values",
  "excel.formula.recalculate",
  "excel.formula.explain",
  "excel.table.list",
  "excel.table.get_info",
  "excel.table.read",
  "excel.table.create",
  "excel.table.resize",
  "excel.table.append_rows",
  "excel.table.update_rows",
  "excel.table.clear_data_keep_formulas",
  "excel.table.clear_filters",
  "excel.table.apply_filters",
  "excel.table.preserve_filters",
  "excel.table.sort",
  "excel.table.set_total_row",
  "excel.table.set_style",
  "excel.table.copy_structure",
  "excel.table.validate_against_template",
  "excel.filter.get_filters",
  "excel.filter.apply",
  "excel.filter.clear",
  "excel.filter.preserve_from_template",
  "excel.filter.validate",
  "excel.sort.apply",
  "excel.sort.clear",
  "excel.sort.preserve_from_template",
  "excel.pivot.list",
  "excel.pivot.get_info",
  "excel.pivot.create",
  "excel.pivot.refresh",
  "excel.pivot.refresh_all",
  "excel.pivot.update_source",
  "excel.pivot.copy_from_template",
  "excel.pivot.delete",
  "excel.pivot.validate_source",
  "excel.pivot.get_capability_matrix",
  "excel.pivot.get_fingerprint",
  "excel.pivot.compare_fingerprint",
  "excel.pivot.diff",
  "excel.pivot.repair_from_template",
  "excel.pivot.rebuild_with_source",
  "excel.chart.list",
  "excel.chart.get_info",
  "excel.chart.create",
  "excel.chart.update_data_source",
  "excel.chart.copy_from_template",
  "excel.chart.refresh",
  "excel.chart.delete",
  "excel.chart.validate_against_template",
  "excel.names.list",
  "excel.names.get",
  "excel.names.create",
  "excel.names.update",
  "excel.names.delete",
  "excel.region.detect",
  "excel.region.register",
  "excel.region.list",
  "excel.region.get",
  "excel.region.clear_values",
  "excel.region.write_values",
  "excel.region.fill",
  "excel.validate.workbook",
  "excel.validate.sheet",
  "excel.validate.template_consistency",
  "excel.validate.formulas",
  "excel.validate.styles",
  "excel.validate.tables",
  "excel.validate.filters",
  "excel.validate.print_layout",
  "excel.validate.no_broken_references",
  "excel.validate.no_formula_errors",
  "excel.validate.no_unintended_changes",
  "excel.repair.style_from_template",
  "excel.repair.formulas_from_template",
  "excel.repair.filters_from_template",
  "excel.repair.table_structure",
  "excel.repair.print_layout",
  "excel.repair.named_ranges",
  "excel.repair.formula_errors",
  "excel.repair.merged_cells",
  "excel.clean.detect_header_row",
  "excel.clean.normalize_headers",
  "excel.clean.trim_whitespace",
  "excel.clean.remove_duplicates",
  "excel.clean.parse_dates",
  "excel.clean.parse_numbers",
  "excel.clean.standardize_currency",
  "excel.clean.fill_missing_values",
  "excel.clean.split_column",
  "excel.clean.merge_columns",
  "excel.clean.detect_outliers",
  "excel.clean.fuzzy_match",
  "excel.permissions.get",
  "excel.permissions.set",
  "excel.permissions.require_confirmation",
  "excel.permissions.set_scope",
  "excel.permissions.allow_destructive_actions",
  "excel.permissions.allow_macro_execution",
  "excel.permissions.lock_regions",
  "excel.permissions.unlock_regions"
]);

const PREVIEW_TOOLS = new Set([
  "excel.runtime.connect_addin",
  "excel.runtime.disconnect_addin",
  "excel.runtime.ping_addin",
  "excel.runtime.get_selection",
  "excel.runtime.set_active_workbook",
  "excel.runtime.set_active_sheet"
]);

const TOOL_NAMES = [
  "excel.runtime.get_status",
  "excel.runtime.connect_addin",
  "excel.runtime.disconnect_addin",
  "excel.runtime.ping_addin",
  "excel.runtime.get_capabilities",
  "excel.runtime.get_active_context",
  "excel.runtime.get_selection",
  "excel.runtime.set_active_workbook",
  "excel.runtime.set_active_sheet",
  "excel.workbook.list_open_workbooks",
  "excel.workbook.get_workbook_info",
  "excel.workbook.get_workbook_map",
  "excel.workbook.snapshot",
  "excel.workbook.refresh_snapshot",
  "excel.workbook.get_snapshot",
  "excel.workbook.detect_external_changes",
  "excel.workbook.calculate",
  "excel.workbook.save",
  "excel.workbook.save_as",
  "excel.workbook.create_backup",
  "excel.workbook.restore_backup",
  "excel.workbook.export_copy",
  "excel.backup.create_file",
  "excel.backup.list",
  "excel.backup.get",
  "excel.backup.verify",
  "excel.backup.restore_file",
  "excel.backup.delete",
  "excel.backup.prune",
  "excel.backup.pin",
  "excel.backup.unpin",
  "excel.workbook.export_local_config",
  "excel.workbook.import_local_config",
  "excel.workbook.embed_local_config",
  "excel.workbook.read_embedded_local_config",
  "excel.workbook.import_embedded_local_config",
  "excel.workbook.close",
  "excel.sheet.list",
  "excel.sheet.get_info",
  "excel.sheet.create",
  "excel.sheet.copy",
  "excel.sheet.rename",
  "excel.sheet.delete",
  "excel.sheet.hide",
  "excel.sheet.unhide",
  "excel.sheet.protect",
  "excel.sheet.unprotect",
  "excel.sheet.clear",
  "excel.sheet.get_used_range",
  "excel.sheet.set_tab_color",
  "excel.range.read_values",
  "excel.range.read_formulas",
  "excel.range.read_number_formats",
  "excel.range.read_display_text",
  "excel.range.read_styles",
  "excel.range.read_hyperlinks",
  "excel.range.read_comments",
  "excel.range.read_notes",
  "excel.range.read_merged_cells",
  "excel.range.read_data_validation",
  "excel.range.read_conditional_formatting",
  "excel.range.read_full",
  "excel.range.search",
  "excel.range.find_blank_cells",
  "excel.range.find_errors",
  "excel.range.write_values",
  "excel.range.write_formulas",
  "excel.range.write_number_formats",
  "excel.range.write_styles",
  "excel.range.clear",
  "excel.range.clear_values",
  "excel.range.clear_formats",
  "excel.range.clear_values_keep_format",
  "excel.range.copy",
  "excel.range.move",
  "excel.range.insert_rows",
  "excel.range.delete_rows",
  "excel.range.insert_columns",
  "excel.range.delete_columns",
  "excel.range.autofit_columns",
  "excel.range.autofit_rows",
  "excel.range.merge",
  "excel.range.unmerge",
  "excel.batch.apply",
  "excel.batch.validate",
  "excel.batch.dry_run",
  "excel.template.detect_templates",
  "excel.template.register",
  "excel.template.unregister",
  "excel.template.get",
  "excel.template.list",
  "excel.template.infer_regions",
  "excel.template.create_sheet_from_template",
  "excel.template.clear_data_regions",
  "excel.template.fill_regions",
  "excel.template.validate_sheet_against_template",
  "excel.template.repair_sheet_from_template",
  "excel.style.get_fingerprint",
  "excel.style.compare_fingerprint",
  "excel.style.copy_from_template",
  "excel.style.apply_style",
  "excel.style.validate_consistency",
  "excel.style.repair_consistency",
  "excel.style.get_theme",
  "excel.style.apply_theme",
  "excel.style.copy_column_widths",
  "excel.style.copy_row_heights",
  "excel.style.copy_borders",
  "excel.style.copy_fills",
  "excel.style.copy_fonts",
  "excel.style.copy_alignment",
  "excel.style.copy_number_formats",
  "excel.style.copy_conditional_formatting",
  "excel.style.copy_data_validation",
  "excel.style.copy_freeze_panes",
  "excel.style.copy_print_settings",
  "excel.style.copy_page_layout",
  "excel.style.copy_hidden_rows_columns",
  "excel.formula.read_patterns",
  "excel.formula.copy_patterns",
  "excel.formula.fill_down",
  "excel.formula.fill_right",
  "excel.formula.validate",
  "excel.formula.validate_against_template",
  "excel.formula.repair_patterns",
  "excel.formula.find_errors",
  "excel.formula.find_circular_references",
  "excel.formula.get_dependency_graph",
  "excel.formula.trace_precedents",
  "excel.formula.trace_dependents",
  "excel.formula.convert_to_values",
  "excel.formula.recalculate",
  "excel.formula.explain",
  "excel.table.list",
  "excel.table.get_info",
  "excel.table.read",
  "excel.table.create",
  "excel.table.resize",
  "excel.table.append_rows",
  "excel.table.update_rows",
  "excel.table.clear_data_keep_formulas",
  "excel.table.clear_filters",
  "excel.table.apply_filters",
  "excel.table.preserve_filters",
  "excel.table.sort",
  "excel.table.set_total_row",
  "excel.table.set_style",
  "excel.table.copy_structure",
  "excel.table.validate_against_template",
  "excel.filter.get_filters",
  "excel.filter.apply",
  "excel.filter.clear",
  "excel.filter.preserve_from_template",
  "excel.filter.validate",
  "excel.sort.apply",
  "excel.sort.clear",
  "excel.sort.preserve_from_template",
  "excel.pivot.list",
  "excel.pivot.get_info",
  "excel.pivot.create",
  "excel.pivot.refresh",
  "excel.pivot.refresh_all",
  "excel.pivot.update_source",
  "excel.pivot.copy_from_template",
  "excel.pivot.delete",
  "excel.pivot.validate_source",
  "excel.pivot.get_capability_matrix",
  "excel.pivot.get_fingerprint",
  "excel.pivot.compare_fingerprint",
  "excel.pivot.diff",
  "excel.pivot.repair_from_template",
  "excel.pivot.rebuild_with_source",
  "excel.chart.list",
  "excel.chart.get_info",
  "excel.chart.create",
  "excel.chart.update_data_source",
  "excel.chart.copy_from_template",
  "excel.chart.refresh",
  "excel.chart.delete",
  "excel.chart.validate_against_template",
  "excel.names.list",
  "excel.names.get",
  "excel.names.create",
  "excel.names.update",
  "excel.names.delete",
  "excel.region.detect",
  "excel.region.register",
  "excel.region.list",
  "excel.region.get",
  "excel.region.clear_values",
  "excel.region.write_values",
  "excel.region.fill",
  "excel.plan.create",
  "excel.plan.preview",
  "excel.plan.refresh_preview",
  "excel.plan.rebase",
  "excel.plan.apply",
  "excel.plan.rollback",
  "excel.task.create",
  "excel.task.claim",
  "excel.task.update",
  "excel.task.set_progress",
  "excel.task.add_blocker",
  "excel.task.resolve_blocker",
  "excel.task.evaluate_schedule",
  "excel.task.resume_ready",
  "excel.task.complete",
  "excel.task.fail",
  "excel.task.cancel",
  "excel.task.list",
  "excel.task.get",
  "excel.collab.get_status",
  "excel.collab.list_agents",
  "excel.collab.list_tasks",
  "excel.collab.list_locks",
  "excel.collab.list_transactions",
  "excel.collab.get_conflicts",
  "excel.collab.get_recent_events",
  "excel.lock.get_policy",
  "excel.lock.set_policy",
  "excel.lock.acquire",
  "excel.lock.renew",
  "excel.lock.release",
  "excel.conflict.get_guidance",
  "excel.conflict.explain",
  "excel.conflict.get_telemetry",
  "excel.conflict.clear_telemetry",
  "excel.transaction.get",
  "excel.transaction.list",
  "excel.transaction.preview_rollback",
  "excel.transaction.rollback",
  "excel.transaction.preview_rollback_chain",
  "excel.transaction.rollback_chain",
  "excel.diff.create",
  "excel.diff.summarize",
  "excel.diff.get_details",
  "excel.diff.export_json",
  "excel.diff.export_html",
  "excel.events.subscribe",
  "excel.events.unsubscribe",
  "excel.events.get_recent",
  "excel.events.clear",
  "excel.events.set_debounce",
  "excel.snapshot.create",
  "excel.snapshot.refresh",
  "excel.snapshot.get",
  "excel.snapshot.compare",
  "excel.snapshot.invalidate",
  "excel.snapshot.list",
  "excel.snapshot.delete",
  "excel.validate.workbook",
  "excel.validate.sheet",
  "excel.validate.template_consistency",
  "excel.validate.formulas",
  "excel.validate.styles",
  "excel.validate.tables",
  "excel.validate.filters",
  "excel.validate.print_layout",
  "excel.validate.no_broken_references",
  "excel.validate.no_formula_errors",
  "excel.validate.no_unintended_changes",
  "excel.repair.style_from_template",
  "excel.repair.formulas_from_template",
  "excel.repair.filters_from_template",
  "excel.repair.table_structure",
  "excel.repair.print_layout",
  "excel.repair.named_ranges",
  "excel.repair.formula_errors",
  "excel.repair.merged_cells",
  "excel.clean.detect_header_row",
  "excel.clean.normalize_headers",
  "excel.clean.trim_whitespace",
  "excel.clean.remove_duplicates",
  "excel.clean.parse_dates",
  "excel.clean.parse_numbers",
  "excel.clean.standardize_currency",
  "excel.clean.fill_missing_values",
  "excel.clean.split_column",
  "excel.clean.merge_columns",
  "excel.clean.detect_outliers",
  "excel.clean.fuzzy_match",
  "excel.permissions.get",
  "excel.permissions.set",
  "excel.permissions.require_confirmation",
  "excel.permissions.set_scope",
  "excel.permissions.allow_destructive_actions",
  "excel.permissions.allow_macro_execution",
  "excel.permissions.lock_regions",
  "excel.permissions.unlock_regions"
] as const;

export const ToolCatalog: ToolContract[] = TOOL_NAMES.map((name) => {
  const namespace = getNamespace(name);
  const mutatesWorkbook = isMutatingTool(name);
  const destructiveLevel = getDestructiveLevel(name, namespace);
  return {
    name,
    title: titleize(name),
    namespace,
    status: getToolStatus(name, namespace),
    mutatesWorkbook,
    destructiveLevel,
    requiresConfirmation: mutatesWorkbook,
    requiredCapabilities: getRequiredCapabilities(name, namespace),
    description: `Open Workbook ${name} tool.`
  };
});

export const InitialToolContracts = getExposedToolCatalog();

export function getExposedToolCatalog(options: ToolCatalogOptions = {}): ToolContract[] {
  return ToolCatalog.filter((tool) => {
    if (tool.status === "stable") {
      return true;
    }
    if (tool.status === "preview") {
      return options.includePreview === true;
    }
    if (tool.status === "planned") {
      return options.includePlanned === true;
    }
    return options.includeUnsupported === true;
  });
}

export function getToolCatalogSummary(options: ToolCatalogOptions = {}) {
  const exposed = getExposedToolCatalog(options);
  return {
    total: ToolCatalog.length,
    exposed: exposed.length,
    stable: ToolCatalog.filter((tool) => tool.status === "stable").length,
    preview: ToolCatalog.filter((tool) => tool.status === "preview").length,
    planned: ToolCatalog.filter((tool) => tool.status === "planned").length,
    unsupported: ToolCatalog.filter((tool) => tool.status === "unsupported").length,
    tools: ToolCatalog
  };
}

export function isToolExposed(name: string, options: ToolCatalogOptions = {}): boolean {
  return getExposedToolCatalog(options).some((tool) => tool.name === name);
}

function getNamespace(name: string): ToolNamespace {
  const namespace = name.split(".")[1];
  if (!namespace) {
    throw new Error(`Invalid tool name: ${name}`);
  }
  return namespace as ToolNamespace;
}

function getToolStatus(name: string, namespace: ToolNamespace): CatalogStatus {
  if (STABLE_TOOLS.has(name)) {
    return "stable";
  }
  if (PREVIEW_TOOLS.has(name)) {
    return "preview";
  }
  return "planned";
}

function isMutatingTool(name: string): boolean {
  return /\.(set_|write_|create|copy|rename|delete|move|hide|unhide|protect|unprotect|clear|apply|repair|fill|append|update|resize|sort|save|restore|close|insert|merge|unmerge|lock|unlock|convert|calculate|recalculate|register|unregister|commit|rollback|cancel|refresh|invalidate|parse|normalize|trim|remove|standardize|split|import|embed)/.test(
    name
  );
}

function getDestructiveLevel(name: string, namespace: ToolNamespace): DestructiveLevel {
  if (!isMutatingTool(name)) {
    return "none";
  }
  if (namespace === "workbook" || name.includes(".close") || name.includes(".save_as") || name.includes(".restore_backup")) {
    return "workbook";
  }
  if (namespace === "sheet" || namespace === "template" || name.includes("insert_") || name.includes("delete_")) {
    return "structure";
  }
  if (namespace === "style" || name.includes("format") || name.includes("theme")) {
    return "format";
  }
  return "values";
}

function getRequiredCapabilities(name: string, namespace: ToolNamespace): string[] {
  if (namespace === "runtime") {
    return ["runtime.session"];
  }
  if (namespace === "workbook") {
    return ["workbook.context"];
  }
  if (namespace === "sheet") {
    return ["worksheet.basic"];
  }
  if (namespace === "range") {
    return [name.includes(".read_") || name.endsWith(".search") || name.endsWith(".find_errors") ? "range.read" : "range.write"];
  }
  return [`${namespace}.basic`];
}

function titleize(name: string): string {
  return name
    .replace(/^excel\./, "")
    .replace(/\./g, " ")
    .replace(/_/g, " ")
    .replace(/\b\w/g, (char) => char.toUpperCase());
}
