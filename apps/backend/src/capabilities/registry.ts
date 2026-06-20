import {
  getInternalCapabilityCatalog,
  getInternalCapabilityCatalogSummary,
  type CapabilityCatalogOptions
} from "@components-kit/open-workbook-protocol";
import { AGENT_ACTION_HANDLERS } from "../agent-action-handlers.js";
import { BACKEND_CAPABILITY_DOMAINS } from "./domains/metadata.js";
import type {
  BackendCapabilityDefinition,
  ExcelCapabilityAgentStatus,
  ExcelCapabilityCoverageEntry,
  ExcelCapabilityCoverageSummary,
  ExcelCapabilityDefinition,
  ExcelCapabilityGroup,
  ExcelCapabilityGroupSummary,
  ExcelCapabilityPlanningStatus
} from "./types.js";

export const EXCEL_CAPABILITY_GROUPS = BACKEND_CAPABILITY_DOMAINS;

const PLANNING_STATUSES: ExcelCapabilityPlanningStatus[] = ["covered", "needs_unit_contract", "future_orchestration_candidate", "host_limited", "defer"];
const AGENT_ACTION_CAPABILITIES = new Set(AGENT_ACTION_HANDLERS.map((handler) => handler.capabilityName));
const AGENT_ORCHESTRATED_CAPABILITIES = new Set([
  ...AGENT_ACTION_CAPABILITIES,
  "excel.runtime.get_status",
  "excel.runtime.get_active_context",
  "excel.runtime.get_selection",
  "excel.workbook.list_open_workbooks",
  "excel.workbook.get_workbook_info",
  "excel.workbook.get_workbook_map",
  "excel.workbook.get_summary",
  "excel.workbook.get_used_range_summary",
  "excel.workbook.snapshot",
  "excel.workbook.refresh_snapshot",
  "excel.workbook.get_snapshot",
  "excel.workbook.detect_external_changes",
  "excel.workbook.create_backup",
  "excel.workbook.restore_backup",
  "excel.workbook.export_local_config",
  "excel.workbook.import_local_config",
  "excel.workbook.embed_local_config",
  "excel.workbook.read_embedded_local_config",
  "excel.workbook.import_embedded_local_config",
  "excel.workbook.close",
  "excel.snapshot.create",
  "excel.snapshot.refresh",
  "excel.snapshot.list",
  "excel.snapshot.get_compact",
  "excel.snapshot.compare_compact",
  "excel.snapshot.invalidate",
  "excel.snapshot.delete",
  "excel.backup.list",
  "excel.backup.create_file",
  "excel.backup.get",
  "excel.backup.verify",
  "excel.backup.restore_file",
  "excel.backup.prune",
  "excel.backup.delete",
  "excel.backup.pin",
  "excel.backup.unpin",
  "excel.sheet.list",
  "excel.sheet.get_info",
  "excel.sheet.get_summary",
  "excel.sheet.get_used_range",
  "excel.sheet.copy_clean_data_regions",
  "excel.range.read_compact",
  "excel.range.get_summary",
  "excel.range.read_hyperlinks",
  "excel.range.read_comments",
  "excel.range.read_notes",
  "excel.range.read_merged_cells",
  "excel.range.read_data_validation",
  "excel.range.read_conditional_formatting",
  "excel.range.search",
  "excel.range.find_blank_cells",
  "excel.range.find_errors",
  "excel.range.write_styles_many",
  "excel.range.write_values_many",
  "excel.range.write_number_formats_many",
  "excel.range.clear_many",
  "excel.range.clear_formats_many",
  "excel.range.autofit_many",
  "excel.range.clear_values",
  "excel.range.insert_rows",
  "excel.range.delete_rows",
  "excel.range.insert_columns",
  "excel.range.delete_columns",
  "excel.range.merge",
  "excel.range.unmerge",
  "excel.workflow.prepare_session",
  "excel.workflow.create_formula_sheet",
  "excel.workflow.create_template_report",
  "excel.workflow.create_pivot_chart_summary",
  "excel.workflow.repair_formula_errors",
  "excel.workflow.preview_risky_edit",
  "excel.workflow.inspect_analyze",
  "excel.workflow.rollback_validate",
  "excel.lookup.search_workbook",
  "excel.lookup.resolve_range",
  "excel.lookup.inspect_match",
  "excel.table.list",
  "excel.table.get_info",
  "excel.table.get_schema",
  "excel.table.read_compact",
  "excel.table.create",
  "excel.table.resize",
  "excel.table.reorder_columns",
  "excel.table.apply_view",
  "excel.table.copy_structure",
  "excel.table.validate_against_template",
  "excel.template.detect_templates",
  "excel.template.register",
  "excel.template.unregister",
  "excel.template.get",
  "excel.template.list",
  "excel.template.infer_regions",
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
  "excel.repair.style_from_template",
  "excel.repair.formulas_from_template",
  "excel.repair.filters_from_template",
  "excel.repair.table_structure",
  "excel.repair.print_layout",
  "excel.repair.named_ranges",
  "excel.repair.formula_errors",
  "excel.repair.merged_cells",
  "excel.names.list",
  "excel.region.detect",
  "excel.region.register",
  "excel.region.list",
  "excel.region.get",
  "excel.region.clear_values",
  "excel.region.write_values",
  "excel.region.fill",
  "excel.validate.workbook",
  "excel.validate.compact",
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
  "excel.formula.read_patterns",
  "excel.formula.copy_patterns",
  "excel.formula.fill_down",
  "excel.formula.fill_right",
  "excel.formula.validate",
  "excel.formula.validate_against_template",
  "excel.formula.repair_patterns",
  "excel.formula.find_errors",
  "excel.formula.get_dependency_graph",
  "excel.formula.trace_precedents",
  "excel.formula.trace_dependents",
  "excel.formula.convert_to_values",
  "excel.formula.recalculate",
  "excel.formula.explain",
  "excel.names.get"
]);

export function listBackendCapabilityRegistry(options: CapabilityCatalogOptions = {}): BackendCapabilityDefinition[] {
  return listExcelCapabilities(options).map((capability) => {
    const domain = getBackendCapabilityDomain(capability.name);
    if (!domain) {
      throw new Error(`Capability is not assigned to one backend domain: ${capability.name}`);
    }
    const runtimeMethod = RUNTIME_METHOD_BY_CAPABILITY[capability.name];
    return {
      capability,
      name: capability.name,
      group: domain.group,
      implementationOwner: domain.implementationOwner,
      ...(runtimeMethod ? { runtimeMethod } : {}),
      agentHandlerIds: AGENT_ACTION_HANDLERS.filter((handler) => handler.capabilityName === capability.name).map((handler) => handler.id),
      hostMethods: hostMethodsForCapability(capability.name),
      operationKinds: operationKindsForCapability(capability.name),
      statefulManagers: statefulManagersForCapability(domain.group),
      coverageStatus: getExcelCapabilityPlanningStatus(capability),
      unitTestFile: domain.unitTestFile
    };
  });
}

export function getBackendCapability(name: string, options: CapabilityCatalogOptions = {}): BackendCapabilityDefinition | undefined {
  return listBackendCapabilityRegistry(options).find((entry) => entry.name === name);
}

export function listExcelCapabilities(options: CapabilityCatalogOptions = {}): ExcelCapabilityDefinition[] {
  return getInternalCapabilityCatalog(options);
}

export function getExcelCapability(name: string, options: CapabilityCatalogOptions = {}): ExcelCapabilityDefinition | undefined {
  return listExcelCapabilities(options).find((capability) => capability.name === name);
}

export function getExcelCapabilitySummary(options: CapabilityCatalogOptions = {}) {
  return getInternalCapabilityCatalogSummary(options);
}

export function getExcelCapabilityGroup(name: string): ExcelCapabilityGroup | undefined {
  return getBackendCapabilityDomain(name)?.group;
}

export function getExcelCapabilityAgentStatus(name: string): ExcelCapabilityAgentStatus {
  if (name === "excel.agent.run") {
    return "agent_entrypoint";
  }
  if (AGENT_ORCHESTRATED_CAPABILITIES.has(name)) {
    return "agent_action_handler";
  }
  return "internal_capability";
}

export function getExcelCapabilityPlanningStatus(capability: ExcelCapabilityDefinition): ExcelCapabilityPlanningStatus {
  if (getExcelCapabilityAgentStatus(capability.name) !== "internal_capability") {
    return "covered";
  }
  if (capability.status === "planned" || capability.status === "unsupported") {
    return "defer";
  }
  return "covered";
}

export function listExcelCapabilitiesByGroup(group: ExcelCapabilityGroup, options: CapabilityCatalogOptions = {}): ExcelCapabilityDefinition[] {
  return listExcelCapabilities(options).filter((capability) => getExcelCapabilityGroup(capability.name) === group);
}

export function listExcelCapabilityCoverage(options: CapabilityCatalogOptions = {}): ExcelCapabilityCoverageEntry[] {
  return listBackendCapabilityRegistry(options).map((entry) => ({
    capability: entry.capability,
    group: entry.group,
    agentStatus: getExcelCapabilityAgentStatus(entry.name),
    planningStatus: entry.coverageStatus
  }));
}

export function summarizeExcelCapabilityCoverage(options: CapabilityCatalogOptions = {}): ExcelCapabilityCoverageSummary {
  const entries = listExcelCapabilityCoverage(options);
  return {
    total: entries.length,
    byPlanningStatus: countByPlanningStatus(entries),
    byGroup: EXCEL_CAPABILITY_GROUPS.map((definition) => {
      const groupEntries = entries.filter((entry) => entry.group === definition.group);
      return {
        group: definition.group,
        label: definition.label,
        total: groupEntries.length,
        byPlanningStatus: countByPlanningStatus(groupEntries)
      };
    }),
    entries
  };
}

export function listExcelCapabilityGroups(options: CapabilityCatalogOptions = {}): ExcelCapabilityGroupSummary[] {
  return EXCEL_CAPABILITY_GROUPS.map((definition) => {
    const capabilities = listExcelCapabilitiesByGroup(definition.group, options);
    return {
      group: definition.group,
      label: definition.label,
      description: definition.description,
      total: capabilities.length,
      readOnly: capabilities.filter((capability) => !capability.mutatesWorkbook).length,
      mutating: capabilities.filter((capability) => capability.mutatesWorkbook).length,
      agentEntrypoint: capabilities.filter((capability) => getExcelCapabilityAgentStatus(capability.name) === "agent_entrypoint").length,
      agentActionHandlers: capabilities.filter((capability) => getExcelCapabilityAgentStatus(capability.name) === "agent_action_handler").length,
      internalOnly: capabilities.filter((capability) => getExcelCapabilityAgentStatus(capability.name) === "internal_capability").length,
      capabilities
    };
  });
}

function getBackendCapabilityDomain(name: string) {
  const matches = EXCEL_CAPABILITY_GROUPS.filter((definition) => definition.prefixes.some((prefix) => name.startsWith(prefix)));
  const match = matches[0];
  return matches.length === 1 && match ? match : undefined;
}

function countByPlanningStatus(entries: ExcelCapabilityCoverageEntry[]): Record<ExcelCapabilityPlanningStatus, number> {
  return Object.fromEntries(PLANNING_STATUSES.map((status) => [status, entries.filter((entry) => entry.planningStatus === status).length])) as Record<ExcelCapabilityPlanningStatus, number>;
}

function hostMethodsForCapability(name: string): string[] {
  return HOST_METHODS_BY_CAPABILITY[name] ?? [];
}

function operationKindsForCapability(name: string): string[] {
  return OPERATION_KINDS_BY_CAPABILITY[name] ?? [];
}

function statefulManagersForCapability(group: ExcelCapabilityGroup): string[] {
  const managers: Partial<Record<ExcelCapabilityGroup, string[]>> = {
    backup: ["BackupManager"],
    snapshot: ["SnapshotManager"],
    template: ["TemplateRegistry"],
    plan: ["PlanManager"],
    batch: ["BatchCompiler", "TransactionManager"],
    job: ["TransactionManager"],
    task: ["TaskRegistry"],
    collaboration: ["TaskRegistry", "LockManager", "TransactionManager"],
    lock: ["LockManager"],
    conflict: ["LockManager"],
    transaction: ["TransactionManager"],
    permissions: ["DefaultPermissionPolicy"]
  };
  return managers[group] ?? [];
}

const RUNTIME_METHOD_BY_CAPABILITY: Record<string, string> = {
  "excel.agent.run": "runAgent",
  "excel.runtime.get_status": "getStatus",
  "excel.runtime.get_capabilities": "getCapabilities",
  "excel.runtime.get_active_context": "getActiveContext",
  "excel.runtime.get_selection": "getSelection",
  "excel.runtime.connect_addin": "connectAddinInfo",
  "excel.runtime.disconnect_addin": "disconnectActiveAddin",
  "excel.runtime.ping_addin": "pingAddin",
  "excel.runtime.set_active_workbook": "setActiveWorkbook",
  "excel.runtime.set_active_sheet": "setActiveSheet",
  "excel.permissions.get": "getPermissions",
  "excel.permissions.set": "setPermissions",
  "excel.permissions.require_confirmation": "requireConfirmation",
  "excel.permissions.set_scope": "setPermissionScope",
  "excel.permissions.allow_destructive_actions": "allowDestructiveActions",
  "excel.permissions.allow_macro_execution": "allowMacroExecution",
  "excel.permissions.lock_regions": "lockRegions",
  "excel.permissions.unlock_regions": "unlockRegions"
};

const HOST_METHODS_BY_CAPABILITY: Record<string, string[]> = Object.fromEntries([
  ["excel.runtime.ping_addin", ["runtime.ping"]],
  ["excel.runtime.get_active_context", ["runtime.get_active_context"]],
  ["excel.runtime.get_selection", ["runtime.get_selection"]],
  ["excel.runtime.set_active_sheet", ["runtime.set_active_sheet"]],
  ["excel.workbook.get_workbook_info", ["workbook.get_info"]],
  ["excel.workbook.get_workbook_map", ["workbook.get_map"]],
  ["excel.workbook.get_summary", ["workbook.get_map"]],
  ["excel.workbook.get_used_range_summary", ["workbook.get_map"]],
  ["excel.workbook.snapshot", ["workbook.snapshot_ranges"]],
  ["excel.workbook.refresh_snapshot", ["workbook.snapshot_ranges"]],
  ["excel.workbook.get_snapshot", ["workbook.snapshot_ranges"]],
  ["excel.workbook.detect_external_changes", ["workbook.snapshot_ranges"]],
  ["excel.workbook.calculate", ["workbook.calculate", "operation.execute_batch"]],
  ["excel.workbook.save", ["workbook.save", "operation.execute_batch"]],
  ["excel.workbook.export_copy", ["workbook.snapshot_ranges", "workbook.get_file"]],
  ["excel.workbook.create_backup", ["workbook.snapshot_ranges"]],
  ["excel.workbook.restore_backup", ["operation.execute_batch"]],
  ["excel.workbook.embed_local_config", ["workbook.embed_local_config"]],
  ["excel.workbook.read_embedded_local_config", ["workbook.read_embedded_local_config"]],
  ["excel.workbook.import_embedded_local_config", ["workbook.embed_local_config", "workbook.read_embedded_local_config"]],
  ["excel.workbook.close", ["workbook.close"]],
  ["excel.sheet.create", ["operation.execute_batch"]],
  ["excel.sheet.copy", ["operation.execute_batch"]],
  ["excel.sheet.copy_clean_data_regions", ["operation.execute_batch"]],
  ["excel.sheet.rename", ["operation.execute_batch"]],
  ["excel.sheet.delete", ["operation.execute_batch"]],
  ["excel.sheet.hide", ["operation.execute_batch"]],
  ["excel.sheet.unhide", ["operation.execute_batch"]],
  ["excel.sheet.protect", ["operation.execute_batch"]],
  ["excel.sheet.unprotect", ["operation.execute_batch"]],
  ["excel.sheet.clear", ["operation.execute_batch"]],
  ["excel.sheet.set_tab_color", ["operation.execute_batch"]],
  ["excel.range.read_hyperlinks", ["range.read_hyperlinks"]],
  ["excel.range.read_comments", ["range.read_comments"]],
  ["excel.range.read_notes", ["range.read_notes"]],
  ["excel.range.read_merged_cells", ["range.read_merged_cells"]],
  ["excel.range.read_data_validation", ["range.read_data_validation"]],
  ["excel.range.read_conditional_formatting", ["range.read_conditional_formatting"]],
  ["excel.range.search", ["range.search"]],
  ["excel.range.find_blank_cells", ["range.find_blank_cells"]],
  ["excel.range.find_errors", ["range.find_errors"]],
  ["excel.range.write_values", ["operation.execute_batch"]],
  ["excel.range.write_values_many", ["operation.execute_batch"]],
  ["excel.range.write_formulas", ["operation.execute_batch"]],
  ["excel.range.write_number_formats", ["operation.execute_batch"]],
  ["excel.range.write_number_formats_many", ["operation.execute_batch"]],
  ["excel.range.write_styles", ["operation.execute_batch"]],
  ["excel.range.write_styles_many", ["operation.execute_batch"]],
  ["excel.range.clear", ["operation.execute_batch"]],
  ["excel.range.clear_many", ["operation.execute_batch"]],
  ["excel.range.clear_values", ["operation.execute_batch"]],
  ["excel.range.clear_formats", ["operation.execute_batch"]],
  ["excel.range.clear_formats_many", ["operation.execute_batch"]],
  ["excel.range.clear_values_keep_format", ["operation.execute_batch"]],
  ["excel.range.copy", ["operation.execute_batch"]],
  ["excel.range.move", ["operation.execute_batch"]],
  ["excel.range.insert_rows", ["operation.execute_batch"]],
  ["excel.range.delete_rows", ["operation.execute_batch"]],
  ["excel.range.insert_columns", ["operation.execute_batch"]],
  ["excel.range.delete_columns", ["operation.execute_batch"]],
  ["excel.range.autofit_columns", ["operation.execute_batch"]],
  ["excel.range.autofit_rows", ["operation.execute_batch"]],
  ["excel.range.autofit_many", ["operation.execute_batch"]],
  ["excel.range.merge", ["operation.execute_batch"]],
  ["excel.range.unmerge", ["operation.execute_batch"]],
  ["excel.batch.apply", ["operation.execute_batch"]],
  ["excel.batch.submit", ["operation.execute_batch"]],
  ["excel.batch.submit_chunked", ["operation.execute_batch"]],
  ["excel.template.detect_templates", ["template.capture"]],
  ["excel.template.register", ["template.capture"]],
  ["excel.template.create_sheet_from_template", ["operation.execute_batch"]],
  ["excel.template.validate_sheet_against_template", ["template.capture_sheet"]],
  ["excel.template.repair_sheet_from_template", ["template.repair"]],
  ["excel.style.get_fingerprint", ["style.capture_fingerprint", "template.capture_sheet"]],
  ["excel.style.compare_fingerprint", ["style.capture_fingerprint"]],
  ["excel.style.copy_from_template", ["style.copy_dimensions"]],
  ["excel.style.copy_column_widths", ["style.copy_dimensions"]],
  ["excel.style.copy_row_heights", ["style.copy_dimensions"]],
  ["excel.style.copy_borders", ["style.copy_dimensions"]],
  ["excel.style.copy_fills", ["style.copy_dimensions"]],
  ["excel.style.copy_fonts", ["style.copy_dimensions"]],
  ["excel.style.copy_alignment", ["style.copy_dimensions"]],
  ["excel.style.copy_number_formats", ["style.copy_dimensions"]],
  ["excel.style.copy_conditional_formatting", ["style.copy_dimensions"]],
  ["excel.style.copy_data_validation", ["style.copy_dimensions"]],
  ["excel.style.repair_consistency", ["template.repair"]],
  ["excel.formula.read_patterns", ["formula.read_patterns"]],
  ["excel.formula.copy_patterns", ["formula.copy_patterns"]],
  ["excel.formula.fill_down", ["formula.fill_pattern"]],
  ["excel.formula.fill_right", ["formula.fill_pattern"]],
  ["excel.formula.find_errors", ["range.find_errors"]],
  ["excel.formula.convert_to_values", ["formula.convert_to_values"]],
  ["excel.formula.recalculate", ["workbook.calculate"]],
  ["excel.table.list", ["table.list"]],
  ["excel.table.get_info", ["table.get_info"]],
  ["excel.table.get_schema", ["table.get_info"]],
  ["excel.table.read_compact", ["table.read"]],
  ["excel.table.create", ["table.create"]],
  ["excel.table.resize", ["table.resize"]],
  ["excel.table.reorder_columns", ["table.reorder_columns"]],
  ["excel.table.append_rows", ["table.append_rows"]],
  ["excel.table.update_rows", ["table.update_rows"]],
  ["excel.table.clear_data_keep_formulas", ["table.clear_data_keep_formulas"]],
  ["excel.table.clear_filters", ["table.clear_filters"]],
  ["excel.table.apply_filters", ["table.apply_filters"]],
  ["excel.table.sort", ["table.sort"]],
  ["excel.table.apply_view", ["table.apply_view"]],
  ["excel.table.set_total_row", ["table.set_total_row"]],
  ["excel.table.set_style", ["table.set_style"]],
  ["excel.table.copy_structure", ["table.copy_structure"]],
  ["excel.pivot.list", ["pivot.list"]],
  ["excel.pivot.get_info", ["pivot.get_info"]],
  ["excel.pivot.create", ["pivot.create"]],
  ["excel.pivot.refresh", ["pivot.refresh"]],
  ["excel.pivot.refresh_all", ["pivot.refresh_all"]],
  ["excel.pivot.copy_from_template", ["pivot.copy_from_template"]],
  ["excel.pivot.delete", ["pivot.delete"]],
  ["excel.pivot.repair_from_template", ["pivot.copy_from_template"]],
  ["excel.chart.list", ["chart.list"]],
  ["excel.chart.get_info", ["chart.get_info"]],
  ["excel.chart.create", ["chart.create"]],
  ["excel.chart.update_data_source", ["chart.update_data_source"]],
  ["excel.chart.copy_from_template", ["chart.copy_from_template"]],
  ["excel.chart.refresh", ["chart.refresh"]],
  ["excel.chart.delete", ["chart.delete"]],
  ["excel.names.list", ["names.list"]],
  ["excel.names.get", ["names.get"]],
  ["excel.names.create", ["names.create"]],
  ["excel.names.update", ["names.update"]],
  ["excel.names.delete", ["names.delete"]]
]);

const OPERATION_KINDS_BY_CAPABILITY: Record<string, string[]> = Object.fromEntries([
  ["excel.workbook.calculate", ["workbook.calculate"]],
  ["excel.workbook.save", ["workbook.save"]],
  ["excel.sheet.create", ["sheet.create"]],
  ["excel.sheet.copy", ["sheet.copy"]],
  ["excel.sheet.copy_clean_data_regions", ["sheet.copy_clean_data_regions"]],
  ["excel.sheet.rename", ["sheet.rename"]],
  ["excel.sheet.delete", ["sheet.delete"]],
  ["excel.sheet.hide", ["sheet.hide"]],
  ["excel.sheet.unhide", ["sheet.unhide"]],
  ["excel.sheet.protect", ["sheet.protect"]],
  ["excel.sheet.unprotect", ["sheet.unprotect"]],
  ["excel.sheet.clear", ["sheet.clear"]],
  ["excel.sheet.set_tab_color", ["sheet.set_tab_color"]],
  ["excel.range.write_values", ["range.write_values"]],
  ["excel.range.write_values_many", ["range.write_values_many"]],
  ["excel.range.write_formulas", ["range.write_formulas"]],
  ["excel.range.write_number_formats", ["range.write_number_formats"]],
  ["excel.range.write_number_formats_many", ["range.write_number_formats_many"]],
  ["excel.range.write_styles", ["range.write_styles"]],
  ["excel.range.write_styles_many", ["range.write_styles_many"]],
  ["excel.range.clear", ["range.clear"]],
  ["excel.range.clear_many", ["range.clear_many"]],
  ["excel.range.clear_values", ["range.clear_values"]],
  ["excel.range.clear_formats", ["range.clear_formats"]],
  ["excel.range.clear_formats_many", ["range.clear_formats_many"]],
  ["excel.range.clear_values_keep_format", ["range.clear_values_keep_format"]],
  ["excel.range.copy", ["range.copy"]],
  ["excel.range.move", ["range.move"]],
  ["excel.range.insert_rows", ["range.insert_rows"]],
  ["excel.range.delete_rows", ["range.delete_rows"]],
  ["excel.range.insert_columns", ["range.insert_columns"]],
  ["excel.range.delete_columns", ["range.delete_columns"]],
  ["excel.range.autofit_columns", ["range.autofit_columns"]],
  ["excel.range.autofit_rows", ["range.autofit_rows"]],
  ["excel.range.autofit_many", ["range.autofit_many"]],
  ["excel.range.merge", ["range.merge"]],
  ["excel.range.unmerge", ["range.unmerge"]],
  ["excel.template.create_sheet_from_template", ["template.create_sheet_from_template"]]
]);
