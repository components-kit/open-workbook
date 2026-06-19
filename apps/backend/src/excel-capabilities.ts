import {
  getInternalCapabilityCatalog,
  getInternalCapabilityCatalogSummary,
  type CapabilityCatalogOptions,
  type CapabilityContract
} from "@components-kit/open-workbook-protocol";
import { AGENT_ACTION_HANDLERS } from "./agent-action-handlers.js";

export type ExcelCapabilityDefinition = CapabilityContract;
export type ExcelCapabilityGroup =
  | "agent"
  | "runtime"
  | "workbook"
  | "backup"
  | "worksheet"
  | "range"
  | "lookup"
  | "batch"
  | "workflow"
  | "plan"
  | "job"
  | "task"
  | "collaboration"
  | "lock"
  | "conflict"
  | "transaction"
  | "events"
  | "snapshot"
  | "template"
  | "formatting"
  | "formula"
  | "table"
  | "pivot"
  | "chart"
  | "names"
  | "region"
  | "validation"
  | "repair"
  | "cleaning"
  | "permissions";

export type ExcelCapabilityAgentStatus = "agent_entrypoint" | "agent_action_handler" | "internal_capability";
export type ExcelCapabilityPlanningStatus = "covered" | "needs_unit_contract" | "future_orchestration_candidate" | "host_limited" | "defer";

export interface ExcelCapabilityGroupDefinition {
  group: ExcelCapabilityGroup;
  label: string;
  description: string;
  prefixes: string[];
}

export interface ExcelCapabilityGroupSummary {
  group: ExcelCapabilityGroup;
  label: string;
  description: string;
  total: number;
  readOnly: number;
  mutating: number;
  agentEntrypoint: number;
  agentActionHandlers: number;
  internalOnly: number;
  capabilities: ExcelCapabilityDefinition[];
}

export interface ExcelCapabilityCoverageEntry {
  capability: ExcelCapabilityDefinition;
  group: ExcelCapabilityGroup;
  agentStatus: ExcelCapabilityAgentStatus;
  planningStatus: ExcelCapabilityPlanningStatus;
}

export interface ExcelCapabilityCoverageSummary {
  total: number;
  byPlanningStatus: Record<ExcelCapabilityPlanningStatus, number>;
  byGroup: Array<{
    group: ExcelCapabilityGroup;
    label: string;
    total: number;
    byPlanningStatus: Record<ExcelCapabilityPlanningStatus, number>;
  }>;
  entries: ExcelCapabilityCoverageEntry[];
}

export const EXCEL_CAPABILITY_GROUPS: ExcelCapabilityGroupDefinition[] = [
  { group: "agent", label: "Agent", description: "Public agent entrypoint and agent-run control surface.", prefixes: ["excel.agent."] },
  { group: "runtime", label: "Runtime", description: "Backend/add-in runtime discovery, session, and active context operations.", prefixes: ["excel.runtime."] },
  { group: "workbook", label: "Workbook", description: "Workbook-level metadata, persistence, calculation, export, and local config operations.", prefixes: ["excel.workbook."] },
  { group: "backup", label: "Backup", description: "Local workbook backup lifecycle, retention, pinning, verification, and restore operations.", prefixes: ["excel.backup."] },
  { group: "worksheet", label: "Worksheet", description: "Worksheet metadata and structural sheet operations.", prefixes: ["excel.sheet."] },
  { group: "range", label: "Range", description: "Cell range reads, writes, formatting, movement, and shape operations.", prefixes: ["excel.range."] },
  { group: "lookup", label: "Lookup", description: "Workbook search, target resolution, header discovery, and candidate inspection operations.", prefixes: ["excel.lookup."] },
  { group: "batch", label: "Batch", description: "Compiled operation preflight, validation, dry-run, submit, and apply operations.", prefixes: ["excel.batch."] },
  { group: "workflow", label: "Workflow", description: "Combined workbook workflows that compose multiple internal capabilities.", prefixes: ["excel.workflow."] },
  { group: "plan", label: "Plan", description: "Plan preview, rebase, apply, and rollback operations.", prefixes: ["excel.plan."] },
  { group: "job", label: "Job", description: "Queued job list, status, wait, and cancellation operations.", prefixes: ["excel.job."] },
  { group: "task", label: "Task", description: "Collaborative task claim, progress, blocker, schedule, completion, and failure operations.", prefixes: ["excel.task."] },
  { group: "collaboration", label: "Collaboration", description: "Agent, task, lock, transaction, conflict, and event collaboration views.", prefixes: ["excel.collab."] },
  { group: "lock", label: "Lock", description: "Manual and policy-driven workbook lock operations.", prefixes: ["excel.lock."] },
  { group: "conflict", label: "Conflict", description: "Conflict guidance, explanation, and telemetry operations.", prefixes: ["excel.conflict."] },
  { group: "transaction", label: "Transaction", description: "Transaction status, wait, cancellation, rollback preview, and rollback operations.", prefixes: ["excel.transaction."] },
  { group: "events", label: "Events", description: "Runtime event subscription, recent-event, clear, and debounce operations.", prefixes: ["excel.events."] },
  { group: "snapshot", label: "Snapshot", description: "Workbook snapshot creation, refresh, comparison, invalidation, list, and delete operations.", prefixes: ["excel.snapshot."] },
  { group: "template", label: "Template", description: "Template detection, registration, region inference, fill, validation, and repair operations.", prefixes: ["excel.template."] },
  { group: "formatting", label: "Formatting", description: "Style fingerprints, themes, template style copy, consistency, and formatting repair operations.", prefixes: ["excel.style."] },
  { group: "formula", label: "Formula", description: "Formula pattern, validation, repair, dependency, trace, recalculation, and explanation operations.", prefixes: ["excel.formula."] },
  { group: "table", label: "Table", description: "Excel table schema, compact read, structure, row, filter, sort, total-row, and style operations.", prefixes: ["excel.table."] },
  { group: "pivot", label: "PivotTables", description: "PivotTable creation, refresh, source, template copy, validation, fingerprint, diff, repair, and rebuild operations.", prefixes: ["excel.pivot."] },
  { group: "chart", label: "Charts", description: "Chart list, metadata, create, source update, template copy, refresh, delete, and validation operations.", prefixes: ["excel.chart."] },
  { group: "names", label: "Names", description: "Named range list, get, create, update, and delete operations.", prefixes: ["excel.names."] },
  { group: "region", label: "Regions", description: "Detected and registered region list, get, clear, write, and fill operations.", prefixes: ["excel.region."] },
  { group: "validation", label: "Validation", description: "Workbook, sheet, template, formula, style, table, filter, layout, reference, error, and unintended-change validation.", prefixes: ["excel.validate."] },
  { group: "repair", label: "Repair", description: "Template-backed style, formula, filter, table, print layout, named range, formula error, and merged-cell repair operations.", prefixes: ["excel.repair."] },
  { group: "cleaning", label: "Cleaning", description: "Header, whitespace, duplicate, date, number, currency, missing value, split, merge, outlier, and fuzzy-match cleanup operations.", prefixes: ["excel.clean."] },
  { group: "permissions", label: "Permissions", description: "Permission, confirmation, scope, destructive-action, macro, and region lock policy operations.", prefixes: ["excel.permissions."] }
];

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
const PLANNING_STATUSES: ExcelCapabilityPlanningStatus[] = ["covered", "needs_unit_contract", "future_orchestration_candidate", "host_limited", "defer"];
const UNIT_CONTRACT_GROUPS = new Set<ExcelCapabilityGroup>([
  "runtime",
  "lookup",
  "batch",
  "plan",
  "job",
  "task",
  "collaboration",
  "lock",
  "conflict",
  "transaction",
  "events",
  "pivot",
  "chart",
  "permissions"
]);

const UNIT_CONTRACT_CAPABILITY_NAMES = new Set([
  "excel.workbook.save_as",
  "excel.workbook.export_copy"
]);

const CONTRACT_TESTED_CAPABILITY_NAMES = new Set([
  "excel.runtime.connect_addin",
  "excel.runtime.disconnect_addin",
  "excel.runtime.ping_addin",
  "excel.runtime.get_capabilities",
  "excel.runtime.set_active_workbook",
  "excel.runtime.set_active_sheet",
  "excel.workbook.save_as",
  "excel.workbook.export_copy",
  "excel.batch.apply",
  "excel.batch.submit",
  "excel.batch.submit_chunked",
  "excel.batch.preflight",
  "excel.batch.validate",
  "excel.batch.dry_run",
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
  "excel.transaction.wait",
  "excel.transaction.cancel",
  "excel.transaction.preview_rollback",
  "excel.transaction.rollback",
  "excel.transaction.preview_rollback_chain",
  "excel.transaction.rollback_chain",
  "excel.job.list",
  "excel.job.get",
  "excel.job.wait",
  "excel.job.cancel",
  "excel.events.subscribe",
  "excel.events.unsubscribe",
  "excel.events.get_recent",
  "excel.events.clear",
  "excel.events.set_debounce",
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
  "excel.permissions.get",
  "excel.permissions.set",
  "excel.permissions.require_confirmation",
  "excel.permissions.set_scope",
  "excel.permissions.allow_destructive_actions",
  "excel.permissions.allow_macro_execution",
  "excel.permissions.lock_regions",
  "excel.permissions.unlock_regions"
]);

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
  const matches = EXCEL_CAPABILITY_GROUPS.filter((definition) => definition.prefixes.some((prefix) => name.startsWith(prefix)));
  const match = matches[0];
  return matches.length === 1 && match ? match.group : undefined;
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
  const agentStatus = getExcelCapabilityAgentStatus(capability.name);
  const group = getExcelCapabilityGroup(capability.name);
  if (agentStatus !== "internal_capability") {
    return "covered";
  }
  if (capability.status === "planned" || capability.status === "unsupported") {
    return "defer";
  }
  if (CONTRACT_TESTED_CAPABILITY_NAMES.has(capability.name)) {
    return "covered";
  }
  if ((group !== undefined && UNIT_CONTRACT_GROUPS.has(group)) || UNIT_CONTRACT_CAPABILITY_NAMES.has(capability.name)) {
    return "needs_unit_contract";
  }
  return "future_orchestration_candidate";
}

export function listExcelCapabilitiesByGroup(group: ExcelCapabilityGroup, options: CapabilityCatalogOptions = {}): ExcelCapabilityDefinition[] {
  return listExcelCapabilities(options).filter((capability) => getExcelCapabilityGroup(capability.name) === group);
}

export function listExcelCapabilityCoverage(options: CapabilityCatalogOptions = {}): ExcelCapabilityCoverageEntry[] {
  return listExcelCapabilities(options).map((capability) => {
    const group = getExcelCapabilityGroup(capability.name);
    if (!group) {
      throw new Error(`Capability is not assigned to one coverage group: ${capability.name}`);
    }
    return {
      capability,
      group,
      agentStatus: getExcelCapabilityAgentStatus(capability.name),
      planningStatus: getExcelCapabilityPlanningStatus(capability)
    };
  });
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

function countByPlanningStatus(entries: ExcelCapabilityCoverageEntry[]): Record<ExcelCapabilityPlanningStatus, number> {
  return Object.fromEntries(PLANNING_STATUSES.map((status) => [status, entries.filter((entry) => entry.planningStatus === status).length])) as Record<ExcelCapabilityPlanningStatus, number>;
}
