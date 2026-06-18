import {
  getInternalCapabilityCatalog,
  getInternalCapabilityCatalogSummary,
  type ToolCatalogOptions,
  type ToolContract
} from "@components-kit/open-workbook-protocol";
import { AGENT_ACTION_HANDLERS } from "./agent-action-handlers.js";

export type ExcelCapabilityDefinition = ToolContract;
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
  | "diff"
  | "events"
  | "snapshot"
  | "compact_resource"
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
  { group: "diff", label: "Diff", description: "Workbook diff creation, summary, and compact diff retrieval operations.", prefixes: ["excel.diff."] },
  { group: "events", label: "Events", description: "Runtime event subscription, recent-event, clear, and debounce operations.", prefixes: ["excel.events."] },
  { group: "snapshot", label: "Snapshot", description: "Workbook snapshot creation, refresh, comparison, invalidation, list, and delete operations.", prefixes: ["excel.snapshot."] },
  { group: "compact_resource", label: "Compact Resources", description: "Stored compact payload, cache, garbage collection, and context statistics operations.", prefixes: ["excel.compact."] },
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

export function listExcelCapabilities(options: ToolCatalogOptions = {}): ExcelCapabilityDefinition[] {
  return getInternalCapabilityCatalog(options);
}

export function getExcelCapability(name: string, options: ToolCatalogOptions = {}): ExcelCapabilityDefinition | undefined {
  return listExcelCapabilities(options).find((capability) => capability.name === name);
}

export function getExcelCapabilitySummary(options: ToolCatalogOptions = {}) {
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
  if (AGENT_ACTION_CAPABILITIES.has(name)) {
    return "agent_action_handler";
  }
  return "internal_capability";
}

export function listExcelCapabilitiesByGroup(group: ExcelCapabilityGroup, options: ToolCatalogOptions = {}): ExcelCapabilityDefinition[] {
  return listExcelCapabilities(options).filter((capability) => getExcelCapabilityGroup(capability.name) === group);
}

export function listExcelCapabilityGroups(options: ToolCatalogOptions = {}): ExcelCapabilityGroupSummary[] {
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
