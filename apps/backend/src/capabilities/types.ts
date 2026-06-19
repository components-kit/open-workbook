import type { CapabilityCatalogOptions, CapabilityContract } from "@components-kit/open-workbook-protocol";
import type { AgentActionHandlerId } from "../agent-action-handlers.js";

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
export type BackendCapabilityCoverageStatus = ExcelCapabilityPlanningStatus;

export interface ExcelCapabilityGroupDefinition {
  group: ExcelCapabilityGroup;
  label: string;
  description: string;
  prefixes: string[];
  implementationOwner: string;
  unitTestFile: string;
}

export interface BackendCapabilityDefinition {
  capability: ExcelCapabilityDefinition;
  name: string;
  group: ExcelCapabilityGroup;
  implementationOwner: string;
  runtimeMethod?: string;
  agentHandlerIds: AgentActionHandlerId[];
  hostMethods: string[];
  operationKinds: string[];
  statefulManagers: string[];
  coverageStatus: BackendCapabilityCoverageStatus;
  unitTestFile: string;
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

export type { CapabilityCatalogOptions };
