import {
  makeId,
  type AgentId,
  type AgentOperationId,
  type AgentRunOutput,
  type ExcelOperation,
  type FormulaCopyPatternsRequest,
  type FormulaFillRequest,
  type FormulaPatternRequest,
  type NameCreateRequest,
  type NameSelector,
  type NameUpdateRequest,
  type RegionRegisterRequest,
  type RegionSelector,
  type StyleCopyRequest,
  type WorkbookBackupRetentionRequest,
  type WorkbookCreateFileBackupRequest,
  type WorkbookLocalConfigImportRequest,
  type WorkbookRestoreFileBackupRequest,
  type TableApplyFiltersRequest,
  type TableAppendRowsRequest,
  type TableCopyStructureRequest,
  type TableCreateRequest,
  type TableReorderColumnsRequest,
  type TableResizeRequest,
  type TableSelector,
  type TableSetStyleRequest,
  type TableSetTotalRowRequest,
  type TableSortRequest,
  type TableUpdateRowsRequest,
  type AddinTemplateRepairRequest,
  type TemplateCaptureRequest,
  type A1Range,
  type BackupId,
  type SnapshotId,
  type TemplateId,
  type WorkbookId
} from "@components-kit/open-workbook-protocol";
import type { AgentOperationRisk } from "./agent-action-policy.js";

export type PendingAgentAction =
  | { kind: "batch"; operations: ExcelOperation[] }
  | { kind: "table.append_rows"; request: TableAppendRowsRequest }
  | { kind: "table.update_rows"; request: TableUpdateRowsRequest }
  | { kind: "table.create"; request: TableCreateRequest }
  | { kind: "table.resize"; request: TableResizeRequest }
  | { kind: "table.reorder_columns"; request: TableReorderColumnsRequest }
  | { kind: "table.clear_data_keep_formulas"; request: TableSelector }
  | { kind: "table.clear_filters"; request: TableSelector }
  | { kind: "table.apply_filters"; request: TableApplyFiltersRequest }
  | { kind: "table.sort"; request: TableSortRequest }
  | { kind: "table.set_total_row"; request: TableSetTotalRowRequest }
  | { kind: "table.set_style"; request: TableSetStyleRequest }
  | { kind: "table.copy_structure"; request: TableCopyStructureRequest }
  | { kind: "template.register"; request: TemplateCaptureRequest }
  | { kind: "template.unregister"; templateId: TemplateId }
  | { kind: "template.repair_sheet"; request: { workbookId: WorkbookId; templateId: TemplateId; targetSheetName: string; repair?: AddinTemplateRepairRequest["repair"] } }
  | { kind: "style.copy_dimensions"; request: StyleCopyRequest }
  | { kind: "style.repair_consistency"; request: { workbookId: WorkbookId; templateId: TemplateId; targetSheetName: string } }
  | { kind: "clean.transform"; action: AgentCleanMutationAction; request: AgentCleanRequest }
  | { kind: "workbook.snapshot"; request: { workbookId: WorkbookId; reason?: string; ranges?: A1Range[] } }
  | { kind: "workbook.create_backup"; request: { workbookId: WorkbookId; reason?: string; ranges?: A1Range[] } }
  | { kind: "snapshot.refresh"; snapshotId: SnapshotId; reason?: string }
  | { kind: "snapshot.invalidate"; snapshotId: SnapshotId }
  | { kind: "snapshot.delete"; snapshotId: SnapshotId }
  | { kind: "backup.create_file"; request: WorkbookCreateFileBackupRequest }
  | { kind: "backup.restore_file"; request: WorkbookRestoreFileBackupRequest }
  | { kind: "backup.prune"; request: WorkbookBackupRetentionRequest }
  | { kind: "backup.pin"; backupId: BackupId }
  | { kind: "backup.unpin"; backupId: BackupId }
  | { kind: "backup.delete"; backupId: BackupId }
  | { kind: "workbook.restore_backup"; backupId: BackupId }
  | { kind: "workbook.import_local_config"; request: WorkbookLocalConfigImportRequest }
  | { kind: "workbook.embed_local_config"; workbookId: WorkbookId; includePermissions?: boolean }
  | { kind: "workbook.import_embedded_local_config"; request: { workbookId: WorkbookId; includeTemplates?: boolean; includeRegions?: boolean; includePermissions?: boolean; overwrite?: boolean } }
  | { kind: "workbook.close"; workbookId: WorkbookId; closeBehavior?: "Save" | "SkipSave" }
  | { kind: "formula.copy_patterns"; request: FormulaCopyPatternsRequest }
  | { kind: "formula.fill_pattern"; request: FormulaFillRequest }
  | { kind: "formula.repair_patterns"; request: { workbookId: WorkbookId; templateId: TemplateId; targetSheetName: string } }
  | { kind: "formula.convert_to_values"; request: FormulaPatternRequest }
  | { kind: "names.create"; request: NameCreateRequest }
  | { kind: "names.update"; request: NameUpdateRequest }
  | { kind: "names.delete"; request: NameSelector }
  | { kind: "region.register"; request: RegionRegisterRequest }
  | { kind: "region.clear_values"; request: RegionSelector }
  | { kind: "region.write_values"; request: RegionSelector & { values: unknown[][] } }
  | { kind: "region.fill"; request: RegionSelector & { values: unknown[][]; clearFirst?: boolean } };

export type AgentCleanMutationAction =
  | "normalize_headers"
  | "trim_whitespace"
  | "remove_duplicates"
  | "parse_dates"
  | "parse_numbers"
  | "standardize_currency"
  | "fill_missing_values"
  | "split_column"
  | "merge_columns";

export interface AgentCleanRequest {
  workbookId: WorkbookId;
  sheetName: string;
  address: string;
  headerRowIndex?: number;
  hasHeader?: boolean;
  keyColumns?: number[];
  strategy?: "value" | "zero" | "previous" | "next";
  value?: unknown;
  columnIndex?: number;
  columnIndexes?: number[];
  delimiter?: string;
  separator?: string;
  targetAddress?: string;
}

export interface PendingAgentOperation {
  operationId: AgentOperationId | string;
  confirmationToken: string;
  workbookContextId: string;
  workbookId: WorkbookId;
  action: PendingAgentAction;
  changes: NonNullable<AgentRunOutput["changes"]>;
  createdAt: number;
  summary: string;
  risk?: AgentOperationRisk;
  agentId?: AgentId | string;
  agentName?: string;
  sourceFingerprintHash?: string;
  sourceTargetFingerprintHash?: string;
}

export class AgentOperationStore {
  private readonly pending = new Map<string, PendingAgentOperation>();

  create(input: Omit<PendingAgentOperation, "operationId" | "confirmationToken" | "createdAt">): PendingAgentOperation {
    const operation: PendingAgentOperation = {
      ...input,
      operationId: makeId<AgentOperationId>("agentop"),
      confirmationToken: makeId("confirm"),
      createdAt: Date.now()
    };
    this.pending.set(operation.operationId, operation);
    return operation;
  }

  get(operationId: string): PendingAgentOperation | undefined {
    return this.pending.get(operationId);
  }

  delete(operationId: string): void {
    this.pending.delete(operationId);
  }
}
