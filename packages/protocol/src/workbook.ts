import type {
  BackupId,
  OperationId,
  PlanId,
  SnapshotId,
  TemplateId,
  WorkbookId,
  WorksheetId
} from "./ids.js";

export type CellPrimitive = string | number | boolean | null;
export type CellValue = CellPrimitive | Date;
export type CellMatrix<T = CellValue> = T[][];

export interface WorkbookRef {
  workbookId: WorkbookId;
  name: string;
  path?: string;
  platform: "mac" | "windows" | "web" | "unknown";
}

export interface WorksheetRef {
  workbookId: WorkbookId;
  worksheetId?: WorksheetId;
  name: string;
}

export interface A1Range {
  workbookId: WorkbookId;
  sheetName: string;
  address: string;
}

export interface RangeFingerprint {
  range: A1Range;
  hash: string;
  cellCount: number;
  capturedAt: string;
}

export interface WorkbookFingerprint {
  workbookId: WorkbookId;
  workbookHash: string;
  structureHash: string;
  capturedAt: string;
}

export interface SnapshotRef {
  snapshotId: SnapshotId;
  workbookId: WorkbookId;
  createdAt: string;
  reason: string;
  affectedRanges: A1Range[];
}

export interface BackupRef {
  backupId: BackupId;
  workbookId: WorkbookId;
  kind: "region" | "sheet" | "workbook-copy";
  createdAt: string;
  reason: string;
  operationId?: OperationId;
}

export interface TemplateRef {
  templateId: TemplateId;
  name: string;
  scope: "workbook" | "local";
  version: number;
}

export interface PlanRef {
  planId: PlanId;
  workbookId: WorkbookId;
  status: "draft" | "previewed" | "applying" | "applied" | "rolled_back" | "failed" | "cancelled";
}
