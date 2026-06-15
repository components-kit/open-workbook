import type { AgentId, BackupId, OperationId, PlanId, SnapshotId, TaskId, TemplateId, TransactionId, WorkbookId } from "./ids.js";
import type { A1Range, CellMatrix, RangeFingerprint, WorkbookFingerprint } from "./workbook.js";
import type { ExcelRuntimeError } from "./errors.js";

export type DestructiveLevel = "none" | "values" | "format" | "structure" | "workbook";

export interface OperationTelemetry {
  durationMs?: number;
  syncCount?: number;
  payloadBytes?: number;
  cellsRead?: number;
  cellsWritten?: number;
  rangeCount?: number;
  chunkCount?: number;
  engineName?: string;
  engineVersion?: string;
  warningCount?: number;
  requestedFacets?: string[];
  loadedFacets?: string[];
}

export interface OperationWarning {
  code: string;
  message: string;
  target?: A1Range;
  details?: Record<string, unknown>;
}

export interface OperationBase {
  operationId: OperationId;
  workbookId: WorkbookId;
  destructiveLevel: DestructiveLevel;
  reason: string;
}

export interface ReadFullOperation extends OperationBase {
  kind: "range.read_full";
  target: A1Range;
  facets?: Array<"values" | "formulas" | "numberFormat" | "text" | "style">;
  includeStyles?: boolean;
  includeFormulas?: boolean;
  includeValidation?: boolean;
  includeComments?: boolean;
}

export interface WriteValuesOperation extends OperationBase {
  kind: "range.write_values";
  target: A1Range;
  values: CellMatrix;
  preserveFormats: true;
}

export interface WriteFormulasOperation extends OperationBase {
  kind: "range.write_formulas";
  target: A1Range;
  formulas: CellMatrix<string | null>;
  preserveFormats: true;
}

export interface ClearValuesKeepFormatOperation extends OperationBase {
  kind: "range.clear_values_keep_format";
  target: A1Range;
}

export interface WriteNumberFormatsOperation extends OperationBase {
  kind: "range.write_number_formats";
  target: A1Range;
  numberFormat: string[][];
  preserveValues: true;
}

export interface WriteStylesOperation extends OperationBase {
  kind: "range.write_styles";
  target: A1Range;
  style: NonNullable<RangeSnapshot["style"]>;
  preserveValues: true;
}

export interface WriteHyperlinksOperation extends OperationBase {
  kind: "range.write_hyperlinks";
  target: A1Range;
  hyperlinks: Array<Array<{ text?: string; address: string; screenTip?: string } | null>>;
}

export interface WriteCommentsOperation extends OperationBase {
  kind: "range.write_comments";
  target: A1Range;
  comments: CellMatrix<string | null>;
}

export interface ClearRangeOperation extends OperationBase {
  kind: "range.clear";
  target: A1Range;
  applyTo?: "all" | "contents" | "formats" | "hyperlinks";
}

export interface ClearValuesOperation extends OperationBase {
  kind: "range.clear_values";
  target: A1Range;
}

export interface ClearFormatsOperation extends OperationBase {
  kind: "range.clear_formats";
  target: A1Range;
}

export interface CopyRangeOperation extends OperationBase {
  kind: "range.copy";
  source: A1Range;
  target: A1Range;
  copyType?: "all" | "values" | "formats" | "formulas";
}

export interface MoveRangeOperation extends OperationBase {
  kind: "range.move";
  source: A1Range;
  target: A1Range;
}

export interface InsertRowsOperation extends OperationBase {
  kind: "range.insert_rows";
  target: A1Range;
  shift?: "down" | "right";
}

export interface DeleteRowsOperation extends OperationBase {
  kind: "range.delete_rows";
  target: A1Range;
  shift?: "up" | "left";
}

export interface InsertColumnsOperation extends OperationBase {
  kind: "range.insert_columns";
  target: A1Range;
  shift?: "right" | "down";
}

export interface DeleteColumnsOperation extends OperationBase {
  kind: "range.delete_columns";
  target: A1Range;
  shift?: "left" | "up";
}

export interface AutofitColumnsOperation extends OperationBase {
  kind: "range.autofit_columns";
  target: A1Range;
}

export interface AutofitRowsOperation extends OperationBase {
  kind: "range.autofit_rows";
  target: A1Range;
}

export interface MergeRangeOperation extends OperationBase {
  kind: "range.merge";
  target: A1Range;
  across?: boolean;
}

export interface UnmergeRangeOperation extends OperationBase {
  kind: "range.unmerge";
  target: A1Range;
}

export interface RestoreRangeSnapshotOperation extends OperationBase {
  kind: "range.restore_snapshot";
  target: A1Range;
  snapshot: RangeSnapshot;
}

export interface CreateSheetOperation extends OperationBase {
  kind: "sheet.create";
  sheetName: string;
  position?: "beginning" | "end" | "before" | "after";
  relativeToSheetName?: string;
  activate?: boolean;
}

export interface CopySheetOperation extends OperationBase {
  kind: "sheet.copy";
  sourceSheetName: string;
  newSheetName: string;
  position?: "beginning" | "end" | "before" | "after";
  relativeToSheetName?: string;
  activate?: boolean;
}

export interface RenameSheetOperation extends OperationBase {
  kind: "sheet.rename";
  sheetName: string;
  newSheetName: string;
}

export interface DeleteSheetOperation extends OperationBase {
  kind: "sheet.delete";
  sheetName: string;
}

export interface MoveSheetOperation extends OperationBase {
  kind: "sheet.move";
  sheetName: string;
  position: "beginning" | "end" | "before" | "after";
  relativeToSheetName?: string;
}

export interface HideSheetOperation extends OperationBase {
  kind: "sheet.hide";
  sheetName: string;
}

export interface UnhideSheetOperation extends OperationBase {
  kind: "sheet.unhide";
  sheetName: string;
}

export interface ProtectSheetOperation extends OperationBase {
  kind: "sheet.protect";
  sheetName: string;
  password?: string;
}

export interface UnprotectSheetOperation extends OperationBase {
  kind: "sheet.unprotect";
  sheetName: string;
  password?: string;
}

export interface ClearSheetOperation extends OperationBase {
  kind: "sheet.clear";
  sheetName: string;
  applyTo?: "all" | "contents" | "formats";
}

export interface SetSheetTabColorOperation extends OperationBase {
  kind: "sheet.set_tab_color";
  sheetName: string;
  color: string;
}

export interface WorkbookCalculateOperation extends OperationBase {
  kind: "workbook.calculate";
  calculationType?: "full" | "recalculate";
}

export interface WorkbookSaveOperation extends OperationBase {
  kind: "workbook.save";
}

export interface CreateSheetFromTemplateOperation extends OperationBase {
  kind: "template.create_sheet_from_template";
  templateId: TemplateId;
  newSheetName: string;
  clearDataRegions: boolean;
}

export type ExcelOperation =
  | ReadFullOperation
  | WriteValuesOperation
  | WriteFormulasOperation
  | ClearValuesKeepFormatOperation
  | WriteNumberFormatsOperation
  | WriteStylesOperation
  | WriteHyperlinksOperation
  | WriteCommentsOperation
  | ClearRangeOperation
  | ClearValuesOperation
  | ClearFormatsOperation
  | CopyRangeOperation
  | MoveRangeOperation
  | InsertRowsOperation
  | DeleteRowsOperation
  | InsertColumnsOperation
  | DeleteColumnsOperation
  | AutofitColumnsOperation
  | AutofitRowsOperation
  | MergeRangeOperation
  | UnmergeRangeOperation
  | RestoreRangeSnapshotOperation
  | CreateSheetOperation
  | CopySheetOperation
  | RenameSheetOperation
  | DeleteSheetOperation
  | MoveSheetOperation
  | HideSheetOperation
  | UnhideSheetOperation
  | ProtectSheetOperation
  | UnprotectSheetOperation
  | ClearSheetOperation
  | SetSheetTabColorOperation
  | WorkbookCalculateOperation
  | WorkbookSaveOperation
  | CreateSheetFromTemplateOperation;

export interface BatchRequest {
  workbookId: WorkbookId;
  operations: ExcelOperation[];
  mode: "validate" | "dry_run" | "apply";
  planId?: PlanId | undefined;
  confirmationToken?: string;
  expectedTargetFingerprints?: RangeFingerprint[];
  baseSnapshotId?: SnapshotId;
  agentId?: AgentId | undefined;
  agentName?: string | undefined;
  taskId?: TaskId | undefined;
  role?: string | undefined;
  progressMessage?: string | undefined;
  retryStrategy?: string | undefined;
  chunksTotal?: number | undefined;
  chunksCompleted?: number | undefined;
}

export interface CompiledBatch {
  workbookId: WorkbookId;
  operations: ExcelOperation[];
  requiredBackups: Array<"region" | "sheet" | "workbook-copy">;
  targetFingerprints: RangeFingerprint[];
  estimatedCellsTouched: number;
  destructiveLevel: DestructiveLevel;
}

export type BatchExecutionMode = "apply" | "submit" | "chunked_submit";

export interface BatchChunkPlan {
  strategy: "none" | "split_style_entries" | "split_matrix_rows" | "mixed";
  chunksTotal: number;
  chunkSize: number;
  operationCount: number;
  chunkedOperationKinds: string[];
  safeToAutoChunk: boolean;
}

export interface BatchPreflightResult {
  ok: boolean;
  workbookId: WorkbookId;
  operationCount: number;
  estimatedCellsTouched: number;
  estimatedPayloadBytes: number;
  destructiveLevel: DestructiveLevel;
  recommendedExecutionMode: BatchExecutionMode;
  safeToAutoChunk: boolean;
  chunkPlan?: BatchChunkPlan | undefined;
  warnings: OperationWarning[];
}

export interface PlanCreateRequest {
  workbookId: WorkbookId;
  goal: string;
  operations: ExcelOperation[];
  baseSnapshotId?: SnapshotId;
  agentId?: AgentId | undefined;
  agentName?: string | undefined;
  taskId?: TaskId | undefined;
  role?: string | undefined;
}

export interface PlanPreview {
  planId: PlanId;
  workbookId: WorkbookId;
  baseSnapshotId: SnapshotId;
  requiredBackups: BackupId[];
  beforeWorkbookFingerprint: WorkbookFingerprint;
  targetFingerprints: RangeFingerprint[];
  diffSummary: DiffSummary;
  warnings: OperationWarning[];
}

export interface PlanRefreshResult {
  ok: boolean;
  planId: PlanId;
  refreshed: boolean;
  preview?: PlanPreview | undefined;
  conflicts: OperationWarning[];
  warnings: OperationWarning[];
}

export interface DiffSummary {
  title: string;
  changedRanges: A1Range[];
  cellsChanged: number;
  formulasChanged: number;
  stylesChanged: number;
  tablesChanged: number;
  sheetsChanged: number;
  destructiveLevel: DestructiveLevel;
}

export interface OperationResult {
  ok: boolean;
  operationId?: OperationId | undefined;
  planId?: PlanId | undefined;
  transactionId?: TransactionId | undefined;
  transactionStatus?: import("./collaboration.js").TransactionStatus | undefined;
  queuePosition?: number | undefined;
  progressMessage?: string | undefined;
  taskId?: TaskId | undefined;
  agentId?: AgentId | undefined;
  diffSummary?: DiffSummary;
  data?: unknown;
  rollbackAvailable: boolean;
  backups: BackupId[];
  warnings: OperationWarning[];
  telemetry: OperationTelemetry;
  error?: ExcelRuntimeError;
}

export interface RangeSnapshot {
  fingerprint: RangeFingerprint;
  values?: CellMatrix;
  formulas?: CellMatrix<string | null>;
  numberFormat?: string[][];
  text?: string[][];
  style?: {
    fillColor?: string;
    fontName?: string;
    fontSize?: number;
    fontColor?: string;
    fontBold?: boolean;
    fontItalic?: boolean;
    horizontalAlignment?: string;
    verticalAlignment?: string;
    rowHeight?: number;
    columnWidth?: number;
  };
}

export type StyleDimension =
  | "columnWidths"
  | "rowHeights"
  | "borders"
  | "fills"
  | "fonts"
  | "alignment"
  | "numberFormats"
  | "conditionalFormatting"
  | "dataValidation"
  | "freezePanes"
  | "printSettings"
  | "pageLayout"
  | "hiddenRowsColumns";

export interface StyleFingerprintRequest {
  workbookId: WorkbookId;
  sheetName: string;
  address?: string;
  maxCellSamples?: number;
}

export interface StyleFingerprintResponse {
  workbookId: WorkbookId;
  sheetName: string;
  address: string;
  capturedAt: string;
  rowCount: number;
  columnCount: number;
  truncated: boolean;
  dimensions: Partial<Record<StyleDimension, unknown>>;
  warnings: OperationWarning[];
}

export interface StyleCopyRequest {
  workbookId: WorkbookId;
  sourceSheetName: string;
  targetSheetName: string;
  sourceAddress?: string;
  targetAddress?: string;
  dimensions: StyleDimension[];
}

export interface StyleCopyResponse {
  ok: boolean;
  copied: StyleDimension[];
  warnings: OperationWarning[];
}

export interface StyleCompareResponse {
  ok: boolean;
  issueCount: number;
  issues: TemplateValidationIssue[];
  sourceFingerprint: StyleFingerprintResponse;
  targetFingerprint: StyleFingerprintResponse;
}

export interface FormulaPatternRequest {
  workbookId: WorkbookId;
  sheetName: string;
  address?: string;
}

export interface FormulaPatternCell {
  rowIndex: number;
  columnIndex: number;
  formula: string;
  formulaR1C1?: string;
  patternHash: string;
}

export interface FormulaSpillRange {
  sheetName?: string | undefined;
  anchorAddress: string;
  spillAddress: string;
}

export interface FormulaPatternResponse {
  workbookId: WorkbookId;
  sheetName: string;
  address: string;
  capturedAt: string;
  rowCount: number;
  columnCount: number;
  formulaCount: number;
  formulas: CellMatrix<string | null>;
  formulasR1C1?: CellMatrix<string | null>;
  patternMatrix: CellMatrix<string | null>;
  patterns: Array<{
    patternHash: string;
    formulaR1C1: string;
    count: number;
    cells: Array<{ rowIndex: number; columnIndex: number }>;
  }>;
  cells: FormulaPatternCell[];
  spillRanges?: FormulaSpillRange[] | undefined;
  warnings: OperationWarning[];
}

export interface FormulaDependencyNode {
  id: string;
  workbookId: WorkbookId;
  kind: "range" | "table" | "external";
  sheetName?: string | undefined;
  address?: string | undefined;
  tableName?: string | undefined;
  structuredReference?: string | undefined;
  externalWorkbook?: string | undefined;
  externalReference?: string | undefined;
  formula?: string | undefined;
}

export interface FormulaDependencyEdge {
  from: FormulaDependencyNode;
  to: FormulaDependencyNode;
  kind: "precedent";
  confidence: "parsed" | "inferred";
}

export interface FormulaDependencyGraph {
  workbookId: WorkbookId;
  sheetName: string;
  address: string;
  capturedAt: string;
  nodes: FormulaDependencyNode[];
  edges: FormulaDependencyEdge[];
  warnings: OperationWarning[];
}

export interface FormulaTraceResponse {
  ok: boolean;
  workbookId: WorkbookId;
  sheetName: string;
  address: string;
  direction: "precedents" | "dependents";
  nodes: FormulaDependencyNode[];
  edges: FormulaDependencyEdge[];
  warnings: OperationWarning[];
}

export interface FormulaCopyPatternsRequest {
  workbookId: WorkbookId;
  sourceSheetName: string;
  targetSheetName: string;
  sourceAddress?: string;
  targetAddress?: string;
}

export interface FormulaFillRequest {
  workbookId: WorkbookId;
  sheetName: string;
  sourceAddress: string;
  targetAddress: string;
  direction: "down" | "right";
}

export interface FormulaMutationResponse {
  ok: boolean;
  formulasChanged: number;
  warnings: OperationWarning[];
}

export interface FormulaCompareResponse {
  ok: boolean;
  issueCount: number;
  issues: TemplateValidationIssue[];
  sourcePatterns: FormulaPatternResponse;
  targetPatterns: FormulaPatternResponse;
}

export interface WorkbookSnapshotResponse {
  workbookFingerprint: WorkbookFingerprint;
  rangeSnapshots: RangeSnapshot[];
}

export interface TemplateExecutionSource {
  templateId: TemplateId;
  sourceSheetName: string;
  dataRegions: string[];
}

export interface AddinExecuteBatchRequest {
  request: BatchRequest;
  compiled: CompiledBatch;
  templateSources?: TemplateExecutionSource[];
}

export interface TemplateCaptureRequest {
  workbookId: WorkbookId;
  name: string;
  scope: "workbook" | "local";
  sourceSheetName: string;
  dataRegions: string[];
}

export interface TemplateCaptureResponse {
  sourceSheetName: string;
  dataRegions: string[];
  fingerprintPayload: {
    structure: unknown;
    formulas: unknown;
    styles: unknown;
    filters: unknown;
    tables: unknown;
    printLayout: unknown;
  };
}

export interface SheetTemplateFingerprintRequest {
  workbookId: WorkbookId;
  sheetName: string;
  dataRegions?: string[];
}

export interface SheetTemplateFingerprintResponse extends TemplateCaptureResponse {
  sheetName: string;
}

export interface TemplateValidationIssue {
  code: string;
  severity: "info" | "warning" | "error";
  message: string;
  component: "structure" | "formulas" | "styles" | "filters" | "tables" | "printLayout";
  target?: A1Range;
  expected?: unknown;
  actual?: unknown;
}

export interface TemplateValidationResponse {
  ok: boolean;
  sheetName: string;
  templateId: TemplateId;
  issueCount: number;
  issues: TemplateValidationIssue[];
  fingerprintPayload: TemplateCaptureResponse["fingerprintPayload"];
}

export interface TemplateRepairRequest {
  workbookId: WorkbookId;
  templateId: TemplateId;
  targetSheetName: string;
  repair: Array<"styles" | "formulas" | "dataRegions" | "layout">;
}

export interface AddinTemplateRepairRequest extends TemplateRepairRequest {
  sourceSheetName: string;
  dataRegions: string[];
}

export interface TableSelector {
  workbookId: WorkbookId;
  tableName: string;
}

export interface TableReadRequest extends TableSelector {
  includeValues?: boolean;
  includeFormulas?: boolean;
  includeText?: boolean;
  includeNumberFormats?: boolean;
  columns?: Array<string | number>;
  rowOffset?: number;
  rowLimit?: number;
}

export interface TableColumnRef {
  id?: number;
  index: number;
  name: string;
}

export interface TableInfo {
  workbookId: WorkbookId;
  tableName: string;
  id?: string;
  sheetName?: string;
  address?: string;
  headerAddress?: string;
  rowCount: number;
  columnCount: number;
  columns: TableColumnRef[];
  style?: string;
  showHeaders?: boolean;
  showTotals?: boolean;
  showFilterButton?: boolean;
  showBandedRows?: boolean;
  showBandedColumns?: boolean;
  filters?: unknown;
  sort?: unknown;
}

export interface TableReadResponse {
  info: TableInfo;
  headers: CellMatrix;
  values?: CellMatrix;
  formulas?: CellMatrix<string | null>;
  text?: string[][];
  numberFormat?: string[][];
  rowOffset?: number;
  rowLimit?: number;
  rowCount?: number;
  truncated?: boolean;
  projectedColumns?: TableColumnRef[];
}

export interface TableCreateRequest {
  workbookId: WorkbookId;
  sheetName: string;
  address: string;
  tableName?: string;
  hasHeaders: boolean;
  values?: CellMatrix;
  style?: string;
  showTotals?: boolean;
}

export interface TableResizeRequest extends TableSelector {
  address: string;
}

export interface TableAppendRowsRequest extends TableSelector {
  values: CellMatrix;
  index?: number;
  alwaysInsert?: boolean;
}

export interface TableUpdateRowsRequest extends TableSelector {
  rows: Array<{
    index: number;
    values: CellMatrix[number];
  }>;
}

export interface TableReorderColumnsRequest extends TableSelector {
  columnOrder: Array<string | number>;
}

export interface TableFilterSpec {
  column: string | number;
  criteria: unknown;
}

export interface TableApplyFiltersRequest extends TableSelector {
  filters: TableFilterSpec[];
}

export interface TableSortField {
  key: number;
  ascending?: boolean;
  sortOn?: "Value" | "CellColor" | "FontColor" | "Icon";
  color?: string;
  dataOption?: "Normal" | "TextAsNumber";
}

export interface TableSortRequest extends TableSelector {
  fields: TableSortField[];
  matchCase?: boolean;
  method?: "PinYin" | "StrokeCount";
}

export interface TableSetTotalRowRequest extends TableSelector {
  showTotals: boolean;
}

export interface TableSetStyleRequest extends TableSelector {
  style: string;
}

export interface TableCopyStructureRequest extends TableSelector {
  targetSheetName: string;
  targetAddress: string;
  newTableName?: string;
  includeStyle?: boolean;
  includeTotals?: boolean;
  includeFilters?: boolean;
}

export interface PivotSelector {
  workbookId: WorkbookId;
  pivotTableName: string;
}

export interface PivotFieldInfo {
  id?: string;
  name: string;
  showAllItems?: boolean;
  subtotals?: Record<string, unknown>;
}

export interface PivotAxisHierarchyInfo {
  id?: string;
  name: string;
  position?: number;
  enableMultipleFilterItems?: boolean;
  fields?: PivotFieldInfo[];
}

export interface PivotDataHierarchyInfo {
  id?: string;
  name: string;
  position?: number;
  numberFormat?: string;
  summarizeBy?: string;
  field?: PivotFieldInfo;
}

export interface PivotLayoutInfo {
  altTextDescription?: string;
  altTextTitle?: string;
  autoFormat?: boolean;
  emptyCellText?: string;
  enableFieldList?: boolean;
  fillEmptyCells?: boolean;
  layoutType?: string;
  preserveFormatting?: boolean;
  showColumnGrandTotals?: boolean;
  showFieldHeaders?: boolean;
  showRowGrandTotals?: boolean;
  subtotalLocation?: string;
}

export interface PivotRangeInfo {
  address: string;
  rowCount: number;
  columnCount: number;
}

export interface PivotTableInfo {
  workbookId: WorkbookId;
  pivotTableName: string;
  id?: string;
  sheetName?: string;
  range?: PivotRangeInfo;
  source?: string;
  sourceType?: string;
  refreshOnOpen?: boolean;
  useCustomSortLists?: boolean;
  enableDataValueEditing?: boolean;
  allowMultipleFiltersPerField?: boolean;
  layout?: PivotLayoutInfo;
  rowHierarchies?: PivotAxisHierarchyInfo[];
  columnHierarchies?: PivotAxisHierarchyInfo[];
  filterHierarchies?: PivotAxisHierarchyInfo[];
  dataHierarchies?: PivotDataHierarchyInfo[];
  hierarchies?: Array<{ id?: string; name: string }>;
}

export type CapabilityStatus = "supported" | "partial" | "unsupported" | "unknown";

export interface CapabilityStatusMetadata {
  capability: string;
  status: CapabilityStatus;
  reason?: string;
}

export type PivotCapabilityName =
  | "create"
  | "read_source_metadata"
  | "read_axis_fields"
  | "write_axis_fields"
  | "write_data_fields"
  | "aggregation"
  | "number_format"
  | "layout_flags"
  | "refresh"
  | "delete"
  | "template_copy"
  | "fingerprint"
  | "diff"
  | "rebuild_with_source"
  | "source_reassignment"
  | "pivot_chart";

export interface PivotCapabilityMatrix {
  workbookId?: WorkbookId;
  hostPlatform?: string;
  apiSets?: string[];
  capabilities: Array<CapabilityStatusMetadata & { capability: PivotCapabilityName }>;
}

export interface PivotOperationCapabilityStatus {
  operation: "pivot.update_source" | "pivot.copy_from_template" | "pivot.repair_from_template" | "pivot.rebuild_with_source";
  capabilities: CapabilityStatusMetadata[];
  warnings: OperationWarning[];
  fallback?: string;
}

export interface PivotCopyFromTemplateResponse {
  ok: boolean;
  copied?: string[];
  source?: PivotTableInfo;
  target?: PivotTableInfo;
  warnings?: OperationWarning[];
  capabilityStatus?: PivotOperationCapabilityStatus;
}

export interface PivotFingerprint {
  workbookId: WorkbookId;
  pivotTableName: string;
  capturedAt: string;
  hash: string;
  source?: {
    type?: string;
    value?: string;
    fields: string[];
  };
  layout: {
    rowFields: string[];
    columnFields: string[];
    filterFields: string[];
    dataFields: Array<{ name: string; sourceFieldName?: string; summarizeBy?: string; numberFormat?: string }>;
    flags?: PivotLayoutInfo;
  };
  output?: PivotRangeInfo & { sheetName?: string };
  warnings: Array<{ code: string; message: string }>;
}

export interface PivotDiff {
  ok: boolean;
  workbookId: WorkbookId;
  sourcePivotTableName: string;
  targetPivotTableName: string;
  source?: PivotFingerprint;
  target?: PivotFingerprint;
  changes: Array<{
    path: string;
    kind: "added" | "removed" | "changed";
    before?: unknown;
    after?: unknown;
  }>;
  warnings: Array<{ code: string; message: string }>;
}

export interface PivotCreateRequest {
  workbookId: WorkbookId;
  pivotTableName: string;
  sourceSheetName?: string;
  sourceAddress?: string;
  sourceTableName?: string;
  destinationSheetName: string;
  destinationAddress: string;
  rowFields?: string[];
  columnFields?: string[];
  filterFields?: string[];
  dataFields?: Array<{
    sourceFieldName: string;
    name?: string;
    summarizeBy?: string;
    numberFormat?: string;
  }>;
  layout?: PivotLayoutInfo;
  refresh?: boolean;
}

export interface PivotCopyFromTemplateRequest extends PivotSelector {
  templatePivotTableName: string;
  templateId?: TemplateId;
  dimensions?: Array<"metadata" | "layout" | "fields" | "dataFields" | "numberFormats" | "filters" | "refresh">;
  strict?: boolean;
}

export interface PivotValidateSourceRequest extends PivotSelector {
  expectedFields?: string[];
  expectedRowFields?: string[];
  expectedColumnFields?: string[];
  expectedFilterFields?: string[];
  expectedDataFields?: string[];
  expectedDataFieldSettings?: Array<{
    sourceFieldName?: string;
    name?: string;
    summarizeBy?: string;
    numberFormat?: string;
  }>;
  expectedLayout?: Partial<PivotLayoutInfo>;
}

export interface PivotCompareFingerprintRequest extends PivotSelector {
  targetPivotTableName: string;
}

export interface PivotRebuildWithSourceRequest extends PivotCreateRequest {
  templatePivotTableName?: string;
  replaceExisting?: boolean;
  strict?: boolean;
}

export type PivotRepairFromTemplateRequest = PivotCopyFromTemplateRequest;

export type WorkbookFileBridgeOperation = "workbook.save_as" | "workbook.export_copy" | "workbook.restore_file_backup";

export interface WorkbookFileBridgeRequest {
  operation: WorkbookFileBridgeOperation;
  workbookId: WorkbookId;
  targetPath?: string;
  backupPath?: string;
  restoreTargetPath?: string;
  restoreMode?: WorkbookFileRestoreMode;
  sourceBackupId?: BackupId;
  ranges?: A1Range[];
  reason?: string;
}

export interface WorkbookFileBridgeStatus {
  available: boolean;
  url?: string;
  path?: string;
  reason?: "not_configured" | "configured";
  reachable?: boolean;
  checkedAt?: string;
  statusCode?: number;
  bridge?: string;
  route?: string;
  adapter?: Record<string, unknown>;
  error?: string;
}

export interface WorkbookFileBridgeResponse {
  ok: boolean;
  operation: WorkbookFileBridgeOperation;
  workbookId: WorkbookId;
  targetPath?: string;
  backupPath?: string;
  restoreTargetPath?: string;
  restoreMode?: WorkbookFileRestoreMode;
  sourceBackupId?: BackupId;
  filePath?: string;
  metadata?: Record<string, unknown>;
  error?: string;
}

export type WorkbookFileBackupMode = "export-copy" | "save-copy-as";
export type WorkbookFileRestoreMode = "open-as-new" | "replace-open-workbook" | "restore-into-open-workbook";

export interface WorkbookFileBackupManifest {
  backupId: BackupId;
  workbookId: WorkbookId;
  createdAt: string;
  reason: string;
  filePath?: string;
  mode: WorkbookFileBackupMode;
  size?: number;
  checksum?: string;
  sourceSnapshotBackupId?: BackupId;
  pinned?: boolean;
  verifiedAt?: string;
  transactionId?: TransactionId;
  taskId?: TaskId;
  agentId?: AgentId;
  restoreStatus?: "available" | "missing" | "checksum_mismatch" | "unsupported";
  bridge?: WorkbookFileBridgeResponse;
  metadata?: Record<string, unknown>;
}

export interface WorkbookCreateFileBackupRequest {
  workbookId: WorkbookId;
  reason?: string;
  targetPath?: string;
  mode?: WorkbookFileBackupMode;
  pin?: boolean;
}

export interface WorkbookRestoreFileBackupRequest {
  workbookId: WorkbookId;
  backupId: BackupId;
  mode?: WorkbookFileRestoreMode;
  restoreTargetPath?: string;
  confirmationToken?: string;
  force?: boolean;
}

export interface WorkbookBackupRetentionRequest {
  workbookId?: WorkbookId;
  maxAgeDays?: number;
  maxBackupsPerWorkbook?: number;
  dryRun?: boolean;
}

export interface WorkbookFileContent {
  ok: true;
  workbookId: WorkbookId;
  fileType: "compressed";
  size: number;
  sliceCount: number;
  base64: string;
  capturedAt: string;
}

export interface ChartSelector {
  workbookId: WorkbookId;
  sheetName: string;
  chartName: string;
}

export interface ChartInfo {
  workbookId: WorkbookId;
  sheetName: string;
  chartName: string;
  id?: string;
  chartType?: string;
  title?: string;
  top?: number;
  left?: number;
  width?: number;
  height?: number;
  style?: number;
  plotBy?: string;
}

export interface ChartCreateRequest {
  workbookId: WorkbookId;
  sheetName: string;
  chartName?: string;
  sourceAddress: string;
  chartType: string;
  seriesBy?: "Auto" | "Columns" | "Rows";
  title?: string;
  position?: {
    startCell: string;
    endCell?: string;
  };
  style?: number;
}

export interface ChartUpdateDataSourceRequest extends ChartSelector {
  sourceAddress: string;
  seriesBy?: "Auto" | "Columns" | "Rows";
}

export interface RangeMetadataRequest {
  workbookId: WorkbookId;
  sheetName: string;
  address: string;
}

export interface RangeAreasSummary {
  address?: string;
  areaCount?: number;
  cellCount?: number;
  isNullObject?: boolean;
}

export interface RangeSearchRequest extends RangeMetadataRequest {
  text: string;
  completeMatch?: boolean;
  matchCase?: boolean;
  searchDirection?: "Forward" | "Backwards";
}

export interface RangeSearchResponse {
  ok: boolean;
  matches: RangeAreasSummary;
}

export interface RangeMetadataResponse {
  ok: boolean;
  target: A1Range;
  data?: unknown;
  warnings: OperationWarning[];
}

export type NameScope = "workbook" | "worksheet";

export interface NameInfo {
  workbookId: WorkbookId;
  name: string;
  scope: NameScope;
  sheetName?: string;
  type?: string;
  value?: unknown;
  formula?: string;
  comment?: string;
  visible?: boolean;
  address?: string;
}

export interface NameSelector {
  workbookId: WorkbookId;
  name: string;
  sheetName?: string;
}

export interface NameCreateRequest extends NameSelector {
  reference?: string;
  formula?: string;
  comment?: string;
  visible?: boolean;
}

export interface NameUpdateRequest extends NameSelector {
  reference?: string;
  formula?: string;
  comment?: string;
  visible?: boolean;
}

export interface WorkbookRegion {
  workbookId: WorkbookId;
  regionId: string;
  name: string;
  sheetName: string;
  address: string;
  kind: "data" | "header" | "formula" | "output" | "template" | "table" | "named-range" | "other";
  source?: "manual" | "detected" | "named-range" | "table" | "template";
  description?: string;
  templateId?: TemplateId;
  namedItem?: string;
  createdAt: string;
  updatedAt: string;
}

export interface RegionSelector {
  workbookId: WorkbookId;
  regionName: string;
}

export interface RegionRegisterRequest {
  workbookId: WorkbookId;
  name: string;
  sheetName: string;
  address: string;
  kind?: WorkbookRegion["kind"];
  description?: string;
  templateId?: TemplateId;
  createNamedRange?: boolean;
}

export interface PermissionScope {
  workbookId?: WorkbookId;
  sheetNames?: string[];
  regionNames?: string[];
}

export interface LockedRegion {
  workbookId: WorkbookId;
  regionName: string;
  sheetName: string;
  address: string;
  reason?: string;
  lockedAt: string;
}

export interface PermissionState {
  allowWrites: boolean;
  allowDestructiveActions: boolean;
  allowWorkbookActions: boolean;
  allowMacroExecution: boolean;
  requireConfirmationFor: DestructiveLevel[];
  scope: PermissionScope;
  lockedRegions: LockedRegion[];
}

export interface WorkbookLocalConfig {
  version: 1;
  workbookId: WorkbookId;
  exportedAt: string;
  source: "open-workbook-local-config";
  templates: Array<Record<string, unknown>>;
  regions: WorkbookRegion[];
  permissions?: PermissionState;
}

export interface WorkbookLocalConfigImportRequest {
  workbookId: WorkbookId;
  config: WorkbookLocalConfig;
  includeTemplates?: boolean;
  includeRegions?: boolean;
  includePermissions?: boolean;
  overwrite?: boolean;
}

export interface WorkbookLocalConfigImportResponse {
  ok: boolean;
  workbookId: WorkbookId;
  imported: {
    templates: number;
    regions: number;
    permissions: boolean;
  };
  skipped: {
    templates: number;
    regions: number;
  };
  error?: ExcelRuntimeError;
}

export interface WorkbookEmbeddedLocalConfigResponse {
  ok: boolean;
  workbookId: WorkbookId;
  embedded: boolean;
  config?: WorkbookLocalConfig;
  partCount?: number;
  warnings?: OperationWarning[];
  error?: ExcelRuntimeError;
}

export interface CleaningReport {
  ok: boolean;
  workbookId: WorkbookId;
  target?: A1Range;
  action: string;
  changedCells: number;
  affectedRows?: number;
  affectedColumns?: number;
  data?: unknown;
  result?: OperationResult;
  warnings: OperationWarning[];
  error?: ExcelRuntimeError;
}

export type ValidationSeverity = "info" | "warning" | "error";
export type ValidationCategory =
  | "workbook"
  | "sheet"
  | "template"
  | "formula"
  | "style"
  | "table"
  | "filter"
  | "printLayout"
  | "reference"
  | "change"
  | "capability";

export interface ValidationIssue {
  code: string;
  severity: ValidationSeverity;
  category: ValidationCategory;
  message: string;
  target?: A1Range;
  details?: Record<string, unknown>;
}

export interface ValidationReport {
  ok: boolean;
  workbookId: WorkbookId;
  scope: string;
  checkedAt: string;
  issueCount: number;
  summary: {
    errorCount: number;
    warningCount: number;
    infoCount: number;
  };
  issues: ValidationIssue[];
  data?: unknown;
}

export interface RepairReport {
  ok: boolean;
  workbookId: WorkbookId;
  repair: string;
  repairedAt: string;
  backups: BackupId[];
  result?: unknown;
  validation?: ValidationReport | TemplateValidationResponse;
  warnings: OperationWarning[];
  error?: ExcelRuntimeError;
}
