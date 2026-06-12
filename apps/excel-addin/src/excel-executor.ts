import { createRangeFingerprint, createWorkbookFingerprint } from "@open-workbook/excel-core";
import type {
  AddinExecuteBatchRequest,
  A1Range,
  DiffSummary,
  ExcelOperation,
  OperationResult,
  OperationTelemetry,
  OperationWarning,
  RangeSnapshot,
  TemplateExecutionSource,
  WorkbookRef,
  WorkbookSnapshotResponse
} from "@open-workbook/protocol";
import { runtimeError } from "@open-workbook/protocol";

interface LoadedRangeSnapshot {
  target: A1Range;
  range: Excel.Range;
}

interface ExecutionCounters {
  syncCount: number;
  cellsRead: number;
  cellsWritten: number;
  rangeCount: number;
  chunkCount: number;
}

const ENGINE_NAME = "office-js-addin";
const ENGINE_VERSION = "0.1.0";
const CHUNK_CELL_LIMIT = 50_000;

export async function getActiveWorkbookContext(): Promise<WorkbookRef | undefined> {
  return Excel.run(async (context) => {
    const workbook = context.workbook;
    const activeWorksheet = workbook.worksheets.getActiveWorksheet();
    workbook.load("name");
    activeWorksheet.load("name");
    await context.sync();

    return {
      workbookId: workbook.name as WorkbookRef["workbookId"],
      name: workbook.name,
      platform: detectPlatform()
    };
  });
}

export async function snapshotRanges(workbookId: string, ranges: A1Range[]): Promise<WorkbookSnapshotResponse> {
  return Excel.run(async (context) => {
    const loaded: LoadedRangeSnapshot[] = [];

    for (const target of ranges) {
      const range = getRange(context, target);
      loadSnapshotProperties(range);
      loaded.push({ target, range });
    }

    const workbook = context.workbook;
    workbook.load("name, worksheets/items/name");
    await context.sync();

    const rangeSnapshots = loaded.map(({ target, range }) => materializeSnapshot(target, range));
    return {
      workbookFingerprint: createWorkbookFingerprint(
        workbookId as WorkbookRef["workbookId"],
        {
          workbookName: workbook.name,
          ranges: rangeSnapshots.map((snapshot) => snapshot.fingerprint.hash)
        },
        {
          sheets: workbook.worksheets.items.map((worksheet) => worksheet.name)
        }
      ),
      rangeSnapshots
    };
  });
}

export async function executeBatch(payload: AddinExecuteBatchRequest): Promise<OperationResult> {
  const started = performance.now();
  const counters: ExecutionCounters = {
    syncCount: 0,
    cellsRead: 0,
    cellsWritten: 0,
    rangeCount: payload.compiled.targetFingerprints.length,
    chunkCount: Math.max(1, Math.ceil(payload.compiled.estimatedCellsTouched / CHUNK_CELL_LIMIT))
  };
  const warnings: OperationWarning[] = [];

  try {
    if (payload.request.expectedTargetFingerprints?.length) {
      const conflictWarnings = await detectTargetConflicts(
        payload.request.workbookId,
        payload.request.expectedTargetFingerprints
      );
      if (conflictWarnings.length > 0) {
        return {
          ok: false,
          rollbackAvailable: true,
          backups: [],
          warnings: conflictWarnings,
          telemetry: createTelemetry(started, counters, conflictWarnings),
          error: runtimeError("EXTERNAL_CHANGE_DETECTED", "Target ranges changed after preview. Refresh the plan before applying.", {
            retryable: true
          })
        };
      }
    }

    const result = await Excel.run(async (context) => {
      const readOperations: Array<{ operation: ExcelOperation; range: Excel.Range }> = [];
      let formulasChanged = 0;
      let sheetsChanged = 0;

      maybeSuspendExcel(context, payload.compiled.estimatedCellsTouched);

      for (const operation of payload.request.operations) {
        switch (operation.kind) {
          case "range.read_full": {
            const range = getRange(context, operation.target);
            loadSnapshotProperties(range);
            readOperations.push({ operation, range });
            counters.cellsRead += payload.compiled.estimatedCellsTouched;
            break;
          }
          case "range.write_values": {
            assertMatrixShape(operation.target, operation.values);
            getRange(context, operation.target).values = operation.values;
            counters.cellsWritten += matrixCellCount(operation.values);
            break;
          }
          case "range.write_formulas": {
            assertMatrixShape(operation.target, operation.formulas);
            getRange(context, operation.target).formulas = operation.formulas;
            counters.cellsWritten += matrixCellCount(operation.formulas);
            formulasChanged += matrixCellCount(operation.formulas);
            break;
          }
          case "range.clear_values_keep_format": {
            getRange(context, operation.target).clear(Excel.ClearApplyTo.contents);
            counters.cellsWritten += payload.compiled.estimatedCellsTouched;
            break;
          }
          case "template.create_sheet_from_template": {
            const warning = applyTemplateSheetOperation(context, operation, payload.templateSources ?? []);
            if (warning) {
              warnings.push(warning);
            } else {
              sheetsChanged += 1;
            }
            break;
          }
        }
      }

      await context.sync();
      counters.syncCount += 1;

      const readData = readOperations.map(({ operation, range }) => ({
        operationId: operation.operationId,
        snapshot: materializeSnapshot("target" in operation ? operation.target : payload.compiled.targetFingerprints[0]!.range, range)
      }));

      const changedRanges = payload.compiled.targetFingerprints.map((fingerprint) => fingerprint.range);
      const diffSummary: DiffSummary = {
        title: "Excel batch applied",
        changedRanges,
        cellsChanged: payload.compiled.estimatedCellsTouched,
        formulasChanged,
        stylesChanged: 0,
        tablesChanged: 0,
        sheetsChanged,
        destructiveLevel: payload.compiled.destructiveLevel
      };

      return { diffSummary, readData };
    });

    return {
      ok: warnings.every((warning) => warning.code !== "TEMPLATE_SOURCE_MISSING"),
      diffSummary: result.diffSummary,
      data: result.readData,
      rollbackAvailable: payload.compiled.requiredBackups.length > 0,
      backups: [],
      warnings,
      telemetry: createTelemetry(started, counters, warnings)
    };
  } catch (error) {
    const message = error instanceof Error ? error.message : String(error);
    return {
      ok: false,
      rollbackAvailable: payload.compiled.requiredBackups.length > 0,
      backups: [],
      warnings,
      telemetry: createTelemetry(started, counters, warnings),
      error: runtimeError("OPERATION_FAILED", message, { retryable: false })
    };
  }
}

async function detectTargetConflicts(workbookId: string, expected: NonNullable<AddinExecuteBatchRequest["request"]["expectedTargetFingerprints"]>) {
  const current = await snapshotRanges(
    workbookId,
    expected.map((fingerprint) => fingerprint.range)
  );
  const warnings: OperationWarning[] = [];

  for (const expectedFingerprint of expected) {
    const currentFingerprint = current.rangeSnapshots.find(
      (snapshot) =>
        snapshot.fingerprint.range.sheetName === expectedFingerprint.range.sheetName &&
        snapshot.fingerprint.range.address === expectedFingerprint.range.address
    )?.fingerprint;

    if (!currentFingerprint || currentFingerprint.hash !== expectedFingerprint.hash) {
      warnings.push({
        code: "TARGET_REGION_CHANGED",
        message: `Target changed after preview: ${expectedFingerprint.range.sheetName}!${expectedFingerprint.range.address}`,
        target: expectedFingerprint.range
      });
    }
  }

  return warnings;
}

function getRange(context: Excel.RequestContext, target: A1Range): Excel.Range {
  return context.workbook.worksheets.getItem(target.sheetName).getRange(stripSheetName(target.address));
}

function loadSnapshotProperties(range: Excel.Range): void {
  range.load("values, formulas, numberFormat, text, rowCount, columnCount");
  range.format.load("horizontalAlignment, verticalAlignment, rowHeight, columnWidth");
  range.format.fill.load("color");
  range.format.font.load("name, size, color, bold, italic");
}

function materializeSnapshot(target: A1Range, range: Excel.Range): RangeSnapshot {
  const values = range.values as RangeSnapshot["values"];
  const formulas = range.formulas as RangeSnapshot["formulas"];
  const numberFormat = range.numberFormat as string[][];
  const text = range.text as string[][];
  const style: NonNullable<RangeSnapshot["style"]> = {};
  assignIfDefined(style, "fillColor", optionalValue(range.format.fill.color));
  assignIfDefined(style, "fontName", optionalValue(range.format.font.name));
  assignIfDefined(style, "fontSize", optionalValue(range.format.font.size));
  assignIfDefined(style, "fontColor", optionalValue(range.format.font.color));
  assignIfDefined(style, "fontBold", optionalValue(range.format.font.bold));
  assignIfDefined(style, "fontItalic", optionalValue(range.format.font.italic));
  assignIfDefined(style, "horizontalAlignment", optionalValue(String(range.format.horizontalAlignment)));
  assignIfDefined(style, "verticalAlignment", optionalValue(String(range.format.verticalAlignment)));
  assignIfDefined(style, "rowHeight", optionalValue(range.format.rowHeight));
  assignIfDefined(style, "columnWidth", optionalValue(range.format.columnWidth));

  const snapshot: RangeSnapshot = {
    fingerprint: createRangeFingerprint(target, {
      values,
      formulas,
      numberFormat,
      text,
      style
    }),
    numberFormat,
    text,
    style
  };
  if (values !== undefined) {
    snapshot.values = values;
  }
  if (formulas !== undefined) {
    snapshot.formulas = formulas;
  }
  return snapshot;
}

function applyTemplateSheetOperation(
  context: Excel.RequestContext,
  operation: Extract<ExcelOperation, { kind: "template.create_sheet_from_template" }>,
  templateSources: TemplateExecutionSource[]
): OperationWarning | undefined {
  const source = templateSources.find((templateSource) => templateSource.templateId === operation.templateId);
  if (!source) {
    return {
      code: "TEMPLATE_SOURCE_MISSING",
      message: `Template source not registered for ${operation.templateId}.`,
      details: { templateId: operation.templateId }
    };
  }

  const sourceSheet = context.workbook.worksheets.getItem(source.sourceSheetName);
  const copiedSheet = sourceSheet.copy(Excel.WorksheetPositionType.after, sourceSheet);
  copiedSheet.name = operation.newSheetName;

  if (operation.clearDataRegions) {
    for (const dataRegion of source.dataRegions) {
      copiedSheet.getRange(stripSheetName(dataRegion)).clear(Excel.ClearApplyTo.contents);
    }
  }

  copiedSheet.activate();
  return undefined;
}

function maybeSuspendExcel(context: Excel.RequestContext, estimatedCellsTouched: number): void {
  if (estimatedCellsTouched >= 10_000) {
    context.workbook.application.suspendApiCalculationUntilNextSync();
    context.workbook.application.suspendScreenUpdatingUntilNextSync();
  }
}

function assertMatrixShape(target: A1Range, matrix: unknown[][]): void {
  if (matrix.length === 0 || matrix.some((row) => row.length !== matrix[0]!.length)) {
    throw new Error(`Invalid matrix shape for ${target.sheetName}!${target.address}`);
  }
}

function matrixCellCount(matrix: unknown[][]): number {
  return matrix.reduce((count, row) => count + row.length, 0);
}

function optionalValue<T>(value: T | null | undefined): T | undefined {
  return value === null || value === undefined ? undefined : value;
}

function assignIfDefined<T extends Record<string, unknown>, K extends keyof T>(target: T, key: K, value: T[K] | undefined): void {
  if (value !== undefined) {
    target[key] = value;
  }
}

function stripSheetName(address: string): string {
  const bangIndex = address.lastIndexOf("!");
  return bangIndex >= 0 ? address.slice(bangIndex + 1) : address;
}

function createTelemetry(started: number, counters: ExecutionCounters, warnings: OperationWarning[]): OperationTelemetry {
  return {
    durationMs: Math.round(performance.now() - started),
    syncCount: counters.syncCount,
    payloadBytes: 0,
    cellsRead: counters.cellsRead,
    cellsWritten: counters.cellsWritten,
    rangeCount: counters.rangeCount,
    chunkCount: counters.chunkCount,
    engineName: ENGINE_NAME,
    engineVersion: ENGINE_VERSION,
    warningCount: warnings.length
  };
}

function detectPlatform(): WorkbookRef["platform"] {
  if (Office.context.platform === Office.PlatformType.Mac) {
    return "mac";
  }
  if (Office.context.platform === Office.PlatformType.PC) {
    return "windows";
  }
  if (Office.context.platform === Office.PlatformType.OfficeOnline) {
    return "web";
  }
  return "unknown";
}
