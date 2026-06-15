import { chunkMatrixRows, createRangeFingerprint, createWorkbookFingerprint, formatA1Cell, hashStable, parseA1Address } from "@components-kit/open-workbook-excel-core";
import type {
  AddinExecuteBatchRequest,
  AddinTemplateRepairRequest,
  A1Range,
  CellPosition,
  ChartCreateRequest,
  ChartInfo,
  ChartSelector,
  ChartUpdateDataSourceRequest,
  DiffSummary,
  ExcelOperation,
  FormulaCopyPatternsRequest,
  FormulaFillRequest,
  FormulaMutationResponse,
  FormulaPatternRequest,
  FormulaPatternResponse,
  NameCreateRequest,
  NameInfo,
  NameSelector,
  NameUpdateRequest,
  OperationResult,
  OperationTelemetry,
  OperationWarning,
  PivotCopyFromTemplateResponse,
  PivotCopyFromTemplateRequest,
  PivotCreateRequest,
  PivotSelector,
  PivotTableInfo,
  RangeMetadataRequest,
  RangeMetadataResponse,
  RangeSearchRequest,
  RangeSearchResponse,
  RangeSnapshot,
  RuntimeSelectionResponse,
  TemplateCaptureRequest,
  TemplateCaptureResponse,
  SheetTemplateFingerprintRequest,
  SheetTemplateFingerprintResponse,
  StyleCopyRequest,
  StyleCopyResponse,
  StyleDimension,
  StyleFingerprintRequest,
  StyleFingerprintResponse,
  TableAppendRowsRequest,
  TableApplyFiltersRequest,
  TableCopyStructureRequest,
  TableCreateRequest,
  TableInfo,
  TableReadRequest,
  TableReadResponse,
  TableReorderColumnsRequest,
  TableResizeRequest,
  TableSelector,
  TableSetStyleRequest,
  TableSetTotalRowRequest,
  TableSortRequest,
  TableUpdateRowsRequest,
  TemplateExecutionSource,
  RuntimeCapabilities,
  WorkbookFileContent,
  WorkbookRef,
  WorkbookLocalConfig,
  WorkbookSnapshotResponse
} from "@components-kit/open-workbook-protocol";
import { runtimeError } from "@components-kit/open-workbook-protocol";

interface LoadedRangeSnapshot {
  target: A1Range;
  range: Excel.Range;
  facets?: RangeSnapshotFacet[];
}

type RangeSnapshotFacet = "values" | "formulas" | "numberFormat" | "text" | "style";

interface ExecutionCounters {
  syncCount: number;
  cellsRead: number;
  cellsWritten: number;
  rangeCount: number;
  chunkCount: number;
}

const ENGINE_NAME = "office-js-addin";
const ENGINE_VERSION = "0.1.9";
const CHUNK_CELL_LIMIT = 50_000;
const OPEN_WORKBOOK_CUSTOM_XML_NAMESPACE = "https://open-workbook.dev/schema/local-config/1";
const EXCEL_API_VERSIONS = ["1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8", "1.9", "1.10", "1.11", "1.12", "1.13", "1.14", "1.15", "1.16", "1.17"] as const;

export function getRuntimeCapabilities(): RuntimeCapabilities {
  const apiSets = EXCEL_API_VERSIONS.map((version) => ({
    set: "ExcelApi",
    version,
    supported: isOfficeApiSetSupported("ExcelApi", version)
  }));
  const supports = (version: (typeof EXCEL_API_VERSIONS)[number]) => apiSets.some((apiSet) => apiSet.version === version && apiSet.supported);
  const officeVersion = typeof Office.context.diagnostics?.version === "string" ? Office.context.diagnostics.version : undefined;
  const platform = detectPlatform();
  const supportsCompressedFileExport = typeof Office.context.document?.getFileAsync === "function" && platform !== "web";
  return {
    engine: {
      name: ENGINE_NAME,
      version: ENGINE_VERSION,
      platform,
      host: String(Office.context.host ?? "Excel"),
      ...(officeVersion !== undefined ? { officeVersion } : {})
    },
    apiSets,
    capabilities: [
      {
        name: "workbook.context",
        supported: supports("1.1"),
        platforms: ["mac", "windows", "web"],
        requires: [{ set: "ExcelApi", version: "1.1" }]
      },
      {
        name: "range.batch.read_write",
        supported: supports("1.9"),
        platforms: ["mac", "windows", "web"],
        requires: [{ set: "ExcelApi", version: "1.9" }]
      },
      {
        name: "table.native",
        supported: supports("1.9"),
        platforms: ["mac", "windows", "web"],
        requires: [{ set: "ExcelApi", version: "1.9" }]
      },
      {
        name: "pivot.native",
        supported: supports("1.8"),
        platforms: ["mac", "windows", "web"],
        requires: [{ set: "ExcelApi", version: "1.8" }]
      },
      {
        name: "chart.native",
        supported: supports("1.9"),
        platforms: ["mac", "windows", "web"],
        requires: [{ set: "ExcelApi", version: "1.9" }]
      },
      {
        name: "range.metadata.advanced",
        supported: supports("1.9"),
        platforms: ["mac", "windows", "web"],
        requires: [{ set: "ExcelApi", version: "1.9" }]
      }
    ],
    hostCapabilities: [
      hostCapability("range-values-formulas-styles", supports("1.9"), "ExcelApi", "1.9"),
      hostCapability("tables-filters-sorts", supports("1.9"), "ExcelApi", "1.9"),
      hostCapability("pivots", supports("1.8"), "ExcelApi", "1.8"),
      hostCapability("charts", supports("1.9"), "ExcelApi", "1.9"),
      {
        name: "workbook-compressed-file-export",
        supported: supportsCompressedFileExport,
        status: supportsCompressedFileExport ? "supported" : "unsupported",
        reason: supportsCompressedFileExport
          ? "Office Document.getFileAsync supports compressed workbook slices on this host."
          : "Compressed workbook file export is supported by Excel desktop hosts, not Excel on the web."
      },
      {
        name: "workbook-save-as-local-path",
        supported: false,
        status: "unsupported",
        reason: "Office.js does not expose a deterministic local Save As path API."
      },
      {
        name: "theme-freeze-print-layout-replay",
        supported: false,
        status: "limited",
        reason: "Current implementation reports layout capability status instead of replaying these dimensions."
      },
      {
        name: "comments-notes-address-mapping",
        supported: false,
        status: "limited",
        reason: "Office.js comment and legacy-note collections need deterministic address mapping before agent writes are enabled."
      }
    ]
  };
}

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

export async function getSelection(): Promise<RuntimeSelectionResponse> {
  return Excel.run(async (context) => {
    const workbook = context.workbook;
    const selectedRange = workbook.getSelectedRange();
    selectedRange.load("address,rowIndex,columnIndex,rowCount,columnCount");
    selectedRange.worksheet.load("name");
    workbook.load("name");
    await context.sync();

    const workbookRef: WorkbookRef = {
      workbookId: workbook.name as WorkbookRef["workbookId"],
      name: workbook.name,
      platform: detectPlatform()
    };

    return {
      workbook: workbookRef,
      selection: {
        workbookId: workbookRef.workbookId,
        sheetName: selectedRange.worksheet.name,
        address: stripSheetName(selectedRange.address),
        startCell: cellPositionFromZeroBased(
          workbookRef.workbookId,
          selectedRange.worksheet.name,
          selectedRange.rowIndex,
          selectedRange.columnIndex
        ),
        endCell: cellPositionFromZeroBased(
          workbookRef.workbookId,
          selectedRange.worksheet.name,
          selectedRange.rowIndex + selectedRange.rowCount - 1,
          selectedRange.columnIndex + selectedRange.columnCount - 1
        ),
        rowCount: selectedRange.rowCount,
        columnCount: selectedRange.columnCount,
        cellCount: selectedRange.rowCount * selectedRange.columnCount,
        isSingleCell: selectedRange.rowCount === 1 && selectedRange.columnCount === 1
      }
    };
  });
}

export async function setActiveSheet(sheetName: string): Promise<{ ok: boolean; activeSheet: string }> {
  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(sheetName);
    worksheet.activate();
    await context.sync();
    return {
      ok: true,
      activeSheet: sheetName
    };
  });
}

export async function getWorkbookInfo(): Promise<{
  workbook: WorkbookRef;
  activeSheet: string;
  worksheetCount: number;
}> {
  return Excel.run(async (context) => {
    const workbook = context.workbook;
    const activeWorksheet = workbook.worksheets.getActiveWorksheet();
    workbook.load("name");
    activeWorksheet.load("name");
    workbook.worksheets.load("items/name");
    await context.sync();

    return {
      workbook: {
        workbookId: workbook.name as WorkbookRef["workbookId"],
        name: workbook.name,
        platform: detectPlatform()
      },
      activeSheet: activeWorksheet.name,
      worksheetCount: workbook.worksheets.items.length
    };
  });
}

export async function getWorkbookMap(): Promise<{
  workbook: WorkbookRef;
  sheets: Array<{
    name: string;
    position: number;
    visibility: string;
    usedRange?: {
      address: string;
      rowCount: number;
      columnCount: number;
    };
    tables: Array<{ name: string }>;
  }>;
}> {
  return Excel.run(async (context) => {
    const workbook = context.workbook;
    workbook.load("name");
    workbook.worksheets.load("items/name,items/position,items/visibility");
    await context.sync();

    const loaded = workbook.worksheets.items.map((worksheet) => {
      const usedRange = worksheet.getUsedRangeOrNullObject();
      const tables = worksheet.tables;
      usedRange.load("address,rowCount,columnCount");
      tables.load("items/name");
      return { worksheet, usedRange, tables };
    });
    await context.sync();

    return {
      workbook: {
        workbookId: workbook.name as WorkbookRef["workbookId"],
        name: workbook.name,
        platform: detectPlatform()
      },
      sheets: loaded.map(({ worksheet, usedRange, tables }) => {
        const sheet: {
          name: string;
          position: number;
          visibility: string;
          usedRange?: {
            address: string;
            rowCount: number;
            columnCount: number;
          };
          tables: Array<{ name: string }>;
        } = {
          name: worksheet.name,
          position: worksheet.position,
          visibility: String(worksheet.visibility),
          tables: tables.items.map((table) => ({ name: table.name }))
        };
        if (!usedRange.isNullObject) {
          sheet.usedRange = {
            address: usedRange.address,
            rowCount: usedRange.rowCount,
            columnCount: usedRange.columnCount
          };
        }
        return sheet;
      })
    };
  });
}

export async function calculateWorkbook(calculationType: "full" | "recalculate" = "full"): Promise<{ ok: boolean; calculationType: string }> {
  return Excel.run(async (context) => {
    context.workbook.application.calculate(
      calculationType === "recalculate" ? Excel.CalculationType.recalculate : Excel.CalculationType.full
    );
    await context.sync();
    return { ok: true, calculationType };
  });
}

export async function saveWorkbook(): Promise<{ ok: boolean }> {
  return Excel.run(async (context) => {
    context.workbook.save(Excel.SaveBehavior.save);
    await context.sync();
    return { ok: true };
  });
}

export async function exportWorkbookFile(workbookId: string, sliceSize = 4 * 1024 * 1024): Promise<WorkbookFileContent | { ok: false; error: ReturnType<typeof runtimeError> }> {
  const document = Office.context.document;
  if (!document || typeof document.getFileAsync !== "function") {
    return {
      ok: false,
      error: runtimeError("CAPABILITY_UNAVAILABLE", "This Excel host does not expose Office Document.getFileAsync.", { retryable: false })
    };
  }
  if (detectPlatform() === "web") {
    return {
      ok: false,
      error: runtimeError("CAPABILITY_UNAVAILABLE", "Excel on the web does not expose compressed workbook export through Office.js.", { retryable: false })
    };
  }

  const file = await getDocumentFile(sliceSize);
  try {
    const chunks: string[] = [];
    for (let index = 0; index < file.sliceCount; index += 1) {
      const slice = await getDocumentFileSlice(file, index);
      chunks.push(sliceDataToBase64(slice.data));
    }
    return {
      ok: true,
      workbookId: workbookId as WorkbookFileContent["workbookId"],
      fileType: "compressed",
      size: file.size,
      sliceCount: file.sliceCount,
      base64: chunks.join(""),
      capturedAt: new Date().toISOString()
    };
  } finally {
    await closeDocumentFile(file);
  }
}

export async function closeWorkbook(closeBehavior: "Save" | "SkipSave" = "Save"): Promise<{ ok: boolean; closeBehavior: string }> {
  return Excel.run(async (context) => {
    context.workbook.close(closeBehavior);
    await context.sync();
    return { ok: true, closeBehavior };
  });
}

export async function listNames(workbookId: string): Promise<{ ok: boolean; names: NameInfo[] }> {
  return Excel.run(async (context) => {
    const workbookNames = context.workbook.names;
    workbookNames.load("items/name,items/scope,items/type,items/value,items/formula,items/comment,items/visible");
    context.workbook.worksheets.load("items/name");
    await context.sync();

    const worksheetCollections = context.workbook.worksheets.items.map((worksheet) => {
      worksheet.names.load("items/name,items/scope,items/type,items/value,items/formula,items/comment,items/visible");
      return { worksheet, names: worksheet.names };
    });
    await context.sync();

    const names = [
      ...workbookNames.items.map((item) => materializeNameInfo(workbookId, item)),
      ...worksheetCollections.flatMap(({ worksheet, names }) => names.items.map((item) => materializeNameInfo(workbookId, item, worksheet.name)))
    ];
    return { ok: true, names };
  });
}

export async function getName(request: NameSelector): Promise<{ ok: boolean; name?: NameInfo }> {
  return Excel.run(async (context) => {
    const item = getNamedItem(context, request);
    const range = loadNameWithRange(item);
    await context.sync();
    if (item.isNullObject) {
      return { ok: false };
    }
    return { ok: true, name: materializeNameInfo(request.workbookId, item, request.sheetName, range) };
  });
}

export async function createName(request: NameCreateRequest): Promise<{ ok: boolean; name: NameInfo }> {
  return Excel.run(async (context) => {
    const collection = request.sheetName ? context.workbook.worksheets.getItem(request.sheetName).names : context.workbook.names;
    const item = collection.add(request.name, nameReference(context, request), request.comment);
    if (request.visible !== undefined) {
      item.visible = request.visible;
    }
    const range = loadNameWithRange(item);
    await context.sync();
    return { ok: true, name: materializeNameInfo(request.workbookId, item, request.sheetName, range) };
  });
}

export async function updateName(request: NameUpdateRequest): Promise<{ ok: boolean; name?: NameInfo }> {
  return Excel.run(async (context) => {
    const item = getNamedItem(context, request);
    item.load("name");
    await context.sync();
    if (item.isNullObject) {
      return { ok: false };
    }
    if (request.reference !== undefined || request.formula !== undefined) {
      item.formula = nameFormula(request);
    }
    if (request.comment !== undefined) {
      item.comment = request.comment;
    }
    if (request.visible !== undefined) {
      item.visible = request.visible;
    }
    const range = loadNameWithRange(item);
    await context.sync();
    return { ok: true, name: materializeNameInfo(request.workbookId, item, request.sheetName, range) };
  });
}

export async function deleteName(request: NameSelector): Promise<{ ok: boolean; deleted: boolean }> {
  return Excel.run(async (context) => {
    const item = getNamedItem(context, request);
    item.load("name");
    await context.sync();
    if (item.isNullObject) {
      return { ok: false, deleted: false };
    }
    item.delete();
    await context.sync();
    return { ok: true, deleted: true };
  });
}

export async function embedWorkbookLocalConfig(request: {
  workbookId: string;
  config: WorkbookLocalConfig;
}): Promise<{ ok: boolean; embedded: boolean; partCount: number; namespaceUri: string; error?: unknown }> {
  return Excel.run(async (context) => {
    const customXmlParts = getCustomXmlParts(context);
    if (!customXmlParts) {
      return {
        ok: false,
        embedded: false,
        partCount: 0,
        namespaceUri: OPEN_WORKBOOK_CUSTOM_XML_NAMESPACE,
        error: runtimeError("CAPABILITY_UNAVAILABLE", "This Excel host does not expose workbook custom XML parts to Office.js.", { retryable: false })
      };
    }
    const existing = customXmlParts.getByNamespace(OPEN_WORKBOOK_CUSTOM_XML_NAMESPACE);
    existing.load("items/id");
    await context.sync();
    for (const part of existing.items ?? []) {
      part.delete();
    }
    await context.sync();
    customXmlParts.add(workbookLocalConfigXml(request.config));
    await context.sync();
    return {
      ok: true,
      embedded: true,
      partCount: 1,
      namespaceUri: OPEN_WORKBOOK_CUSTOM_XML_NAMESPACE
    };
  });
}

export async function readWorkbookEmbeddedLocalConfig(workbookId: string): Promise<{
  ok: boolean;
  workbookId: string;
  embedded: boolean;
  partCount: number;
  config?: WorkbookLocalConfig;
  namespaceUri: string;
  error?: unknown;
}> {
  return Excel.run(async (context) => {
    const customXmlParts = getCustomXmlParts(context);
    if (!customXmlParts) {
      return {
        ok: false,
        workbookId,
        embedded: false,
        partCount: 0,
        namespaceUri: OPEN_WORKBOOK_CUSTOM_XML_NAMESPACE,
        error: runtimeError("CAPABILITY_UNAVAILABLE", "This Excel host does not expose workbook custom XML parts to Office.js.", { retryable: false })
      };
    }
    const parts = customXmlParts.getByNamespace(OPEN_WORKBOOK_CUSTOM_XML_NAMESPACE);
    parts.load("items/id");
    await context.sync();
    if (!parts.items || parts.items.length === 0) {
      return {
        ok: true,
        workbookId,
        embedded: false,
        partCount: 0,
        namespaceUri: OPEN_WORKBOOK_CUSTOM_XML_NAMESPACE
      };
    }
    const xmlResult = parts.items[0]!.getXml();
    await context.sync();
    const config = parseWorkbookLocalConfigXml(xmlResult.value);
    return {
      ok: true,
      workbookId,
      embedded: true,
      partCount: parts.items.length,
      config,
      namespaceUri: OPEN_WORKBOOK_CUSTOM_XML_NAMESPACE
    };
  });
}

export async function listPivotTables(workbookId: string): Promise<{ ok: boolean; pivotTables: PivotTableInfo[] }> {
  return Excel.run(async (context) => {
    const pivots = context.workbook.pivotTables;
    pivots.load("items/name");
    await context.sync();

    const pivotTables: PivotTableInfo[] = [];
    for (const pivot of pivots.items) {
      pivotTables.push(await readPivotTableInfo(context, workbookId, pivot));
    }
    return { ok: true, pivotTables };
  });
}

export async function getPivotTableInfo(request: PivotSelector): Promise<{ ok: boolean; info?: PivotTableInfo }> {
  return Excel.run(async (context) => {
    const pivot = context.workbook.pivotTables.getItemOrNullObject(request.pivotTableName);
    pivot.load("isNullObject");
    await context.sync();
    if (pivot.isNullObject) {
      return { ok: false };
    }
    return { ok: true, info: await readPivotTableInfo(context, request.workbookId, pivot) };
  });
}

export async function createPivotTable(request: PivotCreateRequest): Promise<{ ok: boolean; info: PivotTableInfo }> {
  return Excel.run(async (context) => {
    const source = request.sourceTableName
      ? context.workbook.tables.getItem(request.sourceTableName)
      : context.workbook.worksheets.getItem(request.sourceSheetName ?? request.destinationSheetName).getRange(stripSheetName(request.sourceAddress ?? ""));
    const destination = context.workbook.worksheets.getItem(request.destinationSheetName).getRange(stripSheetName(request.destinationAddress));
    const pivot = context.workbook.pivotTables.add(request.pivotTableName, source, destination);
    await context.sync();
    if (hasPivotCreateLayout(request)) {
      loadPivotTemplateReplayObjects(pivot);
      await context.sync();
      applyPivotCreateLayout(request, pivot);
      await context.sync();
      if (request.refresh !== false) {
        pivot.refresh();
        await context.sync();
      }
    }
    return { ok: true, info: await readPivotTableInfo(context, request.workbookId, pivot) };
  });
}

export async function refreshPivotTable(request: PivotSelector): Promise<{ ok: boolean; info?: PivotTableInfo }> {
  return Excel.run(async (context) => {
    const pivot = context.workbook.pivotTables.getItemOrNullObject(request.pivotTableName);
    pivot.load("isNullObject");
    await context.sync();
    if (pivot.isNullObject) {
      return { ok: false };
    }
    pivot.refresh();
    await context.sync();
    return { ok: true, info: await readPivotTableInfo(context, request.workbookId, pivot) };
  });
}

export async function refreshAllPivotTables(workbookId: string): Promise<{ ok: boolean }> {
  return Excel.run(async (context) => {
    context.workbook.pivotTables.refreshAll();
    await context.sync();
    return { ok: true };
  });
}

export async function copyPivotTableFromTemplate(request: PivotCopyFromTemplateRequest): Promise<PivotCopyFromTemplateResponse> {
  return Excel.run(async (context) => {
    const sourcePivot = context.workbook.pivotTables.getItemOrNullObject(request.templatePivotTableName);
    const targetPivot = context.workbook.pivotTables.getItemOrNullObject(request.pivotTableName);
    sourcePivot.load("isNullObject");
    targetPivot.load("isNullObject");
    await context.sync();
    if (sourcePivot.isNullObject || targetPivot.isNullObject) {
      return { ok: false, copied: [] };
    }

    const source = await readPivotTableInfo(context, request.workbookId, sourcePivot);
    loadPivotTemplateReplayObjects(targetPivot);
    await context.sync();
    const copied = applyPivotTemplateMetadata(source, targetPivot, request.dimensions);
    await context.sync();
    if (!request.dimensions || request.dimensions.includes("refresh")) {
      targetPivot.refresh();
      await context.sync();
    }
    const target = await readPivotTableInfo(context, request.workbookId, targetPivot);
    return {
      ok: true,
      copied,
      source,
      target
    };
  });
}

export async function deletePivotTable(request: PivotSelector): Promise<{ ok: boolean; deleted: boolean }> {
  return Excel.run(async (context) => {
    const pivot = context.workbook.pivotTables.getItemOrNullObject(request.pivotTableName);
    pivot.load("name,isNullObject");
    await context.sync();
    if (pivot.isNullObject) {
      return { ok: false, deleted: false };
    }
    pivot.delete();
    await context.sync();
    return { ok: true, deleted: true };
  });
}

export async function listCharts(workbookId: string): Promise<{ ok: boolean; charts: ChartInfo[] }> {
  return Excel.run(async (context) => {
    context.workbook.worksheets.load("items/name");
    await context.sync();

    for (const worksheet of context.workbook.worksheets.items) {
      worksheet.charts.load("items/name,items/id,items/chartType,items/top,items/left,items/width,items/height,items/style,items/plotBy");
    }
    await context.sync();
    const loaded = context.workbook.worksheets.items.flatMap((worksheet) =>
      worksheet.charts.items.map((chart) => loadChartInfoObjects(workbookId, worksheet.name, chart))
    );
    return { ok: true, charts: loaded.map(materializeChartInfo) };
  });
}

export async function getChartInfo(request: ChartSelector): Promise<{ ok: boolean; info?: ChartInfo }> {
  return Excel.run(async (context) => {
    const chart = context.workbook.worksheets.getItem(request.sheetName).charts.getItemOrNullObject(request.chartName);
    chart.load("name,id,chartType,top,left,width,height,style,plotBy,isNullObject");
    const loaded = loadChartInfoObjects(request.workbookId, request.sheetName, chart);
    await context.sync();
    if (chart.isNullObject) {
      return { ok: false };
    }
    return { ok: true, info: materializeChartInfo(loaded) };
  });
}

export async function createChart(request: ChartCreateRequest): Promise<{ ok: boolean; info: ChartInfo }> {
  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(request.sheetName);
    const source = worksheet.getRange(stripSheetName(request.sourceAddress));
    const chart = worksheet.charts.add(request.chartType as Excel.ChartType, source, request.seriesBy as Excel.ChartSeriesBy | undefined);
    if (request.chartName !== undefined) {
      chart.name = request.chartName;
    }
    if (request.title !== undefined) {
      chart.title.text = request.title;
    }
    if (request.style !== undefined) {
      chart.style = request.style;
    }
    if (request.position !== undefined) {
      chart.setPosition(request.position.startCell, request.position.endCell);
    }
    chart.load("name,id,chartType,top,left,width,height,style,plotBy");
    const loaded = loadChartInfoObjects(request.workbookId, request.sheetName, chart);
    await context.sync();
    return { ok: true, info: materializeChartInfo(loaded) };
  });
}

export async function updateChartDataSource(request: ChartUpdateDataSourceRequest): Promise<{ ok: boolean; info?: ChartInfo }> {
  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(request.sheetName);
    const chart = worksheet.charts.getItemOrNullObject(request.chartName);
    chart.load("name,id,chartType,top,left,width,height,style,plotBy,isNullObject");
    const loaded = loadChartInfoObjects(request.workbookId, request.sheetName, chart);
    await context.sync();
    if (chart.isNullObject) {
      return { ok: false };
    }
    chart.setData(worksheet.getRange(stripSheetName(request.sourceAddress)), request.seriesBy);
    await context.sync();
    return { ok: true, info: materializeChartInfo(loaded) };
  });
}

export async function copyChartFromTemplate(request: ChartSelector & { templateSheetName: string; templateChartName: string }): Promise<{
  ok: boolean;
  copied: string[];
  source?: ChartInfo;
  target?: ChartInfo;
}> {
  return Excel.run(async (context) => {
    const sourceChart = context.workbook.worksheets.getItem(request.templateSheetName).charts.getItemOrNullObject(request.templateChartName);
    const targetChart = context.workbook.worksheets.getItem(request.sheetName).charts.getItemOrNullObject(request.chartName);
    sourceChart.load("name,id,chartType,top,left,width,height,style,plotBy,isNullObject");
    sourceChart.title.load("text");
    targetChart.load("name,id,chartType,top,left,width,height,style,plotBy,isNullObject");
    targetChart.title.load("text");
    await context.sync();
    if (sourceChart.isNullObject || targetChart.isNullObject) {
      return { ok: false, copied: [] };
    }

    const copied: string[] = [];
    targetChart.chartType = sourceChart.chartType;
    copied.push("chartType");
    if (sourceChart.style !== undefined) {
      targetChart.style = sourceChart.style;
      copied.push("style");
    }
    if (sourceChart.title.text !== undefined) {
      targetChart.title.text = sourceChart.title.text;
      copied.push("title");
    }
    targetChart.top = sourceChart.top;
    targetChart.left = sourceChart.left;
    targetChart.width = sourceChart.width;
    targetChart.height = sourceChart.height;
    copied.push("position");

    targetChart.load("name,id,chartType,top,left,width,height,style,plotBy");
    targetChart.title.load("text");
    await context.sync();

    return {
      ok: true,
      copied,
      source: materializeChartInfo({ workbookId: request.workbookId, sheetName: request.templateSheetName, chart: sourceChart }),
      target: materializeChartInfo({ workbookId: request.workbookId, sheetName: request.sheetName, chart: targetChart })
    };
  });
}

export async function refreshChart(request: ChartSelector): Promise<{ ok: boolean; info?: ChartInfo }> {
  return getChartInfo(request);
}

export async function deleteChart(request: ChartSelector): Promise<{ ok: boolean; deleted: boolean }> {
  return Excel.run(async (context) => {
    const chart = context.workbook.worksheets.getItem(request.sheetName).charts.getItemOrNullObject(request.chartName);
    chart.load("name,isNullObject");
    await context.sync();
    if (chart.isNullObject) {
      return { ok: false, deleted: false };
    }
    chart.delete();
    await context.sync();
    return { ok: true, deleted: true };
  });
}

export async function listTables(workbookId: string): Promise<{ ok: boolean; tables: TableInfo[] }> {
  return Excel.run(async (context) => {
    const tables = context.workbook.tables;
    tables.load("items/name,items/id,items/style,items/showHeaders,items/showTotals,items/showFilterButton,items/showBandedRows,items/showBandedColumns");
    await context.sync();

    const loaded = tables.items.map((table) => loadTableInfoObjects(table));
    await context.sync();

    return {
      ok: true,
      tables: loaded.map((loadedTable) => materializeTableInfo(workbookId, loadedTable))
    };
  });
}

export async function getTableInfo(request: TableSelector): Promise<{ ok: boolean; info: TableInfo }> {
  return Excel.run(async (context) => {
    const table = context.workbook.tables.getItem(request.tableName);
    const loaded = loadTableInfoObjects(table);
    await context.sync();
    return { ok: true, info: materializeTableInfo(request.workbookId, loaded) };
  });
}

export async function readTable(request: TableReadRequest): Promise<{ ok: boolean; table: TableReadResponse }> {
  return Excel.run(async (context) => {
    const table = context.workbook.tables.getItem(request.tableName);
    const loaded = loadTableInfoObjects(table);
    await context.sync();

    const headerRange = table.getHeaderRowRange();
    headerRange.load("values");
    const info = materializeTableInfo(request.workbookId, loaded);
    const projectedIndexes = resolveTableColumnIndexes(info, request.columns);
    const rowOffset = Math.max(0, request.rowOffset ?? 0);
    const availableRows = Math.max(0, table.rows.count - rowOffset);
    const requestedRowCount = request.rowLimit === undefined ? availableRows : Math.max(0, Math.min(request.rowLimit, availableRows));
    const bodyRange = table.rows.count > 0 ? table.getDataBodyRange() : undefined;
    const dataRange =
      bodyRange && requestedRowCount > 0 && projectedIndexes.length > 0
        ? bodyRange.getCell(rowOffset, 0).getResizedRange(requestedRowCount - 1, info.columnCount - 1)
        : undefined;
    const dataColumnRanges =
      dataRange && projectedIndexes.length !== info.columnCount
        ? projectedIndexes.map((columnIndex) => dataRange.getColumn(columnIndex))
        : [];
    const shouldLoadValues = request.includeValues ?? true;
    const shouldLoadFormulas = request.includeFormulas ?? true;
    const shouldLoadText = request.includeText ?? true;
    const shouldLoadNumberFormats = request.includeNumberFormats ?? true;
    if (dataRange) {
      const loadProperties = [
        shouldLoadValues ? "values" : undefined,
        shouldLoadFormulas ? "formulas" : undefined,
        shouldLoadText ? "text" : undefined,
        shouldLoadNumberFormats ? "numberFormat" : undefined
      ].filter((property): property is string => property !== undefined);
      if (loadProperties.length > 0) {
        if (dataColumnRanges.length > 0) {
          for (const columnRange of dataColumnRanges) {
            columnRange.load(loadProperties.join(","));
          }
        } else {
          dataRange.load(loadProperties.join(","));
        }
      }
    }
    await context.sync();

    const headers = projectMatrix(headerRange.values as TableReadResponse["headers"], projectedIndexes);
    const response: TableReadResponse = {
      info,
      headers,
      rowOffset,
      ...(request.rowLimit !== undefined ? { rowLimit: request.rowLimit } : {}),
      rowCount: requestedRowCount,
      truncated: rowOffset + requestedRowCount < table.rows.count,
      projectedColumns: projectedIndexes.map((index) => info.columns[index]!).filter(Boolean)
    };
    if (dataRange && shouldLoadValues) {
      response.values = materializeProjectedRangeFacet(dataRange, dataColumnRanges, "values", projectedIndexes) as NonNullable<TableReadResponse["values"]>;
    } else if (shouldLoadValues) {
      response.values = [];
    }
    if (dataRange && shouldLoadFormulas) {
      response.formulas = materializeProjectedRangeFacet(dataRange, dataColumnRanges, "formulas", projectedIndexes) as NonNullable<TableReadResponse["formulas"]>;
    } else if (shouldLoadFormulas) {
      response.formulas = [];
    }
    if (dataRange && shouldLoadText) {
      response.text = materializeProjectedRangeFacet(dataRange, dataColumnRanges, "text", projectedIndexes) as NonNullable<TableReadResponse["text"]>;
    } else if (shouldLoadText) {
      response.text = [];
    }
    if (dataRange && shouldLoadNumberFormats) {
      response.numberFormat = materializeProjectedRangeFacet(dataRange, dataColumnRanges, "numberFormat", projectedIndexes) as NonNullable<TableReadResponse["numberFormat"]>;
    } else if (shouldLoadNumberFormats) {
      response.numberFormat = [];
    }

    return {
      ok: true,
      table: response
    };
  });
}

export async function createTable(request: TableCreateRequest): Promise<{ ok: boolean; info: TableInfo }> {
  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(request.sheetName);
    const range = worksheet.getRange(stripSheetName(request.address));
    if (request.values) {
      range.values = request.values;
    }
    const table = worksheet.tables.add(range, request.hasHeaders);
    if (request.tableName) {
      table.name = request.tableName;
    }
    if (request.style) {
      table.style = request.style;
    }
    if (request.showTotals !== undefined) {
      table.showTotals = request.showTotals;
    }
    const loaded = loadTableInfoObjects(table);
    await context.sync();
    return { ok: true, info: materializeTableInfo(request.workbookId, loaded) };
  });
}

export async function resizeTable(request: TableResizeRequest): Promise<{ ok: boolean; info: TableInfo }> {
  return mutateTableAndReturnInfo(request, (table) => table.resize(request.address));
}

export async function reorderTableColumns(request: TableReorderColumnsRequest): Promise<{ ok: boolean; info: TableInfo; warnings: OperationWarning[] }> {
  return Excel.run(async (context) => {
    const table = context.workbook.tables.getItem(request.tableName);
    const loaded = loadTableInfoObjects(table);
    await context.sync();

    const info = materializeTableInfo(request.workbookId, loaded);
    const sourceIndexes = resolveTableColumnIndexes(info, request.columnOrder);
    if (sourceIndexes.length !== info.columnCount) {
      return {
        ok: false,
        info,
        warnings: [
          {
            code: "TABLE_COLUMN_ORDER_INVALID",
            message: "Column order must resolve to every table column exactly once.",
            details: { columnOrder: request.columnOrder, resolvedIndexes: sourceIndexes }
          }
        ]
      };
    }

    const tableRange = table.getRange();
    tableRange.load("rowCount,columnCount");
    await context.sync();

    const scratchSheet = context.workbook.worksheets.add(`__owb_reorder_${Date.now().toString(36)}`);
    const originalRange = scratchSheet.getRangeByIndexes(0, 0, tableRange.rowCount, tableRange.columnCount);
    const reorderedRange = scratchSheet.getRangeByIndexes(0, tableRange.columnCount + 1, tableRange.rowCount, tableRange.columnCount);
    originalRange.copyFrom(tableRange, Excel.RangeCopyType.all);
    for (const [targetIndex, sourceIndex] of sourceIndexes.entries()) {
      reorderedRange.getColumn(targetIndex).copyFrom(originalRange.getColumn(sourceIndex), Excel.RangeCopyType.all);
    }
    tableRange.copyFrom(reorderedRange, Excel.RangeCopyType.all);
    scratchSheet.delete();

    const reloaded = loadTableInfoObjects(table);
    await context.sync();
    return { ok: true, info: materializeTableInfo(request.workbookId, reloaded), warnings: [] };
  });
}

export async function appendTableRows(request: TableAppendRowsRequest): Promise<{ ok: boolean; info: TableInfo }> {
  return mutateTableAndReturnInfo(request, (table) => {
    table.rows.add(request.index ?? -1, request.values as any, request.alwaysInsert ?? true);
  });
}

export async function updateTableRows(request: TableUpdateRowsRequest): Promise<{ ok: boolean; info: TableInfo }> {
  return mutateTableAndReturnInfo(request, (table) => {
    for (const row of request.rows) {
      table.rows.getItemAt(row.index).values = [row.values] as any;
    }
  });
}

export async function clearTableDataKeepFormulas(request: TableSelector): Promise<{ ok: boolean; info: TableInfo }> {
  return Excel.run(async (context) => {
    const table = context.workbook.tables.getItem(request.tableName);
    table.rows.load("count");
    await context.sync();
    if (table.rows.count > 0) {
      const range = table.getDataBodyRange();
      range.load("formulas");
      await context.sync();
      range.formulas = (range.formulas as string[][]).map((row) =>
        row.map((formula) => (typeof formula === "string" && formula.startsWith("=") ? formula : null))
      );
    }
    const loaded = loadTableInfoObjects(table);
    await context.sync();
    return { ok: true, info: materializeTableInfo(request.workbookId, loaded) };
  });
}

export async function clearTableFilters(request: TableSelector): Promise<{ ok: boolean; info: TableInfo }> {
  return mutateTableAndReturnInfo(request, (table) => table.clearFilters());
}

export async function applyTableFilters(request: TableApplyFiltersRequest): Promise<{ ok: boolean; info: TableInfo }> {
  return mutateTableAndReturnInfo(request, (table) => {
    for (const filter of request.filters) {
      table.columns.getItem(filter.column as any).filter.apply(filter.criteria as Excel.FilterCriteria);
    }
  });
}

export async function sortTable(request: TableSortRequest): Promise<{ ok: boolean; info: TableInfo }> {
  return mutateTableAndReturnInfo(request, (table) => {
    table.sort.apply(request.fields as Excel.SortField[], request.matchCase, request.method as any);
  });
}

export async function clearTableSort(request: TableSelector): Promise<{ ok: boolean; info: TableInfo }> {
  return mutateTableAndReturnInfo(request, (table) => table.sort.clear());
}

export async function setTableTotalRow(request: TableSetTotalRowRequest): Promise<{ ok: boolean; info: TableInfo }> {
  return mutateTableAndReturnInfo(request, (table) => {
    table.showTotals = request.showTotals;
  });
}

export async function setTableStyle(request: TableSetStyleRequest): Promise<{ ok: boolean; info: TableInfo }> {
  return mutateTableAndReturnInfo(request, (table) => {
    table.style = request.style;
  });
}

export async function copyTableStructure(request: TableCopyStructureRequest): Promise<{ ok: boolean; info: TableInfo }> {
  return Excel.run(async (context) => {
    const sourceTable = context.workbook.tables.getItem(request.tableName);
    const headerRange = sourceTable.getHeaderRowRange();
    headerRange.load("values");
    sourceTable.load("name,style,showTotals,showFilterButton,showBandedRows,showBandedColumns");
    await context.sync();

    const targetSheet = context.workbook.worksheets.getItem(request.targetSheetName);
    const targetRange = targetSheet.getRange(stripSheetName(request.targetAddress));
    targetRange.values = headerRange.values;
    const table = targetSheet.tables.add(targetRange, true);
    table.name = request.newTableName ?? `${sourceTable.name}_Copy`;
    if (request.includeStyle ?? true) {
      table.style = sourceTable.style;
      table.showBandedRows = sourceTable.showBandedRows;
      table.showBandedColumns = sourceTable.showBandedColumns;
      table.showFilterButton = sourceTable.showFilterButton;
    }
    if (request.includeTotals ?? false) {
      table.showTotals = sourceTable.showTotals;
    }
    const loaded = loadTableInfoObjects(table);
    await context.sync();
    return { ok: true, info: materializeTableInfo(request.workbookId, loaded) };
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

export async function readRangeHyperlinks(request: RangeMetadataRequest): Promise<RangeMetadataResponse> {
  return readRangeMetadata(request, (range) => {
    range.load("address,hyperlink");
  }, (range) => range.hyperlink);
}

export async function readRangeDataValidation(request: RangeMetadataRequest): Promise<RangeMetadataResponse> {
  return readRangeMetadata(request, (range) => {
    range.load("address");
    range.dataValidation.load("type,rule,prompt,errorAlert,ignoreBlanks,valid");
  }, (range) => range.dataValidation.toJSON());
}

export async function readRangeConditionalFormatting(request: RangeMetadataRequest): Promise<RangeMetadataResponse> {
  return readRangeMetadata(request, (range) => {
    range.load("address");
    range.conditionalFormats.load("items/id,items/type,items/priority,items/stopIfTrue");
  }, (range) => range.conditionalFormats.toJSON());
}

export async function readRangeMergedCells(request: RangeMetadataRequest): Promise<RangeMetadataResponse> {
  return Excel.run(async (context) => {
    const target = targetFromMetadataRequest(request);
    const range = getRange(context, target);
    const merged = range.getMergedAreasOrNullObject();
    merged.load("address,areaCount,cellCount,isNullObject");
    await context.sync();
    return {
      ok: true,
      target,
      data: summarizeRangeAreas(merged),
      warnings: []
    };
  });
}

export async function readRangeComments(request: RangeMetadataRequest): Promise<RangeMetadataResponse> {
  return unsupportedRangeMetadata(request, "RANGE_COMMENTS_UNSUPPORTED", "Office.js comment-to-range mapping is not enabled in this executor yet.");
}

export async function readRangeNotes(request: RangeMetadataRequest): Promise<RangeMetadataResponse> {
  return unsupportedRangeMetadata(request, "RANGE_NOTES_UNSUPPORTED", "Legacy notes are not exposed through this executor yet.");
}

export async function searchRange(request: RangeSearchRequest): Promise<RangeSearchResponse> {
  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(request.sheetName);
    const criteria: Excel.WorksheetSearchCriteria = {};
    if (request.completeMatch !== undefined) {
      criteria.completeMatch = request.completeMatch;
    }
    if (request.matchCase !== undefined) {
      criteria.matchCase = request.matchCase;
    }
    const matches = worksheet.findAllOrNullObject(request.text, criteria);
    matches.load("address,areaCount,cellCount,isNullObject");
    await context.sync();
    return {
      ok: true,
      matches: summarizeRangeAreas(matches)
    };
  });
}

export async function findBlankCells(request: RangeMetadataRequest): Promise<RangeMetadataResponse> {
  return readSpecialCells(request, "Blanks");
}

export async function findFormulaErrors(request: RangeMetadataRequest): Promise<RangeMetadataResponse> {
  return readSpecialCells(request, "Formulas", "Errors");
}

export async function executeBatch(payload: AddinExecuteBatchRequest): Promise<OperationResult> {
  const started = performance.now();
  const counters: ExecutionCounters = {
    syncCount: 0,
    cellsRead: 0,
    cellsWritten: 0,
    rangeCount: payload.compiled.targetFingerprints.length,
    chunkCount: 0
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
      const readOperations: Array<{ operation: ExcelOperation; range: Excel.Range; facets?: RangeSnapshotFacet[] }> = [];
      let formulasChanged = 0;
      let sheetsChanged = 0;

      maybeSuspendExcel(context, payload.compiled.estimatedCellsTouched);

      for (const operation of payload.request.operations) {
        switch (operation.kind) {
          case "range.read_full": {
            const range = getRange(context, operation.target);
            const facets = rangeSnapshotFacetsForOperation(operation);
            loadSnapshotProperties(range, facets);
            readOperations.push({ operation, range, facets });
            counters.cellsRead += payload.compiled.estimatedCellsTouched;
            break;
          }
          case "range.write_values": {
            assertMatrixShape(operation.target, operation.values);
            counters.chunkCount += writeMatrixInChunks(context, operation.target, operation.values, "values");
            counters.cellsWritten += matrixCellCount(operation.values);
            break;
          }
          case "range.write_formulas": {
            assertMatrixShape(operation.target, operation.formulas);
            counters.chunkCount += writeMatrixInChunks(context, operation.target, operation.formulas, "formulas");
            counters.cellsWritten += matrixCellCount(operation.formulas);
            formulasChanged += matrixCellCount(operation.formulas);
            break;
          }
          case "range.write_number_formats": {
            assertMatrixShape(operation.target, operation.numberFormat);
            counters.chunkCount += writeMatrixInChunks(context, operation.target, operation.numberFormat, "numberFormat");
            counters.cellsWritten += matrixCellCount(operation.numberFormat);
            break;
          }
          case "range.write_styles": {
            applyRangeStyle(getRange(context, operation.target), operation.style);
            counters.cellsWritten += payload.compiled.estimatedCellsTouched;
            break;
          }
          case "range.write_hyperlinks":
          case "range.write_comments": {
            warnings.push({
              code: "OPERATION_NOT_SUPPORTED",
              message: `${operation.kind} is defined in the protocol but is not enabled in the Office.js executor yet.`,
              target: operation.target
            });
            break;
          }
          case "range.clear_values_keep_format": {
            getRange(context, operation.target).clear(Excel.ClearApplyTo.contents);
            counters.cellsWritten += payload.compiled.estimatedCellsTouched;
            break;
          }
          case "range.clear":
          case "range.clear_values":
          case "range.clear_formats": {
            const applyTo =
              operation.kind === "range.clear_values"
                ? Excel.ClearApplyTo.contents
                : operation.kind === "range.clear_formats"
                  ? Excel.ClearApplyTo.formats
                  : toClearApplyTo(operation.applyTo ?? "all");
            getRange(context, operation.target).clear(applyTo);
            counters.cellsWritten += payload.compiled.estimatedCellsTouched;
            break;
          }
          case "range.copy": {
            getRange(context, operation.target).copyFrom(
              getRange(context, operation.source),
              toRangeCopyType(operation.copyType ?? "all")
            );
            counters.cellsWritten += payload.compiled.estimatedCellsTouched;
            break;
          }
          case "range.move": {
            const source = getRange(context, operation.source);
            const target = getRange(context, operation.target);
            target.copyFrom(source, Excel.RangeCopyType.all);
            source.clear(Excel.ClearApplyTo.all);
            counters.cellsWritten += payload.compiled.estimatedCellsTouched;
            break;
          }
          case "range.insert_rows":
          case "range.insert_columns": {
            getRange(context, operation.target).insert(
              operation.kind === "range.insert_rows" ? Excel.InsertShiftDirection.down : Excel.InsertShiftDirection.right
            );
            counters.cellsWritten += payload.compiled.estimatedCellsTouched;
            break;
          }
          case "range.delete_rows":
          case "range.delete_columns": {
            getRange(context, operation.target).delete(
              operation.kind === "range.delete_rows" ? Excel.DeleteShiftDirection.up : Excel.DeleteShiftDirection.left
            );
            counters.cellsWritten += payload.compiled.estimatedCellsTouched;
            break;
          }
          case "range.autofit_columns": {
            getRange(context, operation.target).format.autofitColumns();
            break;
          }
          case "range.autofit_rows": {
            getRange(context, operation.target).format.autofitRows();
            break;
          }
          case "range.merge": {
            getRange(context, operation.target).merge(operation.across ?? false);
            break;
          }
          case "range.unmerge": {
            getRange(context, operation.target).unmerge();
            break;
          }
          case "range.restore_snapshot": {
            restoreRangeSnapshot(context, operation);
            counters.cellsWritten += payload.compiled.estimatedCellsTouched;
            break;
          }
          case "workbook.calculate": {
            context.workbook.application.calculate(
              operation.calculationType === "recalculate" ? Excel.CalculationType.recalculate : Excel.CalculationType.full
            );
            break;
          }
          case "workbook.save": {
            context.workbook.save(Excel.SaveBehavior.save);
            break;
          }
          case "sheet.create": {
            const worksheet = context.workbook.worksheets.add(operation.sheetName);
            if (operation.activate ?? true) {
              worksheet.activate();
            }
            sheetsChanged += 1;
            break;
          }
          case "sheet.copy": {
            const sourceSheet = context.workbook.worksheets.getItem(operation.sourceSheetName);
            const relativeSheet = operation.relativeToSheetName
              ? context.workbook.worksheets.getItem(operation.relativeToSheetName)
              : sourceSheet;
            const copiedSheet = sourceSheet.copy(toWorksheetPositionType(operation.position ?? "after"), relativeSheet);
            copiedSheet.name = operation.newSheetName;
            if (operation.activate ?? true) {
              copiedSheet.activate();
            }
            sheetsChanged += 1;
            break;
          }
          case "sheet.rename": {
            context.workbook.worksheets.getItem(operation.sheetName).name = operation.newSheetName;
            sheetsChanged += 1;
            break;
          }
          case "sheet.delete": {
            context.workbook.worksheets.getItem(operation.sheetName).delete();
            sheetsChanged += 1;
            break;
          }
          case "sheet.move": {
            warnings.push({
              code: "OPERATION_NOT_SUPPORTED",
              message: "sheet.move is defined in the protocol but is not enabled in the Office.js executor yet."
            });
            break;
          }
          case "sheet.hide": {
            context.workbook.worksheets.getItem(operation.sheetName).visibility = Excel.SheetVisibility.hidden;
            sheetsChanged += 1;
            break;
          }
          case "sheet.unhide": {
            context.workbook.worksheets.getItem(operation.sheetName).visibility = Excel.SheetVisibility.visible;
            sheetsChanged += 1;
            break;
          }
          case "sheet.protect": {
            context.workbook.worksheets.getItem(operation.sheetName).protection.protect(undefined, operation.password);
            sheetsChanged += 1;
            break;
          }
          case "sheet.unprotect": {
            context.workbook.worksheets.getItem(operation.sheetName).protection.unprotect(operation.password);
            sheetsChanged += 1;
            break;
          }
          case "sheet.clear": {
            const usedRange = context.workbook.worksheets.getItem(operation.sheetName).getUsedRangeOrNullObject();
            usedRange.clear(toClearApplyTo(operation.applyTo ?? "all"));
            sheetsChanged += 1;
            break;
          }
          case "sheet.set_tab_color": {
            context.workbook.worksheets.getItem(operation.sheetName).tabColor = operation.color;
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

      const readData = readOperations.map(({ operation, range, facets }) => ({
        operationId: operation.operationId,
        snapshot: materializeSnapshot("target" in operation ? operation.target : payload.compiled.targetFingerprints[0]!.range, range, facets)
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

export async function captureTemplate(request: TemplateCaptureRequest): Promise<TemplateCaptureResponse> {
  const captured = await captureSheetFingerprint({
    workbookId: request.workbookId,
    sheetName: request.sourceSheetName,
    dataRegions: request.dataRegions
  });
  return {
    sourceSheetName: captured.sourceSheetName,
    dataRegions: captured.dataRegions,
    fingerprintPayload: captured.fingerprintPayload
  };
}

export async function captureSheetFingerprint(request: SheetTemplateFingerprintRequest): Promise<SheetTemplateFingerprintResponse> {
  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(request.sheetName);
    const usedRange = worksheet.getUsedRangeOrNullObject();
    const tables = worksheet.tables;

    worksheet.load("name, position, visibility");
    usedRange.load("address, rowCount, columnCount, values, formulas, numberFormat");
    usedRange.format.load("rowHeight, columnWidth, horizontalAlignment, verticalAlignment");
    usedRange.format.fill.load("color");
    usedRange.format.font.load("name, size, color, bold, italic");
    tables.load("items/name");
    await context.sync();

    const rangePayload = usedRange.isNullObject
      ? null
      : {
          address: usedRange.address,
          rowCount: usedRange.rowCount,
          columnCount: usedRange.columnCount,
          values: usedRange.values,
          formulas: usedRange.formulas,
          numberFormat: usedRange.numberFormat
        };

    const stylePayload = usedRange.isNullObject
      ? null
      : {
          fillColor: usedRange.format.fill.color,
          fontName: usedRange.format.font.name,
          fontSize: usedRange.format.font.size,
          fontColor: usedRange.format.font.color,
          fontBold: usedRange.format.font.bold,
          fontItalic: usedRange.format.font.italic,
          horizontalAlignment: usedRange.format.horizontalAlignment,
          verticalAlignment: usedRange.format.verticalAlignment,
          rowHeight: usedRange.format.rowHeight,
          columnWidth: usedRange.format.columnWidth
        };

    return {
      sheetName: worksheet.name,
      sourceSheetName: worksheet.name,
      dataRegions: request.dataRegions ?? [],
      fingerprintPayload: {
        structure: {
          sheetName: worksheet.name,
          position: worksheet.position,
          visibility: worksheet.visibility,
          usedRange: rangePayload
            ? {
                address: rangePayload.address,
                rowCount: rangePayload.rowCount,
                columnCount: rangePayload.columnCount
              }
            : null,
          dataRegions: request.dataRegions ?? []
        },
        formulas: rangePayload?.formulas ?? null,
        styles: {
          usedRange: stylePayload,
          numberFormat: rangePayload?.numberFormat ?? null
        },
        filters: {
          note: "Filter capture will be expanded with table/filter-specific APIs."
        },
        tables: tables.items.map((table) => ({ name: table.name })),
        printLayout: {
          note: "Print layout capture will be expanded with page layout APIs."
        }
      }
    };
  });
}

export async function captureStyleFingerprint(request: StyleFingerprintRequest): Promise<StyleFingerprintResponse> {
  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(request.sheetName);
    const range = request.address ? worksheet.getRange(stripSheetName(request.address)) : worksheet.getUsedRangeOrNullObject();
    range.load("address, rowCount, columnCount, numberFormat");
    range.format.load("rowHeight, columnWidth, horizontalAlignment, verticalAlignment, wrapText");
    range.format.fill.load("color");
    range.format.font.load("name, size, color, bold, italic, underline");
    range.format.borders.load("items/sideIndex,items/style,items/color,items/weight");
    await context.sync();

    if ("isNullObject" in range && range.isNullObject) {
      return {
        workbookId: request.workbookId,
        sheetName: request.sheetName,
        address: "",
        capturedAt: new Date().toISOString(),
        rowCount: 0,
        columnCount: 0,
        truncated: false,
        dimensions: {},
        warnings: [
          {
            code: "EMPTY_USED_RANGE",
            message: `Sheet ${request.sheetName} has no used range to fingerprint.`
          }
        ]
      };
    }

    const columnRanges: Excel.Range[] = [];
    const rowRanges: Excel.Range[] = [];
    for (let columnIndex = 0; columnIndex < range.columnCount; columnIndex += 1) {
      const column = range.getColumn(columnIndex);
      column.format.load("columnWidth");
      columnRanges.push(column);
    }
    for (let rowIndex = 0; rowIndex < range.rowCount; rowIndex += 1) {
      const row = range.getRow(rowIndex);
      row.format.load("rowHeight");
      rowRanges.push(row);
    }

    const sampleLimit = request.maxCellSamples ?? 500;
    const cellCount = range.rowCount * range.columnCount;
    const sampledCells: Array<{ rowIndex: number; columnIndex: number; cell: Excel.Range }> = [];
    if (cellCount <= sampleLimit) {
      for (let rowIndex = 0; rowIndex < range.rowCount; rowIndex += 1) {
        for (let columnIndex = 0; columnIndex < range.columnCount; columnIndex += 1) {
          const cell = range.getCell(rowIndex, columnIndex);
          cell.load("numberFormat");
          cell.format.load("horizontalAlignment, verticalAlignment, wrapText");
          cell.format.fill.load("color");
          cell.format.font.load("name, size, color, bold, italic, underline");
          sampledCells.push({ rowIndex, columnIndex, cell });
        }
      }
    }

    await context.sync();

    const cellStyles = sampledCells.map(({ rowIndex, columnIndex, cell }) => ({
      rowIndex,
      columnIndex,
      fillColor: optionalValue(cell.format.fill.color),
      fontName: optionalValue(cell.format.font.name),
      fontSize: optionalValue(cell.format.font.size),
      fontColor: optionalValue(cell.format.font.color),
      fontBold: optionalValue(cell.format.font.bold),
      fontItalic: optionalValue(cell.format.font.italic),
      fontUnderline: optionalValue(String(cell.format.font.underline)),
      horizontalAlignment: optionalValue(String(cell.format.horizontalAlignment)),
      verticalAlignment: optionalValue(String(cell.format.verticalAlignment)),
      wrapText: optionalValue((cell.format as unknown as { wrapText?: boolean }).wrapText),
      numberFormat: (cell.numberFormat as string[][])[0]?.[0]
    }));

    const warnings: OperationWarning[] = [];
    if (cellCount > sampleLimit) {
      warnings.push({
        code: "STYLE_SAMPLE_TRUNCATED",
        message: `Cell-level style sampling skipped for ${cellCount} cells; increase maxCellSamples or pass a smaller address.`,
        details: { cellCount, sampleLimit }
      });
    }

    return {
      workbookId: request.workbookId,
      sheetName: request.sheetName,
      address: stripSheetName(range.address),
      capturedAt: new Date().toISOString(),
      rowCount: range.rowCount,
      columnCount: range.columnCount,
      truncated: cellCount > sampleLimit,
      dimensions: {
        columnWidths: columnRanges.map((column, columnIndex) => ({ columnIndex, width: optionalValue(column.format.columnWidth) })),
        rowHeights: rowRanges.map((row, rowIndex) => ({ rowIndex, height: optionalValue(row.format.rowHeight) })),
        fills: {
          rangeColor: optionalValue(range.format.fill.color),
          cells: cellStyles.map(({ rowIndex, columnIndex, fillColor }) => ({ rowIndex, columnIndex, fillColor }))
        },
        fonts: {
          range: {
            name: optionalValue(range.format.font.name),
            size: optionalValue(range.format.font.size),
            color: optionalValue(range.format.font.color),
            bold: optionalValue(range.format.font.bold),
            italic: optionalValue(range.format.font.italic),
            underline: optionalValue(String(range.format.font.underline))
          },
          cells: cellStyles.map(({ rowIndex, columnIndex, fontName, fontSize, fontColor, fontBold, fontItalic, fontUnderline }) => ({
            rowIndex,
            columnIndex,
            fontName,
            fontSize,
            fontColor,
            fontBold,
            fontItalic,
            fontUnderline
          }))
        },
        alignment: {
          range: {
            horizontalAlignment: optionalValue(String(range.format.horizontalAlignment)),
            verticalAlignment: optionalValue(String(range.format.verticalAlignment)),
            wrapText: optionalValue((range.format as unknown as { wrapText?: boolean }).wrapText)
          },
          cells: cellStyles.map(({ rowIndex, columnIndex, horizontalAlignment, verticalAlignment, wrapText }) => ({
            rowIndex,
            columnIndex,
            horizontalAlignment,
            verticalAlignment,
            wrapText
          }))
        },
        numberFormats: {
          matrix: range.numberFormat,
          cells: cellStyles.map(({ rowIndex, columnIndex, numberFormat }) => ({ rowIndex, columnIndex, numberFormat }))
        },
        borders: (range.format.borders as unknown as { items?: unknown[] }).items ?? [],
        conditionalFormatting: {
          note: "Use excel.range.read_conditional_formatting for detailed rule inspection."
        },
        dataValidation: {
          note: "Use excel.range.read_data_validation for detailed rule inspection."
        },
        freezePanes: {
          note: "Office.js freeze pane capture is tracked as a layout capability."
        },
        printSettings: {
          note: "Office.js print setting capture is tracked as a layout capability."
        },
        pageLayout: {
          note: "Office.js page layout capture is tracked as a layout capability."
        },
        hiddenRowsColumns: {
          note: "Hidden row/column capture is tracked as a layout capability."
        }
      },
      warnings
    };
  });
}

export async function copyStyleDimensions(request: StyleCopyRequest): Promise<StyleCopyResponse> {
  return Excel.run(async (context) => {
    const sourceSheet = context.workbook.worksheets.getItem(request.sourceSheetName);
    const targetSheet = context.workbook.worksheets.getItem(request.targetSheetName);
    const sourceRange = request.sourceAddress ? sourceSheet.getRange(stripSheetName(request.sourceAddress)) : sourceSheet.getUsedRangeOrNullObject();
    sourceRange.load("address,rowCount,columnCount");
    await context.sync();

    if ("isNullObject" in sourceRange && sourceRange.isNullObject) {
      return {
        ok: false,
        copied: [],
        warnings: [
          {
            code: "EMPTY_SOURCE_RANGE",
            message: `Sheet ${request.sourceSheetName} has no used range to copy styles from.`
          }
        ]
      };
    }

    const targetAddress = request.targetAddress ?? stripSheetName(sourceRange.address);
    const targetRange = targetSheet.getRange(stripSheetName(targetAddress));
    const copied: StyleDimension[] = [];
    const warnings: OperationWarning[] = [];
    const dimensions = new Set(request.dimensions);
    const copyAll = request.dimensions.length === 0;
    const formatDimensions: StyleDimension[] = ["borders", "fills", "fonts", "alignment", "numberFormats", "conditionalFormatting", "dataValidation"];

    if (copyAll || formatDimensions.some((dimension) => dimensions.has(dimension))) {
      targetRange.copyFrom(sourceRange, Excel.RangeCopyType.formats);
      copied.push(...formatDimensions.filter((dimension) => copyAll || dimensions.has(dimension)));
    }

    const copyColumnWidths = copyAll || dimensions.has("columnWidths");
    const copyRowHeights = copyAll || dimensions.has("rowHeights");
    const sourceColumns: Excel.Range[] = [];
    const targetColumns: Excel.Range[] = [];
    const sourceRows: Excel.Range[] = [];
    const targetRows: Excel.Range[] = [];

    if (copyColumnWidths) {
      for (let columnIndex = 0; columnIndex < sourceRange.columnCount; columnIndex += 1) {
        const sourceColumn = sourceRange.getColumn(columnIndex);
        sourceColumn.format.load("columnWidth");
        sourceColumns.push(sourceColumn);
        targetColumns.push(targetRange.getColumn(columnIndex));
      }
    }

    if (copyRowHeights) {
      for (let rowIndex = 0; rowIndex < sourceRange.rowCount; rowIndex += 1) {
        const sourceRow = sourceRange.getRow(rowIndex);
        sourceRow.format.load("rowHeight");
        sourceRows.push(sourceRow);
        targetRows.push(targetRange.getRow(rowIndex));
      }
    }

    if (copyColumnWidths || copyRowHeights) {
      await context.sync();
    }

    if (copyColumnWidths) {
      sourceColumns.forEach((sourceColumn, index) => {
        targetColumns[index]!.format.columnWidth = sourceColumn.format.columnWidth;
      });
      copied.push("columnWidths");
    }

    if (copyRowHeights) {
      sourceRows.forEach((sourceRow, index) => {
        targetRows[index]!.format.rowHeight = sourceRow.format.rowHeight;
      });
      copied.push("rowHeights");
    }

    for (const unsupported of ["freezePanes", "printSettings", "pageLayout", "hiddenRowsColumns"] as const) {
      if (copyAll || dimensions.has(unsupported)) {
        warnings.push({
          code: "STYLE_LAYOUT_DIMENSION_UNAVAILABLE",
          message: `${unsupported} is tracked in fingerprints but is not safely replayed through the current Office.js path.`,
          details: { dimension: unsupported }
        });
      }
    }

    await context.sync();
    return { ok: warnings.length === 0 || copied.length > 0, copied: [...new Set(copied)], warnings };
  });
}

export async function readFormulaPatterns(request: FormulaPatternRequest): Promise<FormulaPatternResponse> {
  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(request.sheetName);
    const range = request.address ? worksheet.getRange(stripSheetName(request.address)) : worksheet.getUsedRangeOrNullObject();
    range.load("address,rowCount,columnCount,formulas,formulasR1C1");
    await context.sync();

    if ("isNullObject" in range && range.isNullObject) {
      return {
        workbookId: request.workbookId,
        sheetName: request.sheetName,
        address: "",
        capturedAt: new Date().toISOString(),
        rowCount: 0,
        columnCount: 0,
        formulaCount: 0,
        formulas: [],
        patternMatrix: [],
        patterns: [],
        cells: [],
        warnings: [
          {
            code: "EMPTY_USED_RANGE",
            message: `Sheet ${request.sheetName} has no used range to inspect for formulas.`
          }
        ]
      };
    }

    const formulas = normalizeFormulaMatrix(range.formulas as unknown[][]);
    const formulasR1C1 = normalizeFormulaMatrix((range as unknown as { formulasR1C1?: unknown[][] }).formulasR1C1 ?? []);
    const patternMatrix: Array<Array<string | null>> = [];
    const patternMap = new Map<string, { patternHash: string; formulaR1C1: string; count: number; cells: Array<{ rowIndex: number; columnIndex: number }> }>();
    const cells: FormulaPatternResponse["cells"] = [];

    for (let rowIndex = 0; rowIndex < range.rowCount; rowIndex += 1) {
      const row: Array<string | null> = [];
      for (let columnIndex = 0; columnIndex < range.columnCount; columnIndex += 1) {
        const formula = formulas[rowIndex]?.[columnIndex] ?? null;
        const formulaR1C1 = formulasR1C1[rowIndex]?.[columnIndex] ?? formula;
        if (!formula || !formula.startsWith("=")) {
          row.push(null);
          continue;
        }
        const patternHash = hashStable(formulaR1C1 ?? formula);
        row.push(patternHash);
        const pattern = patternMap.get(patternHash) ?? {
          patternHash,
          formulaR1C1: formulaR1C1 ?? formula,
          count: 0,
          cells: []
        };
        pattern.count += 1;
        pattern.cells.push({ rowIndex, columnIndex });
        patternMap.set(patternHash, pattern);
        cells.push({
          rowIndex,
          columnIndex,
          formula,
          ...(formulaR1C1 !== undefined && formulaR1C1 !== null ? { formulaR1C1 } : {}),
          patternHash
        });
      }
      patternMatrix.push(row);
    }

    return {
      workbookId: request.workbookId,
      sheetName: request.sheetName,
      address: stripSheetName(range.address),
      capturedAt: new Date().toISOString(),
      rowCount: range.rowCount,
      columnCount: range.columnCount,
      formulaCount: cells.length,
      formulas,
      formulasR1C1,
      patternMatrix,
      patterns: [...patternMap.values()],
      cells,
      warnings: []
    };
  });
}

export async function copyFormulaPatterns(request: FormulaCopyPatternsRequest): Promise<FormulaMutationResponse> {
  return Excel.run(async (context) => {
    const sourceSheet = context.workbook.worksheets.getItem(request.sourceSheetName);
    const targetSheet = context.workbook.worksheets.getItem(request.targetSheetName);
    const sourceRange = request.sourceAddress ? sourceSheet.getRange(stripSheetName(request.sourceAddress)) : sourceSheet.getUsedRangeOrNullObject();
    sourceRange.load("address,rowCount,columnCount");
    await context.sync();
    if ("isNullObject" in sourceRange && sourceRange.isNullObject) {
      return {
        ok: false,
        formulasChanged: 0,
        warnings: [
          {
            code: "EMPTY_SOURCE_RANGE",
            message: `Sheet ${request.sourceSheetName} has no formulas to copy.`
          }
        ]
      };
    }

    const targetRange = targetSheet.getRange(stripSheetName(request.targetAddress ?? sourceRange.address));
    targetRange.copyFrom(sourceRange, Excel.RangeCopyType.formulas);
    await context.sync();
    return {
      ok: true,
      formulasChanged: sourceRange.rowCount * sourceRange.columnCount,
      warnings: []
    };
  });
}

export async function fillFormulaPattern(request: FormulaFillRequest): Promise<FormulaMutationResponse> {
  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(request.sheetName);
    const sourceRange = worksheet.getRange(stripSheetName(request.sourceAddress));
    const targetRange = worksheet.getRange(stripSheetName(request.targetAddress));
    sourceRange.load("formulasR1C1,rowCount,columnCount");
    targetRange.load("rowCount,columnCount");
    await context.sync();

    const sourceFormulas = normalizeFormulaMatrix((sourceRange as unknown as { formulasR1C1?: unknown[][] }).formulasR1C1 ?? []);
    const warnings: OperationWarning[] = [];
    if (!matrixHasFormula(sourceFormulas)) {
      warnings.push({
        code: "NO_SOURCE_FORMULAS",
        message: `${request.sheetName}!${request.sourceAddress} does not contain formulas to fill.`
      });
      return { ok: false, formulasChanged: 0, warnings };
    }

    const formulas: Array<Array<string | null>> = [];
    for (let rowIndex = 0; rowIndex < targetRange.rowCount; rowIndex += 1) {
      const row: Array<string | null> = [];
      for (let columnIndex = 0; columnIndex < targetRange.columnCount; columnIndex += 1) {
        const sourceRow = request.direction === "down" ? Math.min(rowIndex, sourceFormulas.length - 1) : rowIndex % sourceFormulas.length;
        const sourceColumn = request.direction === "right" ? Math.min(columnIndex, (sourceFormulas[0]?.length ?? 1) - 1) : columnIndex % (sourceFormulas[0]?.length ?? 1);
        row.push(sourceFormulas[sourceRow]?.[sourceColumn] ?? sourceFormulas[0]?.[0] ?? null);
      }
      formulas.push(row);
    }

    (targetRange as unknown as { formulasR1C1: Array<Array<string | null>> }).formulasR1C1 = formulas;
    await context.sync();
    return {
      ok: true,
      formulasChanged: targetRange.rowCount * targetRange.columnCount,
      warnings
    };
  });
}

export async function convertFormulasToValues(request: FormulaPatternRequest): Promise<FormulaMutationResponse> {
  return Excel.run(async (context) => {
    const worksheet = context.workbook.worksheets.getItem(request.sheetName);
    const range = request.address ? worksheet.getRange(stripSheetName(request.address)) : worksheet.getUsedRangeOrNullObject();
    range.load("values,formulas,rowCount,columnCount");
    await context.sync();
    if ("isNullObject" in range && range.isNullObject) {
      return {
        ok: false,
        formulasChanged: 0,
        warnings: [{ code: "EMPTY_USED_RANGE", message: `Sheet ${request.sheetName} has no used range.` }]
      };
    }
    const formulaCount = countFormulas(normalizeFormulaMatrix(range.formulas as unknown[][]));
    range.values = range.values;
    await context.sync();
    return { ok: true, formulasChanged: formulaCount, warnings: [] };
  });
}

export async function repairTemplateConsistency(request: AddinTemplateRepairRequest): Promise<{ ok: boolean; repaired: string[] }> {
  return Excel.run(async (context) => {
    const sourceSheet = context.workbook.worksheets.getItem(request.sourceSheetName);
    const targetSheet = context.workbook.worksheets.getItem(request.targetSheetName);
    const sourceUsedRange = sourceSheet.getUsedRangeOrNullObject();
    sourceUsedRange.load("address");
    await context.sync();

    const repaired: string[] = [];
    if (!sourceUsedRange.isNullObject && request.repair.includes("layout")) {
      const targetRange = targetSheet.getRange(stripSheetName(sourceUsedRange.address));
      targetRange.copyFrom(sourceUsedRange, Excel.RangeCopyType.all);
      repaired.push("layout");
    } else if (!sourceUsedRange.isNullObject) {
      const targetRange = targetSheet.getRange(stripSheetName(sourceUsedRange.address));
      if (request.repair.includes("styles")) {
        targetRange.copyFrom(sourceUsedRange, Excel.RangeCopyType.formats);
        repaired.push("styles");
      }
      if (request.repair.includes("formulas")) {
        targetRange.copyFrom(sourceUsedRange, Excel.RangeCopyType.formulas);
        repaired.push("formulas");
      }
    }

    if (request.repair.includes("dataRegions")) {
      for (const dataRegion of request.dataRegions) {
        targetSheet.getRange(stripSheetName(dataRegion)).clear(Excel.ClearApplyTo.contents);
      }
      repaired.push("dataRegions");
    }

    await context.sync();
    return { ok: true, repaired };
  });
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

async function readRangeMetadata(
  request: RangeMetadataRequest,
  load: (range: Excel.Range) => void,
  materialize: (range: Excel.Range) => unknown
): Promise<RangeMetadataResponse> {
  return Excel.run(async (context) => {
    const target = targetFromMetadataRequest(request);
    const range = getRange(context, target);
    load(range);
    await context.sync();
    return {
      ok: true,
      target,
      data: materialize(range),
      warnings: []
    };
  });
}

async function readSpecialCells(
  request: RangeMetadataRequest,
  cellType: "Blanks" | "Formulas" | "ConditionalFormats" | "DataValidations",
  cellValueType?: "Errors"
): Promise<RangeMetadataResponse> {
  return Excel.run(async (context) => {
    const target = targetFromMetadataRequest(request);
    const range = getRange(context, target);
    const areas = range.getSpecialCellsOrNullObject(cellType, cellValueType);
    areas.load("address,areaCount,cellCount,isNullObject");
    await context.sync();
    return {
      ok: true,
      target,
      data: summarizeRangeAreas(areas),
      warnings: []
    };
  });
}

function unsupportedRangeMetadata(request: RangeMetadataRequest, code: string, message: string): RangeMetadataResponse {
  return {
    ok: false,
    target: targetFromMetadataRequest(request),
    warnings: [
      {
        code,
        message,
        target: targetFromMetadataRequest(request)
      }
    ]
  };
}

function targetFromMetadataRequest(request: RangeMetadataRequest): A1Range {
  return {
    workbookId: request.workbookId,
    sheetName: request.sheetName,
    address: request.address
  };
}

function summarizeRangeAreas(areas: Excel.RangeAreas): { address?: string; areaCount?: number; cellCount?: number; isNullObject?: boolean } {
  const summary: { address?: string; areaCount?: number; cellCount?: number; isNullObject?: boolean } = {};
  assignIfDefined(summary, "address", optionalValue(areas.address));
  assignIfDefined(summary, "areaCount", optionalValue(areas.areaCount));
  assignIfDefined(summary, "cellCount", optionalValue(areas.cellCount));
  assignIfDefined(summary, "isNullObject", optionalValue(areas.isNullObject));
  return summary;
}

function getNamedItem(context: Excel.RequestContext, request: NameSelector): Excel.NamedItem {
  return request.sheetName
    ? context.workbook.worksheets.getItem(request.sheetName).names.getItemOrNullObject(request.name)
    : context.workbook.names.getItemOrNullObject(request.name);
}

function nameReference(context: Excel.RequestContext, request: NameCreateRequest): Excel.Range | string {
  if (request.formula !== undefined) {
    return request.formula;
  }
  if (request.reference === undefined) {
    return "";
  }
  if (request.sheetName !== undefined && !request.reference.startsWith("=")) {
    return context.workbook.worksheets.getItem(request.sheetName).getRange(stripSheetName(request.reference));
  }
  return request.reference;
}

function nameFormula(request: NameUpdateRequest): string {
  if (request.formula !== undefined) {
    return request.formula;
  }
  if (request.reference && request.sheetName && !request.reference.startsWith("=")) {
    return `=${quoteSheetName(request.sheetName)}!${stripSheetName(request.reference)}`;
  }
  return request.reference ?? "";
}

function loadNameWithRange(item: Excel.NamedItem): Excel.Range {
  item.load("name,scope,type,value,formula,comment,visible");
  const range = item.getRangeOrNullObject();
  range.load("address,isNullObject");
  range.worksheet.load("name");
  return range;
}

function materializeNameInfo(workbookId: string, item: Excel.NamedItem, fallbackSheetName?: string, range?: Excel.Range): NameInfo {
  const info: NameInfo = {
    workbookId: workbookId as NameInfo["workbookId"],
    name: item.name,
    scope: item.scope === "Worksheet" ? "worksheet" : "workbook"
  };
  assignIfDefined(info, "sheetName", optionalValue(range?.isNullObject === false ? range.worksheet.name : fallbackSheetName));
  assignIfDefined(info, "type", optionalString(item.type));
  assignIfDefined(info, "value", optionalValue(item.value));
  assignIfDefined(info, "formula", optionalString(item.formula));
  assignIfDefined(info, "comment", optionalValue(item.comment));
  assignIfDefined(info, "visible", optionalValue(item.visible));
  if (range?.isNullObject === false) {
    assignIfDefined(info, "address", optionalValue(range.address));
  }
  return info;
}

async function readPivotTableInfo(context: Excel.RequestContext, workbookId: string, pivot: Excel.PivotTable): Promise<PivotTableInfo> {
  pivot.load("name,id,refreshOnOpen,useCustomSortLists,enableDataValueEditing,allowMultipleFiltersPerField");
  pivot.worksheet.load("name");
  const pivotRange = pivot.layout.getRange();
  pivotRange.load("address,rowCount,columnCount");
  pivot.layout.load("altTextDescription,altTextTitle,autoFormat,emptyCellText,enableFieldList,fillEmptyCells,layoutType,preserveFormatting,showColumnGrandTotals,showFieldHeaders,showRowGrandTotals,subtotalLocation");
  pivot.hierarchies.load("items/name,items/id");
  pivot.rowHierarchies.load("items/name,items/id,items/position");
  pivot.columnHierarchies.load("items/name,items/id,items/position");
  pivot.filterHierarchies.load("items/name,items/id,items/position,items/enableMultipleFilterItems");
  pivot.dataHierarchies.load("items/name,items/id,items/position,items/numberFormat,items/summarizeBy");
  const source = pivot.getDataSourceString();
  const sourceType = pivot.getDataSourceType();
  await context.sync();

  const axisHierarchies = [
    ...pivot.rowHierarchies.items,
    ...pivot.columnHierarchies.items,
    ...pivot.filterHierarchies.items
  ];
  for (const hierarchy of axisHierarchies) {
    hierarchy.fields.load("items/name,items/id,items/showAllItems,items/subtotals");
  }
  for (const hierarchy of pivot.dataHierarchies.items) {
    hierarchy.field.load("name,id,showAllItems,subtotals");
  }
  await context.sync();

  const info: PivotTableInfo = {
    workbookId: workbookId as PivotTableInfo["workbookId"],
    pivotTableName: pivot.name
  };
  assignIfDefined(info, "id", optionalValue(pivot.id));
  assignIfDefined(info, "sheetName", optionalValue(pivot.worksheet.name));
  assignIfDefined(info, "range", {
    address: pivotRange.address,
    rowCount: pivotRange.rowCount,
    columnCount: pivotRange.columnCount
  });
  assignIfDefined(info, "source", optionalValue(source.value));
  assignIfDefined(info, "sourceType", optionalValue(String(sourceType.value)));
  assignIfDefined(info, "refreshOnOpen", optionalValue(pivot.refreshOnOpen));
  assignIfDefined(info, "useCustomSortLists", optionalValue(pivot.useCustomSortLists));
  assignIfDefined(info, "enableDataValueEditing", optionalValue(pivot.enableDataValueEditing));
  assignIfDefined(info, "allowMultipleFiltersPerField", optionalValue(pivot.allowMultipleFiltersPerField));
  const layout: NonNullable<PivotTableInfo["layout"]> = {};
  assignIfDefined(layout, "altTextDescription", optionalValue(pivot.layout.altTextDescription));
  assignIfDefined(layout, "altTextTitle", optionalValue(pivot.layout.altTextTitle));
  assignIfDefined(layout, "autoFormat", optionalValue(pivot.layout.autoFormat));
  assignIfDefined(layout, "emptyCellText", optionalValue(pivot.layout.emptyCellText));
  assignIfDefined(layout, "enableFieldList", optionalValue(pivot.layout.enableFieldList));
  assignIfDefined(layout, "fillEmptyCells", optionalValue(pivot.layout.fillEmptyCells));
  assignIfDefined(layout, "layoutType", optionalString(pivot.layout.layoutType));
  assignIfDefined(layout, "preserveFormatting", optionalValue(pivot.layout.preserveFormatting));
  assignIfDefined(layout, "showColumnGrandTotals", optionalValue(pivot.layout.showColumnGrandTotals));
  assignIfDefined(layout, "showFieldHeaders", optionalValue(pivot.layout.showFieldHeaders));
  assignIfDefined(layout, "showRowGrandTotals", optionalValue(pivot.layout.showRowGrandTotals));
  assignIfDefined(layout, "subtotalLocation", optionalString(pivot.layout.subtotalLocation));
  assignIfDefined(info, "layout", layout);
  assignIfDefined(info, "hierarchies", pivot.hierarchies.items.map((hierarchy) => ({
    ...(hierarchy.id !== undefined ? { id: hierarchy.id } : {}),
    name: hierarchy.name
  })));
  assignIfDefined(info, "rowHierarchies", pivot.rowHierarchies.items.map(materializeAxisHierarchyInfo));
  assignIfDefined(info, "columnHierarchies", pivot.columnHierarchies.items.map(materializeAxisHierarchyInfo));
  assignIfDefined(info, "filterHierarchies", pivot.filterHierarchies.items.map(materializeAxisHierarchyInfo));
  assignIfDefined(info, "dataHierarchies", pivot.dataHierarchies.items.map(materializeDataHierarchyInfo));
  return info;
}

function materializeAxisHierarchyInfo(hierarchy: Excel.RowColumnPivotHierarchy | Excel.FilterPivotHierarchy): NonNullable<PivotTableInfo["rowHierarchies"]>[number] {
  const info: NonNullable<PivotTableInfo["rowHierarchies"]>[number] = {
    name: hierarchy.name
  };
  assignIfDefined(info, "id", optionalValue(hierarchy.id));
  assignIfDefined(info, "position", optionalValue(hierarchy.position));
  if ("enableMultipleFilterItems" in hierarchy) {
    assignIfDefined(info, "enableMultipleFilterItems", optionalValue(hierarchy.enableMultipleFilterItems));
  }
  assignIfDefined(info, "fields", hierarchy.fields.items.map(materializePivotFieldInfo));
  return info;
}

function materializeDataHierarchyInfo(hierarchy: Excel.DataPivotHierarchy): NonNullable<PivotTableInfo["dataHierarchies"]>[number] {
  const info: NonNullable<PivotTableInfo["dataHierarchies"]>[number] = {
    name: hierarchy.name
  };
  assignIfDefined(info, "id", optionalValue(hierarchy.id));
  assignIfDefined(info, "position", optionalValue(hierarchy.position));
  assignIfDefined(info, "numberFormat", optionalValue(hierarchy.numberFormat));
  assignIfDefined(info, "summarizeBy", optionalString(hierarchy.summarizeBy));
  assignIfDefined(info, "field", materializePivotFieldInfo(hierarchy.field));
  return info;
}

function materializePivotFieldInfo(field: Excel.PivotField): NonNullable<NonNullable<PivotTableInfo["rowHierarchies"]>[number]["fields"]>[number] {
  const info: NonNullable<NonNullable<PivotTableInfo["rowHierarchies"]>[number]["fields"]>[number] = {
    name: field.name
  };
  assignIfDefined(info, "id", optionalValue(field.id));
  assignIfDefined(info, "showAllItems", optionalValue(field.showAllItems));
  assignIfDefined(info, "subtotals", optionalValue(field.subtotals as Record<string, unknown>));
  return info;
}

function loadPivotTemplateReplayObjects(pivot: Excel.PivotTable): void {
  pivot.layout.load("altTextDescription,altTextTitle,autoFormat,emptyCellText,enableFieldList,fillEmptyCells,layoutType,preserveFormatting,showColumnGrandTotals,showFieldHeaders,showRowGrandTotals,subtotalLocation");
  pivot.hierarchies.load("items/name,items/id");
  pivot.rowHierarchies.load("items/name,items/id,items/position");
  pivot.columnHierarchies.load("items/name,items/id,items/position");
  pivot.filterHierarchies.load("items/name,items/id,items/position,items/enableMultipleFilterItems");
  pivot.dataHierarchies.load("items/name,items/id,items/position");
}

function hasPivotCreateLayout(request: PivotCreateRequest): boolean {
  return Boolean(
    request.layout !== undefined ||
      request.rowFields?.length ||
      request.columnFields?.length ||
      request.filterFields?.length ||
      request.dataFields?.length
  );
}

function applyPivotCreateLayout(request: PivotCreateRequest, targetPivot: Excel.PivotTable): void {
  if (request.layout !== undefined) {
    targetPivot.layout.set(request.layout as Excel.Interfaces.PivotLayoutUpdateData);
  }
  replayPivotAxis(targetPivot, "row", (request.rowFields ?? []).map(pivotAxisInfoFromName));
  replayPivotAxis(targetPivot, "column", (request.columnFields ?? []).map(pivotAxisInfoFromName));
  replayPivotAxis(targetPivot, "filter", (request.filterFields ?? []).map(pivotAxisInfoFromName));
  replayPivotDataHierarchies(targetPivot, (request.dataFields ?? []).map((field, index) => {
    const info: NonNullable<PivotTableInfo["dataHierarchies"]>[number] = {
      name: field.name ?? field.sourceFieldName,
      position: index,
      field: { name: field.sourceFieldName }
    };
    assignIfDefined(info, "numberFormat", field.numberFormat);
    assignIfDefined(info, "summarizeBy", field.summarizeBy);
    return info;
  }));
}

function pivotAxisInfoFromName(name: string, index: number): NonNullable<PivotTableInfo["rowHierarchies"]>[number] {
  return {
    name,
    position: index
  };
}

function applyPivotTemplateMetadata(source: PivotTableInfo, targetPivot: Excel.PivotTable, dimensions?: PivotCopyFromTemplateRequest["dimensions"]): string[] {
  const includes = (dimension: NonNullable<PivotCopyFromTemplateRequest["dimensions"]>[number]) => !dimensions || dimensions.includes(dimension);
  const copied: string[] = [];
  if (includes("metadata")) {
    targetPivot.refreshOnOpen = source.refreshOnOpen ?? targetPivot.refreshOnOpen;
    targetPivot.useCustomSortLists = source.useCustomSortLists ?? targetPivot.useCustomSortLists;
    targetPivot.enableDataValueEditing = source.enableDataValueEditing ?? targetPivot.enableDataValueEditing;
    targetPivot.allowMultipleFiltersPerField = source.allowMultipleFiltersPerField ?? targetPivot.allowMultipleFiltersPerField;
    copied.push("refreshOnOpen", "useCustomSortLists", "enableDataValueEditing", "allowMultipleFiltersPerField");
  }
  if (includes("layout") && source.layout !== undefined) {
    targetPivot.layout.set(source.layout as Excel.Interfaces.PivotLayoutUpdateData);
    copied.push("layout");
  }
  if (includes("fields")) {
    replayPivotAxis(targetPivot, "row", source.rowHierarchies ?? []);
    replayPivotAxis(targetPivot, "column", source.columnHierarchies ?? []);
    copied.push("rowHierarchyPositions", "columnHierarchyPositions", "fieldSettings");
  }
  if (includes("filters")) {
    replayPivotAxis(targetPivot, "filter", source.filterHierarchies ?? []);
    copied.push("filterHierarchyPositions");
  }
  if (includes("dataFields")) {
    replayPivotDataHierarchies(targetPivot, source.dataHierarchies ?? [], includes("numberFormats"));
    copied.push("dataHierarchySettings");
    if (includes("numberFormats")) {
      copied.push("dataHierarchyNumberFormats");
    }
  }
  if (includes("refresh")) {
    copied.push("refresh");
  }
  return copied;
}

function replayPivotAxis(
  targetPivot: Excel.PivotTable,
  axis: "row" | "column" | "filter",
  hierarchies: NonNullable<PivotTableInfo["rowHierarchies"]>
): void {
  if (axis === "filter") {
    replayFilterPivotAxis(targetPivot, hierarchies);
    return;
  }
  const collection = axis === "row" ? targetPivot.rowHierarchies : targetPivot.columnHierarchies;
  for (const existing of collection.items) {
    collection.remove(existing);
  }
  for (const [index, hierarchyInfo] of hierarchies.entries()) {
    const hierarchy = targetPivot.hierarchies.getItem(hierarchyInfo.name);
    const added = collection.add(hierarchy);
    added.position = hierarchyInfo.position ?? index;
    for (const fieldInfo of hierarchyInfo.fields ?? []) {
      const field = added.fields.getItem(fieldInfo.name);
      applyPivotFieldMetadata(field, fieldInfo);
    }
  }
}

function replayFilterPivotAxis(targetPivot: Excel.PivotTable, hierarchies: NonNullable<PivotTableInfo["filterHierarchies"]>): void {
  for (const existing of targetPivot.filterHierarchies.items) {
    targetPivot.filterHierarchies.remove(existing);
  }
  for (const [index, hierarchyInfo] of hierarchies.entries()) {
    const hierarchy = targetPivot.hierarchies.getItem(hierarchyInfo.name);
    const added = targetPivot.filterHierarchies.add(hierarchy);
    added.position = hierarchyInfo.position ?? index;
    if (hierarchyInfo.enableMultipleFilterItems !== undefined) {
      added.enableMultipleFilterItems = hierarchyInfo.enableMultipleFilterItems;
    }
    for (const fieldInfo of hierarchyInfo.fields ?? []) {
      const field = added.fields.getItem(fieldInfo.name);
      applyPivotFieldMetadata(field, fieldInfo);
    }
  }
}

function replayPivotDataHierarchies(targetPivot: Excel.PivotTable, hierarchies: NonNullable<PivotTableInfo["dataHierarchies"]>, includeNumberFormats = true): void {
  for (const existing of targetPivot.dataHierarchies.items) {
    targetPivot.dataHierarchies.remove(existing);
  }
  for (const [index, hierarchyInfo] of hierarchies.entries()) {
    const sourceFieldName = hierarchyInfo.field?.name ?? hierarchyInfo.name;
    const added = targetPivot.dataHierarchies.add(targetPivot.hierarchies.getItem(sourceFieldName));
    added.position = hierarchyInfo.position ?? index;
    if (hierarchyInfo.name !== undefined) {
      added.name = hierarchyInfo.name;
    }
    if (includeNumberFormats && hierarchyInfo.numberFormat !== undefined) {
      added.numberFormat = hierarchyInfo.numberFormat;
    }
    if (hierarchyInfo.summarizeBy !== undefined) {
      added.summarizeBy = hierarchyInfo.summarizeBy as Excel.AggregationFunction;
    }
    if (hierarchyInfo.field !== undefined) {
      applyPivotFieldMetadata(added.field, hierarchyInfo.field);
    }
  }
}

function applyPivotFieldMetadata(field: Excel.PivotField, fieldInfo: NonNullable<NonNullable<PivotTableInfo["rowHierarchies"]>[number]["fields"]>[number]): void {
  if (fieldInfo.showAllItems !== undefined) {
    field.showAllItems = fieldInfo.showAllItems;
  }
  if (fieldInfo.subtotals !== undefined) {
    field.subtotals = fieldInfo.subtotals as Excel.Subtotals;
  }
}

function loadChartInfoObjects(workbookId: string, sheetName: string, chart: Excel.Chart): {
  workbookId: string;
  sheetName: string;
  chart: Excel.Chart;
} {
  chart.load("name,id,chartType,top,left,width,height,style,plotBy");
  chart.title.load("text");
  return { workbookId, sheetName, chart };
}

function materializeChartInfo(loaded: ReturnType<typeof loadChartInfoObjects>): ChartInfo {
  const info: ChartInfo = {
    workbookId: loaded.workbookId as ChartInfo["workbookId"],
    sheetName: loaded.sheetName,
    chartName: loaded.chart.name
  };
  assignIfDefined(info, "id", optionalValue(loaded.chart.id));
  assignIfDefined(info, "chartType", optionalValue(String(loaded.chart.chartType)));
  assignIfDefined(info, "top", optionalValue(loaded.chart.top));
  assignIfDefined(info, "left", optionalValue(loaded.chart.left));
  assignIfDefined(info, "width", optionalValue(loaded.chart.width));
  assignIfDefined(info, "height", optionalValue(loaded.chart.height));
  assignIfDefined(info, "style", optionalValue(loaded.chart.style));
  assignIfDefined(info, "plotBy", optionalValue(String(loaded.chart.plotBy)));
  assignIfDefined(info, "title", optionalValue(loaded.chart.title.text));
  return info;
}

function loadTableInfoObjects(table: Excel.Table): {
  table: Excel.Table;
  range: Excel.Range;
  headerRange: Excel.Range;
  columns: Excel.TableColumnCollection;
  autoFilter: Excel.AutoFilter;
  sort: Excel.TableSort;
} {
  table.load("name,id,style,showHeaders,showTotals,showFilterButton,showBandedRows,showBandedColumns");
  const range = table.getRange();
  const headerRange = table.getHeaderRowRange();
  const columns = table.columns;
  const autoFilter = table.autoFilter;
  const sort = table.sort;
  range.load("address,rowCount,columnCount");
  range.worksheet.load("name");
  headerRange.load("address");
  columns.load("count,items/id,items/index,items/name");
  table.rows.load("count");
  autoFilter.load("criteria,enabled,isDataFiltered");
  sort.load("fields,matchCase,method");
  return { table, range, headerRange, columns, autoFilter, sort };
}

function materializeTableInfo(workbookId: string, loaded: ReturnType<typeof loadTableInfoObjects>): TableInfo {
  const info: TableInfo = {
    workbookId: workbookId as TableInfo["workbookId"],
    tableName: loaded.table.name,
    rowCount: loaded.table.rows.count,
    columnCount: loaded.columns.count,
    columns: loaded.columns.items.map((column) => ({
      id: column.id,
      index: column.index,
      name: column.name
    }))
  };
  assignIfDefined(info, "id", optionalValue(loaded.table.id));
  assignIfDefined(info, "sheetName", optionalValue(loaded.range.worksheet.name));
  assignIfDefined(info, "address", optionalValue(loaded.range.address));
  assignIfDefined(info, "headerAddress", optionalValue(loaded.headerRange.address));
  assignIfDefined(info, "style", optionalValue(loaded.table.style));
  assignIfDefined(info, "showHeaders", optionalValue(loaded.table.showHeaders));
  assignIfDefined(info, "showTotals", optionalValue(loaded.table.showTotals));
  assignIfDefined(info, "showFilterButton", optionalValue(loaded.table.showFilterButton));
  assignIfDefined(info, "showBandedRows", optionalValue(loaded.table.showBandedRows));
  assignIfDefined(info, "showBandedColumns", optionalValue(loaded.table.showBandedColumns));
  assignIfDefined(info, "filters", {
    enabled: loaded.autoFilter.enabled,
    isDataFiltered: loaded.autoFilter.isDataFiltered,
    criteria: loaded.autoFilter.criteria
  });
  assignIfDefined(info, "sort", {
    fields: loaded.sort.fields,
    matchCase: loaded.sort.matchCase,
    method: loaded.sort.method
  });
  return info;
}

function resolveTableColumnIndexes(info: TableInfo, requestedColumns: Array<string | number> | undefined): number[] {
  if (!requestedColumns?.length) {
    return info.columns.map((column) => column.index);
  }
  const indexes: number[] = [];
  for (const requested of requestedColumns) {
    const column =
      typeof requested === "number"
        ? info.columns.find((candidate) => candidate.index === requested)
        : info.columns.find((candidate) => candidate.name === requested);
    if (column && !indexes.includes(column.index)) {
      indexes.push(column.index);
    }
  }
  return indexes;
}

function projectMatrix<T>(matrix: T[][], columnIndexes: number[]): T[][] {
  return matrix.map((row) => columnIndexes.map((columnIndex) => row[columnIndex] as T));
}

function materializeProjectedRangeFacet(
  dataRange: Excel.Range,
  projectedColumnRanges: Excel.Range[],
  property: "values" | "formulas" | "text" | "numberFormat",
  projectedIndexes: number[]
): unknown[][] {
  if (projectedColumnRanges.length === 0) {
    return projectMatrix((dataRange as unknown as Record<typeof property, unknown[][]>)[property] ?? [], projectedIndexes);
  }
  const columns = projectedColumnRanges.map((columnRange) => ((columnRange as unknown as Record<typeof property, unknown[][]>)[property] ?? []).map((row) => row[0]));
  const rowCount = Math.max(0, ...columns.map((column) => column.length));
  const rows: unknown[][] = [];
  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    rows.push(columns.map((column) => column[rowIndex]));
  }
  return rows;
}

async function mutateTableAndReturnInfo<TRequest extends TableSelector>(
  request: TRequest,
  mutate: (table: Excel.Table, context: Excel.RequestContext) => void
): Promise<{ ok: boolean; info: TableInfo }> {
  return Excel.run(async (context) => {
    const table = context.workbook.tables.getItem(request.tableName);
    mutate(table, context);
    const loaded = loadTableInfoObjects(table);
    await context.sync();
    return { ok: true, info: materializeTableInfo(request.workbookId, loaded) };
  });
}

function getRange(context: Excel.RequestContext, target: A1Range): Excel.Range {
  return context.workbook.worksheets.getItem(target.sheetName).getRange(stripSheetName(target.address));
}

function rangeSnapshotFacetsForOperation(operation: Extract<ExcelOperation, { kind: "range.read_full" }>): RangeSnapshotFacet[] {
  if (operation.facets?.length) {
    return [...new Set(operation.facets)];
  }
  const facets: RangeSnapshotFacet[] = ["values", "numberFormat", "text"];
  if (operation.includeFormulas !== false) {
    facets.push("formulas");
  }
  if (operation.includeStyles !== false) {
    facets.push("style");
  }
  return facets;
}

function allRangeSnapshotFacets(): RangeSnapshotFacet[] {
  return ["values", "formulas", "numberFormat", "text", "style"];
}

function loadSnapshotProperties(range: Excel.Range, facets: RangeSnapshotFacet[] = allRangeSnapshotFacets()): void {
  const properties = ["rowCount", "columnCount"];
  if (facets.includes("values")) {
    properties.push("values");
  }
  if (facets.includes("formulas")) {
    properties.push("formulas");
  }
  if (facets.includes("numberFormat")) {
    properties.push("numberFormat");
  }
  if (facets.includes("text")) {
    properties.push("text");
  }
  range.load(properties.join(","));
  if (facets.includes("style")) {
    range.format.load("horizontalAlignment, verticalAlignment, rowHeight, columnWidth");
    range.format.fill.load("color");
    range.format.font.load("name, size, color, bold, italic");
  }
}

function materializeSnapshot(target: A1Range, range: Excel.Range, facets: RangeSnapshotFacet[] = allRangeSnapshotFacets()): RangeSnapshot {
  const values = facets.includes("values") ? range.values as RangeSnapshot["values"] : undefined;
  const formulas = facets.includes("formulas") ? range.formulas as RangeSnapshot["formulas"] : undefined;
  const numberFormat = facets.includes("numberFormat") ? range.numberFormat as string[][] : undefined;
  const text = facets.includes("text") ? range.text as string[][] : undefined;
  const style: NonNullable<RangeSnapshot["style"]> = {};
  if (facets.includes("style")) {
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
  }

  const snapshot: RangeSnapshot = {
    fingerprint: createRangeFingerprint(target, {
      values,
      formulas,
      numberFormat,
      text,
      style
    })
  };
  if (values !== undefined) {
    snapshot.values = values;
  }
  if (formulas !== undefined) {
    snapshot.formulas = formulas;
  }
  if (numberFormat !== undefined) {
    snapshot.numberFormat = numberFormat;
  }
  if (text !== undefined) {
    snapshot.text = text;
  }
  if (Object.keys(style).length > 0) {
    snapshot.style = style;
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

function restoreRangeSnapshot(
  context: Excel.RequestContext,
  operation: Extract<ExcelOperation, { kind: "range.restore_snapshot" }>
): void {
  const range = getRange(context, operation.target);
  if (operation.snapshot.formulas) {
    range.formulas = operation.snapshot.formulas;
  } else if (operation.snapshot.values) {
    range.values = operation.snapshot.values;
  }
  if (operation.snapshot.numberFormat) {
    range.numberFormat = operation.snapshot.numberFormat;
  }

  const style = operation.snapshot.style;
  if (!style) {
    return;
  }
  if (style.fillColor) {
    range.format.fill.color = style.fillColor;
  }
  if (style.fontName) {
    range.format.font.name = style.fontName;
  }
  if (style.fontSize !== undefined) {
    range.format.font.size = style.fontSize;
  }
  if (style.fontColor) {
    range.format.font.color = style.fontColor;
  }
  if (style.fontBold !== undefined) {
    range.format.font.bold = style.fontBold;
  }
  if (style.fontItalic !== undefined) {
    range.format.font.italic = style.fontItalic;
  }
  if (style.horizontalAlignment) {
    range.format.horizontalAlignment = style.horizontalAlignment as Excel.HorizontalAlignment;
  }
  if (style.verticalAlignment) {
    range.format.verticalAlignment = style.verticalAlignment as Excel.VerticalAlignment;
  }
  if (style.rowHeight !== undefined) {
    range.format.rowHeight = style.rowHeight;
  }
  if (style.columnWidth !== undefined) {
    range.format.columnWidth = style.columnWidth;
  }
}

function applyRangeStyle(range: Excel.Range, style: NonNullable<RangeSnapshot["style"]>): void {
  if (style.fillColor) {
    range.format.fill.color = style.fillColor;
  }
  if (style.fontName) {
    range.format.font.name = style.fontName;
  }
  if (style.fontSize !== undefined) {
    range.format.font.size = style.fontSize;
  }
  if (style.fontColor) {
    range.format.font.color = style.fontColor;
  }
  if (style.fontBold !== undefined) {
    range.format.font.bold = style.fontBold;
  }
  if (style.fontItalic !== undefined) {
    range.format.font.italic = style.fontItalic;
  }
  if (style.horizontalAlignment) {
    range.format.horizontalAlignment = style.horizontalAlignment as Excel.HorizontalAlignment;
  }
  if (style.verticalAlignment) {
    range.format.verticalAlignment = style.verticalAlignment as Excel.VerticalAlignment;
  }
  if (style.rowHeight !== undefined) {
    range.format.rowHeight = style.rowHeight;
  }
  if (style.columnWidth !== undefined) {
    range.format.columnWidth = style.columnWidth;
  }
}

function normalizeFormulaMatrix(matrix: unknown[][]): Array<Array<string | null>> {
  return matrix.map((row) =>
    row.map((value) => {
      if (typeof value !== "string" || value.length === 0) {
        return null;
      }
      return value;
    })
  );
}

function matrixHasFormula(matrix: Array<Array<string | null>>): boolean {
  for (const row of matrix) {
    for (const formula of row) {
      if (formula?.startsWith("=")) {
        return true;
      }
    }
  }
  return false;
}

function countFormulas(matrix: Array<Array<string | null>>): number {
  let count = 0;
  for (const row of matrix) {
    for (const formula of row) {
      if (formula?.startsWith("=")) {
        count += 1;
      }
    }
  }
  return count;
}

function toClearApplyTo(applyTo: "all" | "contents" | "formats" | "hyperlinks"): Excel.ClearApplyTo {
  switch (applyTo) {
    case "contents":
      return Excel.ClearApplyTo.contents;
    case "formats":
      return Excel.ClearApplyTo.formats;
    case "hyperlinks":
      return Excel.ClearApplyTo.hyperlinks;
    case "all":
      return Excel.ClearApplyTo.all;
  }
}

function toRangeCopyType(copyType: "all" | "values" | "formats" | "formulas"): Excel.RangeCopyType {
  switch (copyType) {
    case "values":
      return Excel.RangeCopyType.values;
    case "formats":
      return Excel.RangeCopyType.formats;
    case "formulas":
      return Excel.RangeCopyType.formulas;
    case "all":
      return Excel.RangeCopyType.all;
  }
}

function toWorksheetPositionType(position: "beginning" | "end" | "before" | "after"): Excel.WorksheetPositionType {
  switch (position) {
    case "beginning":
      return Excel.WorksheetPositionType.beginning;
    case "end":
      return Excel.WorksheetPositionType.end;
    case "before":
      return Excel.WorksheetPositionType.before;
    case "after":
      return Excel.WorksheetPositionType.after;
  }
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
  const parsed = parseA1Address(stripSheetName(target.address));
  if (parsed.endRow - parsed.startRow + 1 !== matrix.length || parsed.endColumn - parsed.startColumn + 1 !== matrix[0]!.length) {
    throw new Error(`Matrix dimensions do not match ${target.sheetName}!${target.address}`);
  }
}

function matrixCellCount(matrix: unknown[][]): number {
  return matrix.reduce((count, row) => count + row.length, 0);
}

function writeMatrixInChunks(
  context: Excel.RequestContext,
  target: A1Range,
  matrix: unknown[][],
  property: "values" | "formulas" | "numberFormat"
): number {
  const parsed = parseA1Address(stripSheetName(target.address));
  const columnCount = matrix[0]?.length ?? 0;
  const worksheet = context.workbook.worksheets.getItem(target.sheetName);
  let chunkCount = 0;
  for (const chunk of chunkMatrixRows(matrix, CHUNK_CELL_LIMIT)) {
    const range = worksheet.getRangeByIndexes(parsed.startRow - 1 + chunk.rowOffset, parsed.startColumn - 1, chunk.rows.length, columnCount);
    (range as unknown as Record<typeof property, unknown[][]>)[property] = chunk.rows;
    chunkCount += 1;
  }
  return chunkCount;
}

function optionalValue<T>(value: T | null | undefined): T | undefined {
  return value === null || value === undefined ? undefined : value;
}

function getDocumentFile(sliceSize: number): Promise<Office.File> {
  return new Promise((resolve, reject) => {
    Office.context.document.getFileAsync(Office.FileType.Compressed, { sliceSize }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
        return;
      }
      reject(new Error(result.error?.message ?? "Office Document.getFileAsync failed."));
    });
  });
}

function getDocumentFileSlice(file: Office.File, index: number): Promise<Office.Slice> {
  return new Promise((resolve, reject) => {
    file.getSliceAsync(index, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
        return;
      }
      reject(new Error(result.error?.message ?? `Office File.getSliceAsync failed for slice ${index}.`));
    });
  });
}

function closeDocumentFile(file: Office.File): Promise<void> {
  return new Promise((resolve) => {
    file.closeAsync(() => resolve());
  });
}

function sliceDataToBase64(data: unknown): string {
  if (typeof data === "string") {
    return data;
  }
  if (data instanceof Uint8Array) {
    return uint8ArrayToBase64(data);
  }
  if (Array.isArray(data)) {
    return uint8ArrayToBase64(Uint8Array.from(data as number[]));
  }
  if (data instanceof ArrayBuffer) {
    return uint8ArrayToBase64(new Uint8Array(data));
  }
  throw new Error("Unsupported Office file slice data format.");
}

function uint8ArrayToBase64(bytes: Uint8Array): string {
  let binary = "";
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    const chunk = bytes.subarray(offset, offset + chunkSize);
    binary += String.fromCharCode(...chunk);
  }
  return btoa(binary);
}

function getCustomXmlParts(context: Excel.RequestContext): any | undefined {
  const workbook = context.workbook as unknown as { customXmlParts?: any };
  return workbook.customXmlParts;
}

function workbookLocalConfigXml(config: WorkbookLocalConfig): string {
  const json = JSON.stringify(config);
  return [
    `<owb:localConfig xmlns:owb="${OPEN_WORKBOOK_CUSTOM_XML_NAMESPACE}" version="1" workbookId="${escapeXmlAttribute(config.workbookId)}">`,
    `<owb:json>${cdata(json)}</owb:json>`,
    "</owb:localConfig>"
  ].join("");
}

function parseWorkbookLocalConfigXml(xml: string): WorkbookLocalConfig {
  const cdataMatch = xml.match(/<owb:json>\s*<!\[CDATA\[([\s\S]*)\]\]>\s*<\/owb:json>/);
  const textMatch = xml.match(/<owb:json>([\s\S]*?)<\/owb:json>/);
  const rawJson = cdataMatch?.[1] ?? (textMatch?.[1] ? unescapeXmlText(textMatch[1]) : undefined);
  if (!rawJson) {
    throw new Error("Open Workbook custom XML part does not contain local config JSON.");
  }
  return JSON.parse(rawJson) as WorkbookLocalConfig;
}

function cdata(value: string): string {
  return `<![CDATA[${value.replaceAll("]]>", "]]]]><![CDATA[>")}]]>`;
}

function escapeXmlAttribute(value: string): string {
  return value.replaceAll("&", "&amp;").replaceAll("\"", "&quot;").replaceAll("<", "&lt;").replaceAll(">", "&gt;");
}

function unescapeXmlText(value: string): string {
  return value.replaceAll("&lt;", "<").replaceAll("&gt;", ">").replaceAll("&quot;", "\"").replaceAll("&apos;", "'").replaceAll("&amp;", "&");
}

function assignIfDefined<T extends object, K extends keyof T>(target: T, key: K, value: T[K] | undefined): void {
  if (value !== undefined) {
    target[key] = value;
  }
}

function stripSheetName(address: string): string {
  const bangIndex = address.lastIndexOf("!");
  return bangIndex >= 0 ? address.slice(bangIndex + 1) : address;
}

function cellPositionFromZeroBased(workbookId: WorkbookRef["workbookId"], sheetName: string, rowIndex: number, columnIndex: number): CellPosition {
  const row = rowIndex + 1;
  const column = columnIndex + 1;
  return {
    workbookId,
    sheetName,
    address: formatA1Cell(row, column),
    row,
    column,
    rowIndex,
    columnIndex
  };
}

function quoteSheetName(sheetName: string): string {
  return `'${sheetName.replace(/'/g, "''")}'`;
}

function optionalString(value: unknown): string | undefined {
  return value === null || value === undefined ? undefined : String(value);
}

function createTelemetry(started: number, counters: ExecutionCounters, warnings: OperationWarning[]): OperationTelemetry {
  return {
    durationMs: Math.round(performance.now() - started),
    syncCount: counters.syncCount,
    payloadBytes: 0,
    cellsRead: counters.cellsRead,
    cellsWritten: counters.cellsWritten,
    rangeCount: counters.rangeCount,
    chunkCount: Math.max(1, counters.chunkCount),
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

function isOfficeApiSetSupported(set: string, version: string): boolean {
  try {
    return Office.context.requirements.isSetSupported(set, version);
  } catch {
    return false;
  }
}

function hostCapability(name: string, supported: boolean, set: string, version: string): NonNullable<RuntimeCapabilities["hostCapabilities"]>[number] {
  return {
    name,
    supported,
    status: supported ? "supported" : "unsupported",
    ...(supported ? {} : { reason: `${set} ${version} is not supported by this Excel host.` }),
    requires: [{ set, version }]
  };
}
