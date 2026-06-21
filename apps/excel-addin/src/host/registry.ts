import type { HostMethodDefinition } from "./types.js";
import { batchHostOperations } from "./batch.js";
import { formulaHostOperations } from "./formula.js";
import { namesHostOperations } from "./names.js";
import { pivotChartHostOperations } from "./pivot-chart.js";
import { rangeHostOperations } from "./range.js";
import { runtimeHostOperations } from "./runtime.js";
import { styleHostOperations } from "./style.js";
import { tableHostOperations } from "./table.js";
import { templateHostOperations } from "./template.js";
import { workbookHostOperations } from "./workbook.js";

export const BATCH_OPERATION_KINDS = [
  "range.read_full",
  "range.write_values",
  "range.write_values_many",
  "range.write_formulas",
  "range.write_number_formats",
  "range.write_number_formats_many",
  "range.write_styles",
  "range.write_styles_many",
  "range.write_data_validation",
  "range.write_conditional_formatting",
  "range.clear_style_dimensions",
  "range.clear_style_dimensions_many",
  "range.write_hyperlinks",
  "range.write_comments",
  "range.clear_values_keep_format",
  "range.clear",
  "range.clear_many",
  "range.clear_values",
  "range.clear_formats",
  "range.clear_formats_many",
  "range.copy",
  "range.move",
  "range.reorder_columns",
  "range.insert_rows",
  "range.insert_columns",
  "range.delete_rows",
  "range.delete_columns",
  "range.autofit_columns",
  "range.autofit_rows",
  "range.autofit_many",
  "range.apply_autofilter",
  "range.merge",
  "range.unmerge",
  "range.restore_snapshot",
  "workbook.calculate",
  "workbook.save",
  "sheet.create",
  "sheet.copy",
  "sheet.copy_clean_data_regions",
  "sheet.rename",
  "sheet.delete",
  "sheet.move",
  "sheet.hide",
  "sheet.unhide",
  "sheet.protect",
  "sheet.unprotect",
  "sheet.clear",
  "sheet.set_tab_color",
  "template.create_sheet_from_template"
] as const;

export const HOST_METHOD_REGISTRY: HostMethodDefinition[] = [
  {
    method: "runtime.ping",
    implementationOwner: "RuntimeHostOperations",
    handler: runtimeHostOperations.ping,
    relatedBackendCapabilities: ["excel.runtime.ping_addin"],
    operationKinds: [],
    hostDependency: "backend-only",
    unitTestFile: "apps/excel-addin/src/host/runtime.test.ts"
  },
  {
    method: "runtime.get_active_context",
    implementationOwner: "RuntimeHostOperations",
    handler: runtimeHostOperations.getActiveWorkbookContext,
    relatedBackendCapabilities: ["excel.runtime.get_active_context"],
    operationKinds: [],
    hostDependency: "office-js",
    unitTestFile: "apps/excel-addin/src/host/runtime.test.ts"
  },
  {
    method: "runtime.get_selection",
    implementationOwner: "RuntimeHostOperations",
    handler: runtimeHostOperations.getSelection,
    relatedBackendCapabilities: ["excel.runtime.get_selection"],
    operationKinds: [],
    hostDependency: "office-js",
    unitTestFile: "apps/excel-addin/src/host/runtime.test.ts"
  },
  {
    method: "runtime.set_active_sheet",
    implementationOwner: "RuntimeHostOperations",
    handler: runtimeHostOperations.setActiveSheet,
    relatedBackendCapabilities: ["excel.runtime.set_active_sheet"],
    operationKinds: [],
    hostDependency: "office-js",
    unitTestFile: "apps/excel-addin/src/host/runtime.test.ts"
  },
  {
    method: "workbook.get_info",
    implementationOwner: "WorkbookHostOperations",
    handler: workbookHostOperations.getWorkbookInfo,
    relatedBackendCapabilities: ["excel.workbook.get_workbook_info"],
    operationKinds: [],
    hostDependency: "office-js",
    unitTestFile: "apps/excel-addin/src/host/workbook.test.ts"
  },
  {
    method: "workbook.get_map",
    implementationOwner: "WorkbookHostOperations",
    handler: workbookHostOperations.getWorkbookMap,
    relatedBackendCapabilities: ["excel.workbook.get_workbook_map", "excel.workbook.get_summary", "excel.workbook.get_used_range_summary"],
    operationKinds: [],
    hostDependency: "office-js",
    unitTestFile: "apps/excel-addin/src/host/workbook.test.ts"
  },
  {
    method: "workbook.calculate",
    implementationOwner: "WorkbookHostOperations",
    handler: workbookHostOperations.calculateWorkbook,
    relatedBackendCapabilities: ["excel.workbook.calculate", "excel.formula.recalculate"],
    operationKinds: ["workbook.calculate"],
    hostDependency: "office-js",
    unitTestFile: "apps/excel-addin/src/host/workbook.test.ts"
  },
  {
    method: "workbook.save",
    implementationOwner: "WorkbookHostOperations",
    handler: workbookHostOperations.saveWorkbook,
    relatedBackendCapabilities: ["excel.workbook.save"],
    operationKinds: ["workbook.save"],
    hostDependency: "office-js",
    unitTestFile: "apps/excel-addin/src/host/workbook.test.ts"
  },
  {
    method: "workbook.get_file",
    implementationOwner: "WorkbookHostOperations",
    handler: workbookHostOperations.exportWorkbookFile,
    relatedBackendCapabilities: ["excel.workbook.export_copy"],
    operationKinds: [],
    hostDependency: "office-document-file",
    unitTestFile: "apps/excel-addin/src/host/workbook.test.ts"
  },
  {
    method: "workbook.close",
    implementationOwner: "WorkbookHostOperations",
    handler: workbookHostOperations.closeWorkbook,
    relatedBackendCapabilities: ["excel.workbook.close"],
    operationKinds: [],
    hostDependency: "office-js",
    unitTestFile: "apps/excel-addin/src/host/workbook.test.ts"
  },
  {
    method: "workbook.snapshot_ranges",
    implementationOwner: "WorkbookHostOperations",
    handler: workbookHostOperations.snapshotRanges,
    relatedBackendCapabilities: [
      "excel.workbook.snapshot",
      "excel.workbook.refresh_snapshot",
      "excel.workbook.get_snapshot",
      "excel.workbook.detect_external_changes",
      "excel.workbook.create_backup",
      "excel.workbook.restore_backup",
      "excel.snapshot.create",
      "excel.snapshot.refresh",
      "excel.snapshot.get_compact",
      "excel.snapshot.compare_compact"
    ],
    operationKinds: [],
    hostDependency: "office-js",
    unitTestFile: "apps/excel-addin/src/host/workbook.test.ts"
  },
  {
    method: "workbook.embed_local_config",
    implementationOwner: "WorkbookHostOperations",
    handler: workbookHostOperations.embedWorkbookLocalConfig,
    relatedBackendCapabilities: ["excel.workbook.embed_local_config", "excel.workbook.import_embedded_local_config"],
    operationKinds: [],
    hostDependency: "office-js",
    unitTestFile: "apps/excel-addin/src/host/workbook.test.ts"
  },
  {
    method: "workbook.read_embedded_local_config",
    implementationOwner: "WorkbookHostOperations",
    handler: workbookHostOperations.readWorkbookEmbeddedLocalConfig,
    relatedBackendCapabilities: ["excel.workbook.read_embedded_local_config"],
    operationKinds: [],
    hostDependency: "office-js",
    unitTestFile: "apps/excel-addin/src/host/workbook.test.ts"
  },
  ...methodGroup("names", "NamesHostOperations", "apps/excel-addin/src/host/names.test.ts", [
    ["names.list", namesHostOperations.listNames, ["excel.names.list"]],
    ["names.get", namesHostOperations.getName, ["excel.names.get"]],
    ["names.create", namesHostOperations.createName, ["excel.names.create"]],
    ["names.update", namesHostOperations.updateName, ["excel.names.update"]],
    ["names.delete", namesHostOperations.deleteName, ["excel.names.delete"]]
  ]),
  ...methodGroup("pivot", "PivotChartHostOperations", "apps/excel-addin/src/host/pivot-chart.test.ts", [
    ["pivot.list", pivotChartHostOperations.listPivotTables, ["excel.pivot.list"]],
    ["pivot.get_info", pivotChartHostOperations.getPivotTableInfo, ["excel.pivot.get_info"]],
    ["pivot.create", pivotChartHostOperations.createPivotTable, ["excel.pivot.create"]],
    ["pivot.refresh", pivotChartHostOperations.refreshPivotTable, ["excel.pivot.refresh"]],
    ["pivot.refresh_all", pivotChartHostOperations.refreshAllPivotTables, ["excel.pivot.refresh_all"]],
    ["pivot.copy_from_template", pivotChartHostOperations.copyPivotTableFromTemplate, ["excel.pivot.copy_from_template", "excel.pivot.repair_from_template"]],
    ["pivot.delete", pivotChartHostOperations.deletePivotTable, ["excel.pivot.delete"]]
  ]),
  ...methodGroup("chart", "PivotChartHostOperations", "apps/excel-addin/src/host/pivot-chart.test.ts", [
    ["chart.list", pivotChartHostOperations.listCharts, ["excel.chart.list"]],
    ["chart.get_info", pivotChartHostOperations.getChartInfo, ["excel.chart.get_info"]],
    ["chart.create", pivotChartHostOperations.createChart, ["excel.chart.create"]],
    ["chart.update_data_source", pivotChartHostOperations.updateChartDataSource, ["excel.chart.update_data_source"]],
    ["chart.copy_from_template", pivotChartHostOperations.copyChartFromTemplate, ["excel.chart.copy_from_template"]],
    ["chart.refresh", pivotChartHostOperations.refreshChart, ["excel.chart.refresh"]],
    ["chart.delete", pivotChartHostOperations.deleteChart, ["excel.chart.delete"]]
  ]),
  ...methodGroup("range", "RangeHostOperations", "apps/excel-addin/src/host/range.test.ts", [
    ["range.read_hyperlinks", rangeHostOperations.readRangeHyperlinks, ["excel.range.read_hyperlinks"]],
    ["range.read_comments", rangeHostOperations.readRangeComments, ["excel.range.read_comments"]],
    ["range.read_notes", rangeHostOperations.readRangeNotes, ["excel.range.read_notes"]],
    ["range.read_merged_cells", rangeHostOperations.readRangeMergedCells, ["excel.range.read_merged_cells"]],
    ["range.read_data_validation", rangeHostOperations.readRangeDataValidation, ["excel.range.read_data_validation"]],
    ["range.read_conditional_formatting", rangeHostOperations.readRangeConditionalFormatting, ["excel.range.read_conditional_formatting"]],
    ["range.search", rangeHostOperations.searchRange, ["excel.range.search"]],
    ["range.find_blank_cells", rangeHostOperations.findBlankCells, ["excel.range.find_blank_cells"]],
    ["range.find_errors", rangeHostOperations.findFormulaErrors, ["excel.range.find_errors", "excel.formula.find_errors"]]
  ]),
  ...methodGroup("formula", "FormulaHostOperations", "apps/excel-addin/src/host/formula.test.ts", [
    ["formula.read_patterns", formulaHostOperations.readFormulaPatterns, ["excel.formula.read_patterns"]],
    ["formula.copy_patterns", formulaHostOperations.copyFormulaPatterns, ["excel.formula.copy_patterns"]],
    ["formula.fill_pattern", formulaHostOperations.fillFormulaPattern, ["excel.formula.fill_down", "excel.formula.fill_right"]],
    ["formula.convert_to_values", formulaHostOperations.convertFormulasToValues, ["excel.formula.convert_to_values"]]
  ]),
  ...methodGroup("table", "TableHostOperations", "apps/excel-addin/src/host/table.test.ts", [
    ["table.list", tableHostOperations.listTables, ["excel.table.list"]],
    ["table.get_info", tableHostOperations.getTableInfo, ["excel.table.get_info", "excel.table.get_schema"]],
    ["table.read", tableHostOperations.readTable, ["excel.table.read_compact"]],
    ["table.create", tableHostOperations.createTable, ["excel.table.create"]],
    ["table.resize", tableHostOperations.resizeTable, ["excel.table.resize"]],
    ["table.reorder_columns", tableHostOperations.reorderTableColumns, ["excel.table.reorder_columns"]],
    ["table.append_rows", tableHostOperations.appendTableRows, ["excel.table.append_rows"]],
    ["table.update_rows", tableHostOperations.updateTableRows, ["excel.table.update_rows"]],
    ["table.clear_data_keep_formulas", tableHostOperations.clearTableDataKeepFormulas, ["excel.table.clear_data_keep_formulas"]],
    ["table.clear_filters", tableHostOperations.clearTableFilters, ["excel.table.clear_filters"]],
    ["table.apply_filters", tableHostOperations.applyTableFilters, ["excel.table.apply_filters"]],
    ["table.sort", tableHostOperations.sortTable, ["excel.table.sort"]],
    ["table.apply_view", tableHostOperations.applyTableView, ["excel.table.apply_view"]],
    ["table.clear_sort", tableHostOperations.clearTableSort, []],
    ["table.set_total_row", tableHostOperations.setTableTotalRow, ["excel.table.set_total_row"]],
    ["table.set_style", tableHostOperations.setTableStyle, ["excel.table.set_style"]],
    ["table.copy_structure", tableHostOperations.copyTableStructure, ["excel.table.copy_structure"]]
  ]),
  {
    method: "operation.execute_batch",
    implementationOwner: "BatchHostOperations",
    handler: batchHostOperations.executeBatch,
    relatedBackendCapabilities: [
      "excel.batch.apply",
      "excel.batch.submit",
      "excel.batch.submit_chunked",
      "excel.batch.preflight",
      "excel.batch.validate",
      "excel.batch.dry_run",
      "excel.range.write_values",
      "excel.range.write_values_many",
      "excel.range.write_formulas",
      "excel.range.write_number_formats",
      "excel.range.write_number_formats_many",
      "excel.range.write_styles",
      "excel.range.write_styles_many",
      "excel.range.write_data_validation",
      "excel.range.write_conditional_formatting",
      "excel.range.clear_style_dimensions",
      "excel.range.clear_style_dimensions_many",
      "excel.range.clear",
      "excel.range.clear_many",
      "excel.range.clear_values",
      "excel.range.clear_formats",
      "excel.range.clear_formats_many",
      "excel.range.clear_values_keep_format",
      "excel.range.copy",
      "excel.range.move",
      "excel.range.reorder_columns",
      "excel.range.insert_rows",
      "excel.range.delete_rows",
      "excel.range.insert_columns",
      "excel.range.delete_columns",
      "excel.range.autofit_columns",
      "excel.range.autofit_rows",
      "excel.range.autofit_many",
      "excel.range.merge",
      "excel.range.unmerge",
      "excel.sheet.create",
      "excel.sheet.copy",
      "excel.sheet.copy_clean_data_regions",
      "excel.sheet.rename",
      "excel.sheet.delete",
      "excel.sheet.hide",
      "excel.sheet.unhide",
      "excel.sheet.protect",
      "excel.sheet.unprotect",
      "excel.sheet.clear",
      "excel.sheet.set_tab_color",
      "excel.template.create_sheet_from_template"
    ],
    operationKinds: [...BATCH_OPERATION_KINDS],
    hostDependency: "office-js",
    unitTestFile: "apps/excel-addin/src/host/batch.test.ts"
  },
  ...methodGroup("template", "TemplateHostOperations", "apps/excel-addin/src/host/template.test.ts", [
    ["template.capture", templateHostOperations.captureTemplate, ["excel.template.register", "excel.template.detect_templates"]],
    ["template.capture_sheet", templateHostOperations.captureSheetFingerprint, ["excel.template.validate_sheet_against_template", "excel.style.get_fingerprint"]],
    ["template.repair", templateHostOperations.repairTemplateConsistency, ["excel.template.repair_sheet_from_template", "excel.style.repair_consistency"]]
  ]),
  ...methodGroup("style", "StyleHostOperations", "apps/excel-addin/src/host/style.test.ts", [
    ["style.capture_fingerprint", styleHostOperations.captureStyleFingerprint, ["excel.style.get_fingerprint", "excel.style.compare_fingerprint"]],
    ["style.copy_dimensions", styleHostOperations.copyStyleDimensions, ["excel.style.copy_from_template", "excel.style.copy_column_widths", "excel.style.copy_row_heights", "excel.style.copy_borders", "excel.style.copy_fills", "excel.style.copy_fonts", "excel.style.copy_alignment", "excel.style.copy_number_formats", "excel.style.copy_conditional_formatting", "excel.style.copy_data_validation"]],
    ["style.copy_dimensions_many", styleHostOperations.copyStyleDimensionsMany, ["excel.style.copy_from_template", "excel.style.copy_column_widths", "excel.style.copy_row_heights", "excel.style.copy_borders", "excel.style.copy_fills", "excel.style.copy_fonts", "excel.style.copy_alignment", "excel.style.copy_number_formats", "excel.style.copy_conditional_formatting", "excel.style.copy_data_validation"]]
  ])
];

export const HOST_METHODS = Object.fromEntries(HOST_METHOD_REGISTRY.map((entry) => [entry.method, entry.handler]));
export const HOST_METHOD_NAMES = HOST_METHOD_REGISTRY.map((entry) => entry.method);

export function getHostMethod(method: string): HostMethodDefinition | undefined {
  return HOST_METHOD_REGISTRY.find((entry) => entry.method === method);
}

function methodGroup(
  _prefix: string,
  implementationOwner: string,
  unitTestFile: string,
  entries: Array<[string, HostMethodDefinition["handler"], string[]]>
): HostMethodDefinition[] {
  return entries.map(([method, handler, relatedBackendCapabilities]) => ({
    method,
    implementationOwner,
    handler,
    relatedBackendCapabilities,
    operationKinds: [],
    hostDependency: "office-js",
    unitTestFile
  }));
}
