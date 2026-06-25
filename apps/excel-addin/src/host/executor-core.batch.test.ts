import { BatchCompiler } from "@components-kit/open-workbook-excel-core";
import type { BatchRequest, ExcelOperation, WorkbookId } from "@components-kit/open-workbook-protocol";
import { readFileSync } from "node:fs";
import { afterEach, describe, expect, it } from "vitest";
import { executeBatch } from "./executor-core.js";

const workbookId = "workbook_executor_test" as WorkbookId;
const EXPLICITLY_UNSUPPORTED_OPERATION_KINDS = new Set(["range.write_comments", "range.write_hyperlinks", "sheet.move"]);

describe("Office.js batch executor production operations", () => {
  afterEach(() => {
    delete (globalThis as { Excel?: unknown }).Excel;
  });

  it("executes style, validation, conditional formatting, insert-column, and reorder-column operations", async () => {
    const fixture = installExcelFixture();
    const operations: ExcelOperation[] = [
      {
        kind: "range.write_styles",
        operationId: "op_style",
        workbookId,
        destructiveLevel: "format",
        reason: "Style header",
        target: range("Sales", "A1:E1"),
        preserveValues: true,
        style: { fillColor: "#000000", fontColor: "#FFFFFF", fontBold: true, horizontalAlignment: "center" }
      },
      {
        kind: "range.write_styles_many",
        operationId: "op_style_many",
        workbookId,
        destructiveLevel: "format",
        reason: "Style grouped headers",
        entries: [
          { target: range("Sales", "A1:A1"), style: { fontBold: true } },
          { target: range("Sales", "E1:E1"), style: { fillColor: "#4472C4" } }
        ]
      },
      {
        kind: "range.insert_columns",
        operationId: "op_insert_column",
        workbookId,
        destructiveLevel: "structure",
        reason: "Insert reviewer column",
        target: range("Sales", "F:F")
      },
      {
        kind: "range.reorder_columns",
        operationId: "op_reorder_columns",
        workbookId,
        destructiveLevel: "structure",
        reason: "Swap first two columns",
        target: range("Sales", "A1:B6"),
        columnOrder: [2, 1]
      },
      {
        kind: "range.write_data_validation",
        operationId: "op_validation",
        workbookId,
        destructiveLevel: "format",
        reason: "Add status dropdown",
        target: range("Sales", "E2:E6"),
        validation: { type: "list", source: ["Open", "Reviewed", "Closed"], inCellDropDown: true }
      },
      {
        kind: "range.write_conditional_formatting",
        operationId: "op_conditional_format",
        workbookId,
        destructiveLevel: "format",
        reason: "Highlight open rows",
        target: range("Sales", "A2:E6"),
        rule: { type: "custom", formula: '=$E2="Open"', style: { fillColor: "#FFFF00", fontBold: true } }
      }
    ];
    const request: BatchRequest = { workbookId, mode: "apply", operations };
    const compiled = new BatchCompiler({ now: () => "2026-06-21T00:00:00.000Z" }).compile(request);

    const result = await executeBatch({ request, compiled });

    expect(result.ok).toBe(true);
    expect(result.warnings).toEqual([]);
    expect(result.diffSummary?.destructiveLevel).toBe("structure");
    expect(result.telemetry).toMatchObject({ syncCount: 2, rangeCount: compiled.targetFingerprints.length, warningCount: 0 });
    expect(fixture.syncCount).toBe(2);
    expect(fixture.calls).toEqual(expect.arrayContaining([
      { type: "getRange", sheetName: "Sales", address: "A1:E1" },
      { type: "style", address: "A1:E1", property: "fill.color", value: "#000000" },
      { type: "style", address: "A1:E1", property: "font.color", value: "#FFFFFF" },
      { type: "style", address: "A1:E1", property: "font.bold", value: true },
      { type: "style", address: "A1:E1", property: "horizontalAlignment", value: "center" },
      { type: "insert", address: "F:F", shift: "right" },
      { type: "dataValidation.rule", address: "E2:E6", source: "Open,Reviewed,Closed" },
      { type: "conditionalFormats.add", address: "A2:E6", formatType: "custom" },
      { type: "conditionalRule", address: "A2:E6", formula: '=$E2="Open"' },
      { type: "conditionalStyle", address: "A2:E6", property: "fill.color", value: "#FFFF00" },
      { type: "conditionalStyle", address: "A2:E6", property: "font.bold", value: true }
    ]));
    expect(fixture.calls.some((call) => call.type === "worksheet.add" && call.sheetName.startsWith("__owb_reorder_"))).toBe(true);
    expect(fixture.calls.some((call) => call.type === "copyFrom" && call.address === "A1:B6" && String(call.source).includes("__owb_reorder_"))).toBe(true);
    expect(fixture.calls.some((call) => call.type === "worksheet.delete" && String(call.sheetName).includes("__owb_reorder_"))).toBe(true);
  });

  it("returns an explicit warning for host-limited operations instead of silent success", async () => {
    installExcelFixture();
    const operations: ExcelOperation[] = [
      {
        kind: "range.write_comments",
        operationId: "op_comments",
        workbookId,
        destructiveLevel: "values",
        reason: "Write comments",
        target: range("Sales", "A1:A1"),
        comments: [["Needs review"]]
      }
    ];
    const request: BatchRequest = { workbookId, mode: "apply", operations };
    const compiled = new BatchCompiler({ now: () => "2026-06-21T00:00:00.000Z" }).compile(request);

    const result = await executeBatch({ request, compiled });

    expect(result.ok).toBe(false);
    expect(result.warnings).toEqual([
      expect.objectContaining({
        code: "OPERATION_NOT_SUPPORTED",
        message: expect.stringContaining("range.write_comments")
      })
    ]);
    expect(result.error).toBeUndefined();
    expect(result.telemetry?.warningCount).toBe(1);
  });

  it("returns expected and received matrix dimensions for grouped value write failures", async () => {
    installExcelFixture();
    const operations: ExcelOperation[] = [
      op({ kind: "range.write_values_many", entries: [{ target: range("Sales", "'Sales'!B2:C2"), values: [["Only one value"]], preserveFormats: true }] })
    ];
    const request: BatchRequest = { workbookId, mode: "apply", operations };
    const compiled = new BatchCompiler({ now: () => "2026-06-21T00:00:00.000Z" }).compile(request);

    const result = await executeBatch({ request, compiled });

    expect(result.ok).toBe(false);
    expect((result.error as any)?.message).toContain("Matrix dimensions do not match Sales!B2:C2");
    expect((result.error as any)?.message).toContain("expected 1 row(s) x 2 column(s), received 1 row(s) x 1 column(s)");
  });

  it("executes the remaining supported batch operations through Office.js APIs", async () => {
    const fixture = installExcelFixture();
    const operations: ExcelOperation[] = [
      op({ kind: "range.write_values", target: range("Ops", "A1:B2"), values: [["A", "B"], [1, 2]], preserveFormats: true }),
      op({ kind: "range.write_values_many", entries: [{ target: range("Ops", "C1:C2"), values: [["C"], [3]], preserveFormats: true }] }),
      op({ kind: "range.write_formulas", target: range("Ops", "D2:D2"), formulas: [["=A2+B2"]], preserveFormats: true }),
      op({ kind: "range.write_number_formats", target: range("Ops", "B2:B2"), numberFormat: [["$#,##0"]], preserveValues: true }),
      op({ kind: "range.write_number_formats_many", entries: [{ target: range("Ops", "C2:C2"), numberFormat: [["0.0"]], preserveValues: true }] }),
      op({ kind: "range.clear_style_dimensions", target: range("Ops", "A1:B2"), dimensions: ["fills", "fonts", "alignment", "numberFormats", "rowHeights", "columnWidths"] }),
      op({ kind: "range.clear_style_dimensions_many", entries: [{ target: range("Ops", "C1:C2"), dimensions: ["fills"] }] }),
      op({ kind: "range.clear_values_keep_format", target: range("Ops", "A2:A2") }),
      op({ kind: "range.clear", target: range("Ops", "B2:B2"), applyTo: "all" }),
      op({ kind: "range.clear_many", entries: [{ target: range("Ops", "C2:C2"), applyTo: "contents" }] }),
      op({ kind: "range.clear_values", target: range("Ops", "D2:D2") }),
      op({ kind: "range.clear_formats", target: range("Ops", "E2:E2") }),
      op({ kind: "range.clear_formats_many", targets: [range("Ops", "F2:F2")] }),
      op({ kind: "range.copy", source: range("Ops", "A1:B2"), target: range("Ops", "H1:I2"), copyType: "values" }),
      op({ kind: "range.move", source: range("Ops", "H1:I2"), target: range("Ops", "J1:K2") }),
      op({ kind: "range.insert_rows", target: range("Ops", "3:3") }),
      op({ kind: "range.delete_rows", target: range("Ops", "4:4") }),
      op({ kind: "range.delete_columns", target: range("Ops", "L:L") }),
      op({ kind: "range.hide_columns", target: range("Ops", "M:N") }),
      op({ kind: "range.unhide_columns", target: range("Ops", "O:P") }),
      op({ kind: "range.autofit_columns", target: range("Ops", "A1:B10") }),
      op({ kind: "range.autofit_rows", target: range("Ops", "A1:A2") }),
      op({ kind: "range.autofit_many", entries: [{ target: range("Ops", "A1:B10"), dimension: "both" }] }),
      op({ kind: "range.apply_autofilter", target: range("Ops", "A1:B10") }),
      op({ kind: "range.clear_autofilter", target: range("Ops", "A1:B10") }),
      op({ kind: "range.merge", target: range("Ops", "M1:N1"), across: false }),
      op({ kind: "range.unmerge", target: range("Ops", "M1:N1") }),
      op({
        kind: "range.restore_snapshot",
        target: range("Ops", "O1:P1"),
        snapshot: {
          fingerprint: { range: range("Ops", "O1:P1"), hash: "hash", cellCount: 2, capturedAt: "2026-06-21T00:00:00.000Z" },
          values: [["Restored", 42]],
          numberFormat: [["@", "0"]],
          style: { fillColor: "#D9EAD3", fontColor: "#274E13", fontBold: true }
        }
      }),
      op({ kind: "workbook.calculate", calculationType: "full" }),
      op({ kind: "workbook.save" }),
      op({ kind: "sheet.create", sheetName: "Created", activate: true }),
      op({ kind: "sheet.copy", sourceSheetName: "Ops", newSheetName: "Ops Copy", position: "after", activate: true }),
      op({ kind: "sheet.copy_clean_data_regions", sourceSheetName: "Ops", newSheetName: "Ops Clean", dataRegions: ["Ops!A2:B3"], position: "after", activate: true }),
      op({ kind: "sheet.rename", sheetName: "Created", newSheetName: "Created Renamed" }),
      op({ kind: "sheet.hide", sheetName: "Ops Copy" }),
      op({ kind: "sheet.unhide", sheetName: "Ops Copy" }),
      op({ kind: "sheet.protect", sheetName: "Ops", password: "secret", options: { allowAutoFilter: true, allowSort: true } }),
      op({ kind: "sheet.unprotect", sheetName: "Ops", password: "secret" }),
      op({ kind: "sheet.clear", sheetName: "Ops", applyTo: "contents" }),
      op({ kind: "sheet.set_tab_color", sheetName: "Ops", color: "#00B050" }),
      op({ kind: "template.create_sheet_from_template", templateId: "template_ops" as any, newSheetName: "From Template", clearDataRegions: true }),
      op({ kind: "sheet.delete", sheetName: "Ops Clean" })
    ];
    const request: BatchRequest = { workbookId, mode: "apply", operations };
    const compiled = new BatchCompiler({ now: () => "2026-06-21T00:00:00.000Z" }).compile(request);

    const result = await executeBatch({
      request,
      compiled,
      templateSources: [{ templateId: "template_ops" as any, sourceSheetName: "Ops", dataRegions: ["Ops!A2:B3"] }]
    });

    expect(result.ok).toBe(true);
    expect(result.warnings).toEqual([]);
    expect(fixture.calls).toEqual(expect.arrayContaining([
      { type: "set", address: "Ops:0:2", property: "values", value: [["A", "B"], [1, 2]] },
      { type: "set", address: "Ops:1:4", property: "formulas", value: [["=A2+B2"]] },
      { type: "set", address: "Ops:1:2", property: "numberFormat", value: [["$#,##0"]] },
      { type: "clear", address: "A2:A2", applyTo: "contents" },
      { type: "copyFrom", address: "H1:I2", source: "A1:B2", copyType: "values" },
      { type: "delete", address: "L:L", shift: "left" },
      { type: "set", address: "M:N", property: "columnHidden", value: true },
      { type: "set", address: "O:P", property: "columnHidden", value: false },
      { type: "autofitColumns", address: "A1:B10" },
      { type: "autofitRows", address: "A1:A2" },
      { type: "autoFilter.apply", address: "A1:B10", range: "A1:B10" },
      { type: "autoFilter.clearCriteria", address: "A1:B10" },
      { type: "merge", address: "M1:N1", across: false },
      { type: "unmerge", address: "M1:N1" },
      { type: "application.calculate", calculationType: "full" },
      { type: "workbook.save", behavior: "save" },
      { type: "worksheet.add", sheetName: "Created" },
      { type: "worksheet.copy", sheetName: "Ops", position: "after", relativeTo: "Ops" },
      { type: "worksheet.name", sheetName: "Ops", value: "Ops Copy" },
      { type: "worksheet.activate", sheetName: "Ops Copy" },
      { type: "worksheet.visibility", sheetName: "Ops Copy", value: "hidden" },
      { type: "worksheet.visibility", sheetName: "Ops Copy", value: "visible" },
      { type: "worksheet.protect", sheetName: "Ops", options: { allowAutoFilter: true, allowSort: true }, password: "secret" },
      { type: "worksheet.unprotect", sheetName: "Ops", password: "secret" },
      { type: "getUsedRangeOrNullObject", sheetName: "Ops" },
      { type: "worksheet.tabColor", sheetName: "Ops", value: "#00B050" },
      { type: "worksheet.delete", sheetName: "Ops Clean" }
    ]));
  });

  it("treats sheet clear on an empty sheet as a no-op", async () => {
    const fixture = installExcelFixture({ emptyUsedRangeSheets: ["Blank"] });
    const operations: ExcelOperation[] = [
      op({ kind: "sheet.clear", sheetName: "Blank", applyTo: "all" })
    ];
    const request: BatchRequest = { workbookId, mode: "apply", operations };
    const compiled = new BatchCompiler({ now: () => "2026-06-21T00:00:00.000Z" }).compile(request);

    const result = await executeBatch({ request, compiled });

    expect(result.ok).toBe(true);
    expect(result.warnings).toEqual([]);
    expect(fixture.calls).toEqual(expect.arrayContaining([
      { type: "getUsedRangeOrNullObject", sheetName: "Blank" },
      { type: "load", address: "Blank:used", propertyPath: "isNullObject" }
    ]));
    expect(fixture.calls).not.toEqual(expect.arrayContaining([
      { type: "clear", address: "Blank:used", applyTo: "all" }
    ]));
  });

  it("keeps executor behavior fixtures aligned with every protocol operation kind", () => {
    const protocolKinds = operationKinds(readFileSync(new URL("../../../../packages/protocol/src/operations.ts", import.meta.url), "utf8"));
    const covered = new Set([
      ...BEHAVIOR_COVERED_OPERATION_KINDS,
      ...EXPLICITLY_UNSUPPORTED_OPERATION_KINDS
    ]);

    expect([...protocolKinds].filter((kind) => !covered.has(kind)).sort()).toEqual([]);
  });
});

function range(sheetName: string, address: string) {
  return { workbookId, sheetName, address };
}

function op<T extends Omit<ExcelOperation, "operationId" | "workbookId" | "destructiveLevel" | "reason">>(operation: T): ExcelOperation {
  return {
    operationId: `op_${operation.kind}` as any,
    workbookId,
    destructiveLevel: operation.kind.startsWith("sheet.") || operation.kind === "template.create_sheet_from_template" ? "structure" : "values",
    reason: operation.kind,
    ...operation
  } as ExcelOperation;
}

const BEHAVIOR_COVERED_OPERATION_KINDS = new Set([
  "range.read_full",
  "range.write_values",
  "range.write_values_many",
  "range.write_formulas",
  "range.clear_values_keep_format",
  "range.write_number_formats",
  "range.write_number_formats_many",
  "range.write_styles",
  "range.write_styles_many",
  "range.write_data_validation",
  "range.write_conditional_formatting",
  "range.clear",
  "range.clear_many",
  "range.clear_values",
  "range.clear_formats",
  "range.clear_formats_many",
  "range.clear_style_dimensions",
  "range.clear_style_dimensions_many",
  "range.copy",
  "range.move",
  "range.reorder_columns",
  "range.insert_rows",
  "range.delete_rows",
  "range.insert_columns",
  "range.delete_columns",
  "range.hide_columns",
  "range.unhide_columns",
  "range.autofit_columns",
  "range.autofit_rows",
  "range.autofit_many",
  "range.apply_autofilter",
  "range.clear_autofilter",
  "range.merge",
  "range.unmerge",
  "range.restore_snapshot",
  "sheet.create",
  "sheet.copy",
  "sheet.copy_clean_data_regions",
  "sheet.rename",
  "sheet.delete",
  "sheet.hide",
  "sheet.unhide",
  "sheet.protect",
  "sheet.unprotect",
  "sheet.clear",
  "sheet.set_tab_color",
  "workbook.calculate",
  "workbook.save",
  "template.create_sheet_from_template"
]);

function operationKinds(source: string): Set<string> {
  const operationDefinitions = source.slice(0, source.indexOf("export type ExcelOperation"));
  return new Set([...operationDefinitions.matchAll(/kind:\s*"([^"]+)"/g)].map((match) => match[1]!));
}

function installExcelFixture(options: { emptyUsedRangeSheets?: string[] } = {}) {
  const fixture = new ExcelFixture(options);
  (globalThis as { Excel?: unknown }).Excel = {
    run: async (callback: (context: FakeContext) => Promise<unknown>) => callback(fixture.context),
    ClearApplyTo: { all: "all", contents: "contents", formats: "formats" },
    ConditionalFormatType: { custom: "custom" },
    DataValidationAlertStyle: { stop: "stop", warning: "warning", information: "information" },
    DeleteShiftDirection: { left: "left", up: "up" },
    InsertShiftDirection: { right: "right", down: "down" },
    RangeCopyType: { all: "all", values: "values", formats: "formats", formulas: "formulas" },
    SaveBehavior: { save: "save" },
    CalculationType: { recalculate: "recalculate", full: "full" },
    WorksheetPositionType: { after: "after", before: "before" },
    SheetVisibility: { hidden: "hidden", visible: "visible" },
    HorizontalAlignment: { general: "general" },
    VerticalAlignment: { bottom: "bottom" },
    RangeUnderlineStyle: { none: "none" },
    BorderLineStyle: { none: "none", continuous: "continuous", dash: "dash", dashDot: "dashDot", dashDotDot: "dashDotDot", dot: "dot", double: "double", slantDashDot: "slantDashDot" },
    BorderWeight: { hairline: "hairline", thin: "thin", medium: "medium", thick: "thick" }
  };
  return fixture;
}

class ExcelFixture {
  readonly calls: Array<Record<string, unknown>> = [];
  syncCount = 0;
  readonly context = new FakeContext(this);
  readonly emptyUsedRangeSheets: Set<string>;

  constructor(options: { emptyUsedRangeSheets?: string[] } = {}) {
    this.emptyUsedRangeSheets = new Set(options.emptyUsedRangeSheets ?? []);
  }
}

class FakeContext {
  readonly workbook: FakeWorkbook;

  constructor(private readonly fixture: ExcelFixture) {
    this.workbook = new FakeWorkbook(fixture);
  }

  async sync() {
    this.fixture.syncCount += 1;
  }
}

class FakeWorkbook {
  readonly worksheets: FakeWorksheets;
  readonly application = {
    suspendApiCalculationUntilNextSync: () => undefined,
    suspendScreenUpdatingUntilNextSync: () => undefined,
    calculate: (calculationType: string) => this.fixture.calls.push({ type: "application.calculate", calculationType })
  };

  constructor(private readonly fixture: ExcelFixture) {
    this.worksheets = new FakeWorksheets(fixture);
  }

  save(behavior?: string) {
    this.fixture.calls.push({ type: "workbook.save", behavior });
  }
}

class FakeWorksheets {
  private readonly sheets = new Map<string, FakeWorksheet>();

  constructor(private readonly fixture: ExcelFixture) {}

  getItem(sheetName: string) {
    if (!this.sheets.has(sheetName)) {
      this.sheets.set(sheetName, new FakeWorksheet(this.fixture, sheetName));
    }
    return this.sheets.get(sheetName)!;
  }

  add(sheetName: string) {
    this.fixture.calls.push({ type: "worksheet.add", sheetName });
    const sheet = new FakeWorksheet(this.fixture, sheetName);
    this.sheets.set(sheetName, sheet);
    return sheet;
  }
}

class FakeWorksheet {
  readonly protection = {
    protect: (options: unknown, password?: string) => this.fixture.calls.push({ type: "worksheet.protect", sheetName: this.name, options, password }),
    unprotect: (password?: string) => this.fixture.calls.push({ type: "worksheet.unprotect", sheetName: this.name, password })
  };

  constructor(private readonly fixture: ExcelFixture, private sheetName: string) {}

  get name() {
    return this.sheetName;
  }

  set name(value: string) {
    this.fixture.calls.push({ type: "worksheet.name", sheetName: this.sheetName, value });
    this.sheetName = value;
  }

  set visibility(value: string) {
    this.fixture.calls.push({ type: "worksheet.visibility", sheetName: this.name, value });
  }

  set tabColor(value: string) {
    this.fixture.calls.push({ type: "worksheet.tabColor", sheetName: this.name, value });
  }

  getRange(address: string) {
    this.fixture.calls.push({ type: "getRange", sheetName: this.name, address });
    return new FakeRange(this.fixture, address);
  }

  getRangeByIndexes(rowIndex: number, columnIndex: number, rowCount: number, columnCount: number) {
    const address = `${this.name}:${rowIndex}:${columnIndex + columnCount}`;
    this.fixture.calls.push({ type: "getRangeByIndexes", sheetName: this.name, rowIndex, columnIndex, rowCount, columnCount });
    return new FakeRange(this.fixture, address, rowCount, columnCount);
  }

  delete() {
    this.fixture.calls.push({ type: "worksheet.delete", sheetName: this.name });
  }

  activate() {
    this.fixture.calls.push({ type: "worksheet.activate", sheetName: this.name });
  }

  copy(position: string, relativeTo: FakeWorksheet) {
    this.fixture.calls.push({ type: "worksheet.copy", sheetName: this.name, position, relativeTo: relativeTo.name });
    return new FakeWorksheet(this.fixture, this.name);
  }

  getUsedRangeOrNullObject() {
    this.fixture.calls.push({ type: "getUsedRangeOrNullObject", sheetName: this.name });
    return new FakeRange(this.fixture, `${this.name}:used`, undefined, undefined, this.fixture.emptyUsedRangeSheets.has(this.name));
  }
}

class FakeRange {
  readonly format: FakeRangeFormat;
  readonly dataValidation: FakeDataValidation;
  readonly conditionalFormats: FakeConditionalFormats;
  readonly worksheet = {
    autoFilter: {
      apply: (range: FakeRange) => this.fixture.calls.push({ type: "autoFilter.apply", address: this.address, range: range.address }),
      clearCriteria: () => this.fixture.calls.push({ type: "autoFilter.clearCriteria", address: this.address })
    }
  };

  constructor(
    private readonly fixture: ExcelFixture,
    readonly address: string,
    readonly rowCount = rangeShape(address).rowCount,
    readonly columnCount = rangeShape(address).columnCount,
    readonly isNullObject = false
  ) {
    this.format = new FakeRangeFormat(fixture, address);
    this.dataValidation = new FakeDataValidation(fixture, address);
    this.conditionalFormats = new FakeConditionalFormats(fixture, address);
  }

  load(propertyPath: string) {
    this.fixture.calls.push({ type: "load", address: this.address, propertyPath });
  }

  set values(value: unknown[][]) {
    this.fixture.calls.push({ type: "set", address: this.address, property: "values", value });
  }

  set formulas(value: unknown[][]) {
    this.fixture.calls.push({ type: "set", address: this.address, property: "formulas", value });
  }

  set numberFormat(value: unknown[][]) {
    this.fixture.calls.push({ type: "set", address: this.address, property: "numberFormat", value });
  }

  set columnHidden(value: boolean) {
    this.fixture.calls.push({ type: "set", address: this.address, property: "columnHidden", value });
  }

  getEntireColumn() {
    this.fixture.calls.push({ type: "getEntireColumn", address: this.address });
    return this;
  }

  getEntireRow() {
    this.fixture.calls.push({ type: "getEntireRow", address: this.address });
    return this;
  }

  insert(shift: string) {
    this.fixture.calls.push({ type: "insert", address: this.address, shift });
  }

  delete(shift: string) {
    this.fixture.calls.push({ type: "delete", address: this.address, shift });
  }

  getColumn(index: number) {
    this.fixture.calls.push({ type: "getColumn", address: this.address, index });
    return new FakeRange(this.fixture, `${this.address}:col${index}`, this.rowCount, 1);
  }

  copyFrom(source: FakeRange, copyType?: string) {
    this.fixture.calls.push({ type: "copyFrom", address: this.address, source: source.address, ...(copyType ? { copyType } : {}) });
  }

  clear(applyTo: string) {
    this.fixture.calls.push({ type: "clear", address: this.address, applyTo });
  }

  merge(across: boolean) {
    this.fixture.calls.push({ type: "merge", address: this.address, across });
  }

  unmerge() {
    this.fixture.calls.push({ type: "unmerge", address: this.address });
  }
}

class FakeRangeFormat {
  readonly fill: Record<string, unknown>;
  readonly font: Record<string, unknown>;
  readonly borders: { getItem: (edge: string) => Record<string, unknown> };

  constructor(private readonly fixture: ExcelFixture, private readonly address: string) {
    this.fill = new Proxy(propertyRecorder(this.fixture, this.address, "style", "fill"), {
      get: (target, property) => property === "clear"
        ? () => this.fixture.calls.push({ type: "style", address: this.address, property: "fill.clear" })
        : Reflect.get(target, property)
    }) as Record<string, unknown>;
    this.font = propertyRecorder(this.fixture, this.address, "style", "font");
    this.borders = {
      getItem: (edge: string) => propertyRecorder(this.fixture, this.address, "border", edge)
    };
  }

  set horizontalAlignment(value: string) {
    this.fixture.calls.push({ type: "style", address: this.address, property: "horizontalAlignment", value });
  }

  set verticalAlignment(value: string) {
    this.fixture.calls.push({ type: "style", address: this.address, property: "verticalAlignment", value });
  }

  autofitColumns() {
    this.fixture.calls.push({ type: "autofitColumns", address: this.address });
  }

  autofitRows() {
    this.fixture.calls.push({ type: "autofitRows", address: this.address });
  }
}

class FakeDataValidation {
  constructor(private readonly fixture: ExcelFixture, private readonly address: string) {}

  set rule(value: { list?: { source?: string } }) {
    this.fixture.calls.push({ type: "dataValidation.rule", address: this.address, source: value.list?.source });
  }

  set ignoreBlanks(value: boolean) {
    this.fixture.calls.push({ type: "dataValidation.ignoreBlanks", address: this.address, value });
  }
}

class FakeConditionalFormats {
  constructor(private readonly fixture: ExcelFixture, private readonly address: string) {}

  add(formatType: string) {
    this.fixture.calls.push({ type: "conditionalFormats.add", address: this.address, formatType });
    return new FakeConditionalFormat(this.fixture, this.address);
  }
}

class FakeConditionalFormat {
  readonly custom: { format: FakeConditionalFormatStyle; rule?: { formula: string } };

  constructor(fixture: ExcelFixture, address: string) {
    const format = new FakeConditionalFormatStyle(fixture, address);
    this.custom = new Proxy({ format }, {
      set(target, property, value) {
        if (property === "rule") {
          fixture.calls.push({ type: "conditionalRule", address, formula: (value as { formula?: string }).formula });
        }
        return Reflect.set(target, property, value);
      }
    }) as { format: FakeConditionalFormatStyle; rule?: { formula: string } };
  }
}

class FakeConditionalFormatStyle {
  readonly fill: Record<string, unknown>;
  readonly font: Record<string, unknown>;

  constructor(private readonly fixture: ExcelFixture, private readonly address: string) {
    this.fill = propertyRecorder(this.fixture, this.address, "conditionalStyle", "fill");
    this.font = propertyRecorder(this.fixture, this.address, "conditionalStyle", "font");
  }
}

function propertyRecorder(fixture: ExcelFixture, address: string, type: string, prefix: string) {
  return new Proxy({}, {
    set(target, property, value) {
      fixture.calls.push({ type, address, property: `${prefix}.${String(property)}`, value });
      return Reflect.set(target, property, value);
    }
  }) as Record<string, unknown>;
}

function rangeShape(address: string) {
  const rowOnly = /^(\d+)(?::(\d+))?$/.exec(address);
  if (rowOnly) return { rowCount: Number(rowOnly[2] ?? rowOnly[1]) - Number(rowOnly[1]) + 1, columnCount: 1 };
  const match = /^([A-Z]+)(\d+)?(?::([A-Z]+)(\d+)?)?$/.exec(address);
  if (!match) return { rowCount: 1, columnCount: 1 };
  const startColumn = columnNumber(match[1]!);
  const endColumn = columnNumber(match[3] ?? match[1]!);
  const startRow = Number(match[2] ?? 1);
  const endRow = Number(match[4] ?? match[2] ?? startRow);
  return {
    rowCount: endRow - startRow + 1,
    columnCount: endColumn - startColumn + 1
  };
}

function columnNumber(columnName: string) {
  let value = 0;
  for (const char of columnName) {
    value = value * 26 + (char.charCodeAt(0) - 64);
  }
  return value;
}
