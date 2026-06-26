import type {
  BatchRequest,
  CompiledBatch,
  DestructiveLevel,
  ExcelOperation,
  RangeFingerprint,
  A1Range
} from "@components-kit/open-workbook-protocol";
import { createRangeFingerprint } from "./fingerprint.js";
import { cellCount, columnNameToNumber, formatA1Address, parseA1Address, stripSheetName } from "./range-address.js";

const EXCEL_MAX_ROW = 1_048_576;
const EXCEL_MAX_COLUMN = 16_384;

const DESTRUCTIVE_RANK: Record<DestructiveLevel, number> = {
  none: 0,
  values: 1,
  format: 2,
  structure: 3,
  workbook: 4
};

export interface BatchCompilerOptions {
  now?: () => string;
}

export class BatchCompiler {
  constructor(private readonly options: BatchCompilerOptions = {}) {}

  compile(request: BatchRequest): CompiledBatch {
    const requiredBackups = new Set<"region" | "sheet" | "workbook-copy">();
    const targetFingerprints: RangeFingerprint[] = [];
    let estimatedCellsTouched = 0;
    let destructiveLevel: DestructiveLevel = "none";

    for (const operation of request.operations) {
      destructiveLevel = maxDestructiveLevel(destructiveLevel, operation.destructiveLevel);

      if (operation.kind.startsWith("range.")) {
        requiredBackups.add("region");
      }
      if (operation.kind === "sheet.delete") {
        requiredBackups.add("workbook-copy");
      }
      if (operation.kind === "sheet.copy_clean_data_regions") {
        requiredBackups.add("sheet");
        requiredBackups.add("workbook-copy");
      }
      if (operation.kind === "sheet.clear") {
        requiredBackups.add("sheet");
      }
      if (operation.kind === "template.create_sheet_from_template") {
        requiredBackups.add("sheet");
        requiredBackups.add("workbook-copy");
      }
      if (operation.destructiveLevel === "structure" || operation.destructiveLevel === "workbook") {
        requiredBackups.add("workbook-copy");
      }

      for (const target of getOperationTargets(operation)) {
        estimatedCellsTouched += cellCount(target.address);
        targetFingerprints.push(createRangeFingerprint(target, { pending: operation.kind }, this.now()));
      }
    }

    return {
      workbookId: request.workbookId,
      operations: request.operations,
      requiredBackups: [...requiredBackups],
      targetFingerprints,
      estimatedCellsTouched,
      destructiveLevel
    };
  }

  private now(): string {
    return this.options.now?.() ?? new Date().toISOString();
  }
}

function maxDestructiveLevel(left: DestructiveLevel, right: DestructiveLevel): DestructiveLevel {
  return DESTRUCTIVE_RANK[right] > DESTRUCTIVE_RANK[left] ? right : left;
}

function getOperationTargets(operation: ExcelOperation): A1Range[] {
  switch (operation.kind) {
    case "range.read_full":
    case "range.write_values":
    case "range.write_formulas":
    case "range.write_number_formats":
    case "range.write_styles":
    case "range.write_data_validation":
    case "range.write_conditional_formatting":
    case "range.clear_style_dimensions":
    case "range.write_hyperlinks":
    case "range.write_comments":
    case "range.clear":
    case "range.clear_values":
    case "range.clear_formats":
    case "range.clear_values_keep_format":
    case "range.autofit_columns":
    case "range.autofit_rows":
    case "range.apply_autofilter":
    case "range.clear_autofilter":
    case "range.merge":
    case "range.unmerge":
    case "range.restore_snapshot":
    case "range.reorder_columns":
      return operation.kind === "range.write_data_validation" && operation.entries?.length
        ? operation.entries.map((entry) => entry.target)
        : operation.target
          ? [operation.target]
          : [];
    case "range.write_values_many":
      return operation.entries.map((entry) => entry.target);
    case "range.write_number_formats_many":
      return operation.entries.map((entry) => entry.target);
    case "range.write_styles_many":
      return operation.entries.map((entry) => entry.target);
    case "range.clear_style_dimensions_many":
      return operation.entries.map((entry) => entry.target);
    case "range.clear_many":
      return operation.entries.map((entry) => entry.target);
    case "range.clear_formats_many":
      return operation.targets;
    case "range.autofit_many":
      return operation.entries.map((entry) => entry.target);
    case "range.insert_rows":
    case "range.delete_rows":
      return [shiftedRowRange(operation.target)];
    case "range.insert_columns":
    case "range.delete_columns":
      return [shiftedColumnRange(operation.target)];
    case "range.hide_columns":
    case "range.unhide_columns":
      return [shiftedRowRange(operation.target)];
    case "range.copy":
    case "range.move":
      return [operation.source, operation.target];
    case "template.create_sheet_from_template":
    case "sheet.copy_clean_data_regions":
    case "workbook.calculate":
    case "workbook.save":
    case "sheet.create":
    case "sheet.copy":
    case "sheet.rename":
    case "sheet.delete":
    case "sheet.move":
    case "sheet.hide":
    case "sheet.unhide":
    case "sheet.protect":
    case "sheet.unprotect":
    case "sheet.clear":
    case "sheet.set_tab_color":
    case "sheet.freeze_panes":
      return [];
    default:
      return [];
  }
}

function shiftedRowRange(target: A1Range): A1Range {
  const parsed = parseStructuralAddress(stripSheetName(target.address));
  return {
    ...target,
    address: formatA1Address({
      startRow: parsed.startRow,
      endRow: EXCEL_MAX_ROW,
      startColumn: parsed.startColumn,
      endColumn: parsed.endColumn
    })
  };
}

function shiftedColumnRange(target: A1Range): A1Range {
  const parsed = parseStructuralAddress(stripSheetName(target.address));
  return {
    ...target,
    address: formatA1Address({
      startRow: parsed.startRow,
      endRow: parsed.endRow,
      startColumn: parsed.startColumn,
      endColumn: EXCEL_MAX_COLUMN
    })
  };
}

function parseStructuralAddress(address: string) {
  const wholeColumn = /^([A-Z]+)(?::([A-Z]+))?$/i.exec(address);
  if (wholeColumn) {
    return {
      startRow: 1,
      endRow: EXCEL_MAX_ROW,
      startColumn: columnNameToNumber(wholeColumn[1]!),
      endColumn: columnNameToNumber(wholeColumn[2] ?? wholeColumn[1]!)
    };
  }
  const wholeRow = /^(\d+)(?::(\d+))?$/.exec(address);
  if (wholeRow) {
    return {
      startRow: Number(wholeRow[1]),
      endRow: Number(wholeRow[2] ?? wholeRow[1]),
      startColumn: 1,
      endColumn: EXCEL_MAX_COLUMN
    };
  }
  return parseA1Address(address);
}
