import type {
  BatchRequest,
  CompiledBatch,
  DestructiveLevel,
  ExcelOperation,
  RangeFingerprint
} from "@component-kit/open-workbook-protocol";
import { createRangeFingerprint } from "./fingerprint.js";
import { cellCount } from "./range-address.js";

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

      const target = getOperationTarget(operation);
      if (target) {
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

function getOperationTarget(operation: ExcelOperation) {
  switch (operation.kind) {
    case "range.read_full":
    case "range.write_values":
    case "range.write_formulas":
    case "range.write_number_formats":
    case "range.write_styles":
    case "range.write_hyperlinks":
    case "range.write_comments":
    case "range.clear":
    case "range.clear_values":
    case "range.clear_formats":
    case "range.clear_values_keep_format":
    case "range.insert_rows":
    case "range.delete_rows":
    case "range.insert_columns":
    case "range.delete_columns":
    case "range.autofit_columns":
    case "range.autofit_rows":
    case "range.merge":
    case "range.unmerge":
    case "range.restore_snapshot":
      return operation.target;
    case "range.copy":
    case "range.move":
      return operation.source;
    case "template.create_sheet_from_template":
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
      return undefined;
  }
}
