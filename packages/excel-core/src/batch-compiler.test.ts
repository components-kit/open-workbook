import { describe, expect, it } from "vitest";
import type { BatchRequest, OperationId, WorkbookId } from "@components-kit/open-workbook-protocol";
import { BatchCompiler } from "./batch-compiler.js";

describe("BatchCompiler", () => {
  it("requires region backups and estimates touched cells for range writes", () => {
    const workbookId = "workbook_test" as WorkbookId;
    const request: BatchRequest = {
      workbookId,
      mode: "dry_run",
      operations: [
        {
          kind: "range.write_values",
          operationId: "op_test" as OperationId,
          workbookId,
          destructiveLevel: "values",
          reason: "Write sample values",
          target: {
            workbookId,
            sheetName: "Sheet1",
            address: "A1:B2"
          },
          values: [
            [1, 2],
            [3, 4]
          ],
          preserveFormats: true
        }
      ]
    };

    const compiled = new BatchCompiler({ now: () => "2026-06-12T00:00:00.000Z" }).compile(request);

    expect(compiled.requiredBackups).toContain("region");
    expect(compiled.estimatedCellsTouched).toBe(4);
    expect(compiled.destructiveLevel).toBe("values");
    expect(compiled.targetFingerprints).toHaveLength(1);
  });

  it("tracks both source and destination ranges for copy and move operations", () => {
    const workbookId = "workbook_test" as WorkbookId;
    const request: BatchRequest = {
      workbookId,
      mode: "dry_run",
      operations: [
        {
          kind: "range.copy",
          operationId: "op_copy" as OperationId,
          workbookId,
          destructiveLevel: "values",
          reason: "Copy values",
          source: { workbookId, sheetName: "Sheet1", address: "A1:B2" },
          target: { workbookId, sheetName: "Sheet1", address: "D1:E2" },
          copyType: "all"
        },
        {
          kind: "range.move",
          operationId: "op_move" as OperationId,
          workbookId,
          destructiveLevel: "values",
          reason: "Move values",
          source: { workbookId, sheetName: "Sheet1", address: "A5:A6" },
          target: { workbookId, sheetName: "Sheet1", address: "C5:C6" }
        }
      ]
    };

    const compiled = new BatchCompiler({ now: () => "2026-06-12T00:00:00.000Z" }).compile(request);

    expect(compiled.requiredBackups).toContain("region");
    expect(compiled.estimatedCellsTouched).toBe(12);
    expect(compiled.targetFingerprints.map((fingerprint) => fingerprint.range.address)).toEqual(["A1:B2", "D1:E2", "A5:A6", "C5:C6"]);
  });

  it("tracks every target inside grouped range operations", () => {
    const workbookId = "workbook_test" as WorkbookId;
    const request: BatchRequest = {
      workbookId,
      mode: "dry_run",
      operations: [
        {
          kind: "range.write_values_many",
          operationId: "op_values_many" as OperationId,
          workbookId,
          destructiveLevel: "values",
          reason: "Write related ranges",
          entries: [
            { target: { workbookId, sheetName: "Sheet1", address: "A1:B2" }, values: [[1, 2], [3, 4]] },
            { target: { workbookId, sheetName: "Sheet1", address: "D1:D2" }, values: [[5], [6]] }
          ]
        },
        {
          kind: "range.write_styles_many",
          operationId: "op_styles_many" as OperationId,
          workbookId,
          destructiveLevel: "format",
          reason: "Style related ranges",
          entries: [
            { target: { workbookId, sheetName: "Sheet1", address: "A1:B1" }, style: { font: { bold: true } } },
            { target: { workbookId, sheetName: "Sheet1", address: "D1:D1" }, style: { fill: { color: "#4472C4" } } }
          ]
        },
        {
          kind: "range.autofit_many",
          operationId: "op_autofit_many" as OperationId,
          workbookId,
          destructiveLevel: "format",
          reason: "Autofit related ranges",
          entries: [
            { target: { workbookId, sheetName: "Sheet1", address: "A1:D20" }, dimension: "columns" }
          ]
        }
      ]
    };

    const compiled = new BatchCompiler({ now: () => "2026-06-12T00:00:00.000Z" }).compile(request);

    expect(compiled.requiredBackups).toContain("region");
    expect(compiled.destructiveLevel).toBe("format");
    expect(compiled.targetFingerprints.map((fingerprint) => fingerprint.range.address)).toEqual([
      "A1:B2",
      "D1:D2",
      "A1:B1",
      "D1:D1",
      "A1:D20"
    ]);
  });

  it("requires sheet and workbook-copy backups for copy-clean sheet operations", () => {
    const workbookId = "workbook_test" as WorkbookId;
    const request: BatchRequest = {
      workbookId,
      mode: "dry_run",
      operations: [
        {
          kind: "sheet.copy_clean_data_regions",
          operationId: "op_copy_clean" as OperationId,
          workbookId,
          destructiveLevel: "structure",
          reason: "Duplicate template and clear data regions",
          sourceSheetName: "Template",
          newSheetName: "Report",
          dataRegions: ["B2:B20", "D2:D20"],
          position: "after",
          relativeToSheetName: "Template",
          activate: true
        }
      ]
    };

    const compiled = new BatchCompiler({ now: () => "2026-06-12T00:00:00.000Z" }).compile(request);

    expect(compiled.requiredBackups).toEqual(expect.arrayContaining(["sheet", "workbook-copy"]));
    expect(compiled.destructiveLevel).toBe("structure");
    expect(compiled.targetFingerprints).toHaveLength(0);
  });

  it("expands shifted row and column operations to affected remainder ranges", () => {
    const workbookId = "workbook_test" as WorkbookId;
    const request: BatchRequest = {
      workbookId,
      mode: "dry_run",
      operations: [
        {
          kind: "range.insert_rows",
          operationId: "op_insert_rows" as OperationId,
          workbookId,
          destructiveLevel: "structure",
          reason: "Insert cells downward",
          target: { workbookId, sheetName: "Sheet1", address: "B3:D4" }
        },
        {
          kind: "range.delete_columns",
          operationId: "op_delete_columns" as OperationId,
          workbookId,
          destructiveLevel: "structure",
          reason: "Delete cells leftward",
          target: { workbookId, sheetName: "Sheet1", address: "F10:G12" }
        }
      ]
    };

    const compiled = new BatchCompiler({ now: () => "2026-06-12T00:00:00.000Z" }).compile(request);

    expect(compiled.requiredBackups).toContain("workbook-copy");
    expect(compiled.targetFingerprints.map((fingerprint) => fingerprint.range.address)).toEqual(["B3:D1048576", "F10:XFD12"]);
    expect(compiled.estimatedCellsTouched).toBe((1_048_576 - 3 + 1) * 3 + (16_384 - 6 + 1) * 3);
  });
});
