import { describe, expect, it } from "vitest";
import type { BatchRequest, OperationId, WorkbookId } from "@open-workbook/protocol";
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
});
