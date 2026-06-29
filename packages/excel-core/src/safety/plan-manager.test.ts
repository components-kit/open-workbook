import { describe, expect, it } from "vitest";
import type { ExcelOperation, OperationId, WorkbookId } from "@components-kit/open-workbook-protocol";
import { PlanManager } from "./plan-manager.js";

describe("PlanManager", () => {
  it("counts copied sheets in plan preview summaries", () => {
    const workbookId = "workbook_plan_manager" as WorkbookId;
    const operations: ExcelOperation[] = [
      {
        kind: "sheet.copy",
        operationId: "op_copy" as OperationId,
        workbookId,
        destructiveLevel: "structure",
        reason: "Copy with data",
        sourceSheetName: "Template",
        newSheetName: "Template Data Copy"
      },
      {
        kind: "sheet.copy_clean_data_regions",
        operationId: "op_copy_clean" as OperationId,
        workbookId,
        destructiveLevel: "structure",
        reason: "Copy clean template",
        sourceSheetName: "Template",
        newSheetName: "Template Clean Copy",
        dataRegions: ["Template!A2:B10"]
      }
    ];

    const manager = new PlanManager();
    const plan = manager.createPlan({ workbookId, goal: "Copy templates", operations });
    const preview = manager.previewPlan(plan.planId);

    expect(preview.diffSummary.sheetsChanged).toBe(2);
  });
});
