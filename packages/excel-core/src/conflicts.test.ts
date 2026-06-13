import { describe, expect, it } from "vitest";
import type { WorkbookId } from "@components-kit/open-workbook-protocol";
import { attachConflictGuidance, classifyScopeConflict, makeLockConflict } from "./conflicts.js";

const workbookId = "workbook_conflicts" as WorkbookId;

describe("classifyScopeConflict", () => {
  it("classifies table conflicts separately from generic range overlap", () => {
    const conflict = classifyScopeConflict(
      { type: "range", workbookId, sheetName: "Transactions", address: "A1:F20" },
      { type: "table", workbookId, sheetName: "Transactions", tableName: "TxTable" }
    );

    expect(conflict?.code).toBe("TABLE_CONFLICT");
  });

  it("classifies formula dependency conflicts", () => {
    const conflict = classifyScopeConflict(
      { type: "range", workbookId, sheetName: "Transactions", address: "A1:F20" },
      { type: "formula", workbookId, sheetName: "Transactions", address: "D1:D20" }
    );

    expect(conflict?.code).toBe("FORMULA_DEPENDENCY_CONFLICT");
  });

  it("classifies chart and pivot conflicts as derived object conflicts", () => {
    const chartConflict = classifyScopeConflict(
      { type: "chart", workbookId, sheetName: "Dashboard", chartName: "RevenueTrend" },
      { type: "range", workbookId, sheetName: "Dashboard", address: "A1:E20" }
    );
    const pivotConflict = classifyScopeConflict(
      { type: "pivot", workbookId, sheetName: "Dashboard", pivotName: "RevenuePivot" },
      { type: "range", workbookId, sheetName: "Dashboard", address: "A1:E20" }
    );

    expect(chartConflict?.code).toBe("DERIVED_OBJECT_CONFLICT");
    expect(pivotConflict?.code).toBe("DERIVED_OBJECT_CONFLICT");
  });

  it("recommends waiting and handoff for lock conflicts with owners", () => {
    const conflict = makeLockConflict({
      workbookId,
      left: { type: "range", workbookId, sheetName: "Transactions", address: "A1:F20" },
      right: { type: "range", workbookId, sheetName: "Transactions", address: "D1:H20" },
      taskId: "task_owner" as any
    });

    const guided = attachConflictGuidance({
      ...conflict!,
      lockId: "lock_owner" as any,
      lockExpiresAt: "2026-06-12T16:00:00.000Z"
    });

    expect(guided.guidance?.primaryAction).toBe("retry_after");
    expect(guided.guidance?.steps.some((step) => step.action === "handoff_task")).toBe(true);
    expect(guided.guidance?.steps.some((step) => step.action === "split_scope")).toBe(true);
  });

  it("recommends manual review for structure conflicts", () => {
    const conflict = makeLockConflict({
      workbookId,
      left: { type: "sheet", workbookId, sheetName: "Report" },
      right: { type: "range", workbookId, sheetName: "Report", address: "A1:D10" }
    });

    const guided = attachConflictGuidance(conflict!);

    expect(guided.guidance?.primaryAction).toBe("manual_review");
    expect(guided.guidance?.severity).toBe("blocked");
  });
});
