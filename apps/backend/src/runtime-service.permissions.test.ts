import { describe, expect, it } from "vitest";
import { RuntimeService } from "./runtime-service.js";
import { operationOk, snapshotResponse, writeValuesOperation } from "./runtime-service.test-support.js";
import type { OperationId, WorkbookId } from "./runtime-service.test-support.js";

describe("RuntimeService permission contracts", () => {
  it("updates permission policy fields through dedicated helpers", () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_permissions_policy" as WorkbookId;

    runtime.setPermissions({ allowWrites: false, allowWorkbookActions: false });
    runtime.requireConfirmation(["values", "structure", "values"]);
    runtime.setPermissionScope({ workbookId, sheetNames: ["Allowed"] });
    runtime.allowDestructiveActions(false);
    runtime.allowMacroExecution(true);

    const permissions = runtime.getPermissions().permissions;
    expect(permissions.allowWrites).toBe(false);
    expect(permissions.allowWorkbookActions).toBe(false);
    expect(permissions.allowDestructiveActions).toBe(false);
    expect(permissions.allowMacroExecution).toBe(true);
    expect(permissions.requireConfirmationFor).toEqual(["values", "structure"]);
    expect(permissions.scope).toMatchObject({ workbookId, sheetNames: ["Allowed"] });
  });

  it("blocks writes, confirmations, workbook scope, sheet scope, destructive actions, and workbook actions", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_permissions_apply" as WorkbookId;
    const otherWorkbookId = "workbook_other" as WorkbookId;
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string, params: any) => {
        if (method === "workbook.snapshot_ranges") {
          return snapshotResponse(workbookId, params.ranges);
        }
        if (method === "operation.execute_batch") {
          return operationOk();
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);
    runtime.setPermissions({
      allowWrites: false,
      allowDestructiveActions: false,
      allowWorkbookActions: false,
      requireConfirmationFor: ["values"],
      scope: { workbookId: otherWorkbookId, sheetNames: ["Allowed"] }
    });

    const result = await runtime.applyBatch({
      workbookId,
      mode: "apply",
      operations: [
        {
          ...writeValuesOperation(workbookId),
          destructiveLevel: "workbook",
          target: { workbookId, sheetName: "Blocked", address: "A1" }
        }
      ]
    });

    expect(result.ok).toBe(false);
    expect(result.warnings.map((warning) => warning.code)).toEqual(expect.arrayContaining([
      "WRITES_DISABLED",
      "DESTRUCTIVE_ACTION_BLOCKED",
      "WORKBOOK_ACTION_BLOCKED",
      "WORKBOOK_SCOPE_BLOCKED",
      "SHEET_SCOPE_BLOCKED"
    ]));
  });

  it("allows confirmed scoped writes and blocks locked regions until unlocked", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_permissions_locked_region" as WorkbookId;
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string, params: any) => {
        if (method === "workbook.snapshot_ranges") {
          return snapshotResponse(workbookId, params.ranges);
        }
        if (method === "operation.execute_batch") {
          return operationOk();
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);
    await runtime.registerRegion({
      workbookId,
      name: "Output",
      sheetName: "Report",
      address: "B2:C3"
    });
    runtime.setPermissions({
      allowWrites: true,
      allowDestructiveActions: true,
      allowWorkbookActions: true,
      requireConfirmationFor: ["values"],
      scope: { workbookId, regionNames: ["Output"] }
    });

    const missingConfirmation = await runtime.applyBatch({
      workbookId,
      mode: "apply",
      operations: [permissionWriteOperation(workbookId, "Report", "B2")]
    });
    await runtime.lockRegions({ workbookId, regions: [{ regionName: "Output", reason: "Review in progress" }] });
    const locked = await runtime.applyBatch({
      workbookId,
      mode: "apply",
      confirmationToken: "confirmed",
      operations: [permissionWriteOperation(workbookId, "Report", "B2")]
    });
    runtime.unlockRegions({ workbookId, regionNames: ["Output"] });
    const allowed = await runtime.applyBatch({
      workbookId,
      mode: "apply",
      confirmationToken: "confirmed",
      operations: [permissionWriteOperation(workbookId, "Report", "B2")]
    });
    const outOfScope = await runtime.applyBatch({
      workbookId,
      mode: "apply",
      confirmationToken: "confirmed",
      operations: [permissionWriteOperation(workbookId, "Report", "D4")]
    });

    expect(missingConfirmation.ok).toBe(false);
    expect(missingConfirmation.warnings.some((warning) => warning.code === "CONFIRMATION_REQUIRED")).toBe(true);
    expect(locked.ok).toBe(false);
    expect(locked.warnings.some((warning) => warning.code === "LOCKED_REGION_BLOCKED")).toBe(true);
    expect(allowed.ok).toBe(true);
    expect(outOfScope.ok).toBe(false);
    expect(outOfScope.warnings.some((warning) => warning.code === "REGION_SCOPE_BLOCKED")).toBe(true);
    expect(runtime.getPermissions().permissions.lockedRegions).toHaveLength(0);
  });
});

function permissionWriteOperation(workbookId: WorkbookId, sheetName: string, address: string) {
  return {
    kind: "range.write_values" as const,
    operationId: `op_permissions_${sheetName}_${address}` as OperationId,
    workbookId,
    destructiveLevel: "values" as const,
    reason: "Write permission test value",
    target: { workbookId, sheetName, address },
    values: [["ok"]],
    preserveFormats: true
  };
}
