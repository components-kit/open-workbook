import { describe, expect, it } from "vitest";
import type { WorkbookId } from "@component-kit/open-workbook-protocol";
import { LockManager } from "./lock-manager.js";

const workbookId = "workbook_test" as WorkbookId;

describe("LockManager", () => {
  it("allows non-overlapping range locks on the same sheet", () => {
    const locks = new LockManager();
    const first = locks.acquire({
      workbookId,
      scopes: [{ type: "range", workbookId, sheetName: "Transactions", address: "A1:F20" }],
      mode: "write_values",
      reason: "clean transaction log"
    });
    expect(first.ok).toBe(true);

    const second = locks.acquire({
      workbookId,
      scopes: [{ type: "range", workbookId, sheetName: "Transactions", address: "H1:N20" }],
      mode: "write_formulas",
      reason: "create summary formulas"
    });
    expect(second.ok).toBe(true);
  });

  it("blocks overlapping range locks", () => {
    const locks = new LockManager();
    locks.acquire({
      workbookId,
      scopes: [{ type: "range", workbookId, sheetName: "Transactions", address: "A1:F20" }],
      mode: "write_values",
      reason: "clean transaction log"
    });

    const blocked = locks.acquire({
      workbookId,
      scopes: [{ type: "range", workbookId, sheetName: "Transactions", address: "D5:H10" }],
      mode: "write_values",
      reason: "overlapping update"
    });
    expect(blocked.ok).toBe(false);
    if (!blocked.ok) {
      expect(blocked.conflicts[0]?.code).toBe("LOCK_CONFLICT");
      expect(blocked.conflicts[0]?.lockId).toBeDefined();
      expect(blocked.conflicts[0]?.lockExpiresAt).toBeDefined();
    }
  });

  it("lets released locks be acquired again", () => {
    const locks = new LockManager();
    const acquired = locks.acquire({
      workbookId,
      scopes: [{ type: "sheet", workbookId, sheetName: "Dashboard" }],
      mode: "structure",
      reason: "repair dashboard sheet"
    });
    expect(acquired.ok).toBe(true);
    if (acquired.ok) {
      locks.release(acquired.locks.map((lock) => lock.lockId));
    }

    const next = locks.acquire({
      workbookId,
      scopes: [{ type: "range", workbookId, sheetName: "Dashboard", address: "A1:D10" }],
      mode: "write_values",
      reason: "write dashboard values"
    });
    expect(next.ok).toBe(true);
  });

  it("reports missing lock ids during renew and release", () => {
    const locks = new LockManager();
    const acquired = locks.acquire({
      workbookId,
      scopes: [{ type: "range", workbookId, sheetName: "Dashboard", address: "A1:D10" }],
      mode: "write_values",
      reason: "write dashboard values"
    });
    expect(acquired.ok).toBe(true);
    if (!acquired.ok) {
      return;
    }

    const renewed = locks.renewWithMissing([acquired.locks[0]!.lockId, "lock_missing" as any], 240_000);
    const released = locks.releaseWithMissing([acquired.locks[0]!.lockId, "lock_missing" as any]);

    expect(renewed.renewed).toHaveLength(1);
    expect(renewed.missingLockIds).toEqual(["lock_missing"]);
    expect(released.released).toHaveLength(1);
    expect(released.missingLockIds).toEqual(["lock_missing"]);
  });
});
