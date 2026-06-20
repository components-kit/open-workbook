import { describe, expect, it } from "vitest";
import type { WorkbookId } from "@components-kit/open-workbook-protocol";
import { AgentOperationStore } from "./agent-operation-store.js";

describe("AgentOperationStore", () => {
  it("expires previewed operations", () => {
    let now = 1_000;
    const store = new AgentOperationStore({ now: () => now, previewTtlMs: 100 });
    const operation = store.create({
      workbookContextId: "ctx_test",
      workbookId: "workbook_test" as WorkbookId,
      action: { kind: "batch", operations: [] },
      changes: [],
      summary: "Preview"
    });

    expect(store.get(String(operation.operationId))).toBeDefined();
    now = 1_101;

    expect(store.get(String(operation.operationId))).toBeUndefined();
  });

  it("keeps terminal output for retry until terminal TTL expires", () => {
    let now = 2_000;
    const store = new AgentOperationStore({ now: () => now, previewTtlMs: 100, terminalTtlMs: 200 });
    const operation = store.create({
      workbookContextId: "ctx_test",
      workbookId: "workbook_test" as WorkbookId,
      action: { kind: "batch", operations: [] },
      changes: [],
      summary: "Preview"
    });

    store.markCompleted(String(operation.operationId), {
      status: "SUCCESS",
      mode: "apply_update",
      workbookContextId: "ctx_test",
      summary: "Applied",
      proof: [],
      resourceLinks: [],
      nextAction: "answer_now",
      warnings: []
    });
    now = 2_199;
    expect(store.get(String(operation.operationId))?.terminalOutput?.status).toBe("SUCCESS");
    now = 2_201;
    expect(store.get(String(operation.operationId))).toBeUndefined();
  });

  it("dumps and restores pending and terminal operation records", () => {
    let now = 3_000;
    const first = new AgentOperationStore({ now: () => now, previewTtlMs: 500, terminalTtlMs: 1_000 });
    const pending = first.create({
      workbookContextId: "ctx_test",
      workbookId: "workbook_test" as WorkbookId,
      action: { kind: "batch", operations: [] },
      changes: [],
      summary: "Pending preview"
    });
    const terminal = first.create({
      workbookContextId: "ctx_test",
      workbookId: "workbook_test" as WorkbookId,
      action: { kind: "batch", operations: [] },
      changes: [],
      summary: "Terminal preview"
    });
    first.markCompleted(String(terminal.operationId), {
      status: "SUCCESS",
      mode: "apply_update",
      workbookContextId: "ctx_test",
      summary: "Applied",
      proof: [],
      resourceLinks: [],
      nextAction: "answer_now",
      warnings: []
    });

    const restored = new AgentOperationStore({ now: () => now });
    restored.load(first.dump());

    expect(restored.get(String(pending.operationId))?.summary).toBe("Pending preview");
    expect(restored.get(String(terminal.operationId))?.terminalOutput?.status).toBe("SUCCESS");

    now = 4_001;
    const filtered = new AgentOperationStore({ now: () => now });
    filtered.load(first.dump());
    expect(filtered.get(String(pending.operationId))).toBeUndefined();
    expect(filtered.get(String(terminal.operationId))).toBeUndefined();
  });
});
