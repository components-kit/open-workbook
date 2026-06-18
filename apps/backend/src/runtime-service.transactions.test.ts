import { describe, expect, it } from "vitest";
import {
  NativeFileBridge,
  RuntimeService,
  appliedTransaction,
  existsSync,
  mkdirSync,
  mkdtempSync,
  operationOk,
  path,
  readFileSync,
  runtimeWithDynamicSnapshotHash,
  runtimeWithExecutingAddin,
  runtimeWithFormulaGraph,
  runtimeWithPersistentAddin,
  runtimeWithSnapshotHash,
  sleepForTest,
  snapshotResponse,
  tmpdir,
  writeFileSync,
  writeFormulaOperation,
  writeStyleOperation,
  writeValuesOperation
} from "./runtime-service.test-support.js";
import type { AgentId, OperationId, PlanId, RuntimeCapabilities, WorkbookId } from "./runtime-service.test-support.js";

describe("RuntimeService transaction rollback preview", () => {
  it("allows rollback preview when no later transaction overlaps", () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_rollback" as WorkbookId;
    const planId = "plan_rollback" as PlanId;
    const scopes: WorkbookScope[] = [{ type: "range", workbookId, sheetName: "Transactions", address: "A1:F20" }];
    const transaction = runtime.transactions.create({
      workbookId,
      planId,
      goal: "Clean transactions",
      scopes,
      destructiveLevel: "values"
    });
    runtime.transactions.markApplied(transaction.transactionId, { backups: [], warnings: [] });

    const preview = runtime.previewTransactionRollback(transaction.transactionId);

    expect(preview.ok).toBe(true);
    expect(preview.rollbackAvailable).toBe(true);
    expect(preview.rollbackMethod).toBe("plan");
  });

  it("blocks rollback when a later applied transaction overlaps the same scope", () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_rollback_conflict" as WorkbookId;
    const earlier = appliedTransaction(runtime, workbookId, "plan_earlier" as PlanId, [
      { type: "range", workbookId, sheetName: "Transactions", address: "A1:F20" }
    ]);
    appliedTransaction(runtime, workbookId, "plan_later" as PlanId, [
      { type: "range", workbookId, sheetName: "Transactions", address: "D1:H20" }
    ]);

    const preview = runtime.previewTransactionRollback(earlier.transactionId);

    expect(preview.ok).toBe(false);
    expect(preview.conflicts.some((conflict) => conflict.code === "ROLLBACK_CONFLICT")).toBe(true);
  });

  it("blocks rollback when transaction has no plan metadata", () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_no_plan" as WorkbookId;
    const transaction = appliedTransaction(runtime, workbookId, undefined, [
      { type: "range", workbookId, sheetName: "Transactions", address: "A1:F20" }
    ]);

    const preview = runtime.previewTransactionRollback(transaction.transactionId);

    expect(preview.ok).toBe(false);
    expect(preview.conflicts.some((conflict) => conflict.code === "ROLLBACK_UNAVAILABLE")).toBe(true);
  });

  it("reports dependency rollback conflict for later formula work", () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_formula_dependency" as WorkbookId;
    const earlier = appliedTransaction(runtime, workbookId, "plan_values" as PlanId, [
      { type: "range", workbookId, sheetName: "Transactions", address: "A1:F20" }
    ]);
    appliedTransaction(runtime, workbookId, "plan_formulas" as PlanId, [
      { type: "formula", workbookId, sheetName: "Transactions", address: "D1:D20" }
    ]);

    const preview = runtime.previewTransactionRollback(earlier.transactionId);

    expect(preview.ok).toBe(false);
    expect(preview.conflicts.some((conflict) => conflict.code === "ROLLBACK_DEPENDENCY_CONFLICT")).toBe(true);
  });

  it("previews dependent rollback chains newest first", () => {
    const runtime = new RuntimeService({ persistState: false });
    const workbookId = "workbook_rollback_chain" as WorkbookId;
    const earlier = appliedTransaction(runtime, workbookId, "plan_chain_earlier" as PlanId, [
      { type: "range", workbookId, sheetName: "Transactions", address: "A1:F20" }
    ]);
    const later = appliedTransaction(runtime, workbookId, "plan_chain_later" as PlanId, [
      { type: "range", workbookId, sheetName: "Transactions", address: "D1:H20" }
    ]);

    const preview = runtime.previewTransactionRollbackChain(earlier.transactionId);

    expect(preview.ok).toBe(true);
    expect(preview.requiresConfirmation).toBe(true);
    expect(preview.rollbackOrder.map((transaction) => transaction.transactionId)).toEqual([later.transactionId, earlier.transactionId]);
    expect(preview.confirmationToken).toContain(earlier.transactionId);
  });
});

describe("RuntimeService plan refresh", () => {
  it("refreshes preview fingerprints when target ranges are unchanged", async () => {
    const runtime = runtimeWithSnapshotHash("hash_same");
    const workbookId = "workbook_refresh" as WorkbookId;
    const plan = runtime.createPlan({
      workbookId,
      goal: "Write values",
      operations: [writeValuesOperation(workbookId)]
    });
    await runtime.previewPlan(plan.planId);

    const result = await runtime.refreshPlanPreview(plan.planId);

    expect(result.ok).toBe(true);
    expect(result.refreshed).toBe(true);
    expect(result.conflicts).toHaveLength(0);
  });

  it("blocks refresh when target ranges changed since preview", async () => {
    let hash = "hash_before";
    const runtime = runtimeWithDynamicSnapshotHash(() => hash);
    const workbookId = "workbook_refresh_conflict" as WorkbookId;
    const plan = runtime.createPlan({
      workbookId,
      goal: "Write values",
      operations: [writeValuesOperation(workbookId)]
    });
    await runtime.previewPlan(plan.planId);
    hash = "hash_after";

    const result = await runtime.refreshPlanPreview(plan.planId);

    expect(result.ok).toBe(false);
    expect(result.refreshed).toBe(false);
    expect(result.conflicts.some((conflict) => conflict.code === "TARGET_REGION_CHANGED")).toBe(true);
  });
});

describe("RuntimeService transaction progress", () => {
  it("reports queued transaction metadata and cancels queued work before Excel execution", async () => {
    const workbookId = "workbook_transaction_queue" as WorkbookId;
    const runtime = new RuntimeService({ persistState: false });
    const session = runtime.sessions.createSession();
    let executeCount = 0;
    let releaseFirstExecution!: () => void;
    const firstExecutionStarted = new Promise<void>((resolve) => {
      runtime.attachAddinClient(session.connectionId, {
        request: async (method: string, params: any) => {
          if (method === "workbook.snapshot_ranges") {
            return snapshotResponse(workbookId, params.ranges);
          }
          if (method === "operation.execute_batch") {
            executeCount += 1;
            if (executeCount === 1) {
              resolve();
              await new Promise<void>((release) => {
                releaseFirstExecution = release;
              });
            }
            return operationOk();
          }
          throw new Error(`Unexpected method ${method}`);
        }
      } as any);
    });

    const first = runtime.applyBatch({ workbookId, mode: "apply", operations: [writeValuesOperation(workbookId)] });
    await firstExecutionStarted;
    const second = runtime.submitBatch({ workbookId, mode: "apply", operations: [writeFormulaOperation(workbookId)] });
    await sleepForTest(10);
    const queued = runtime.getTransaction(second.transactionId!).transaction;

    expect(queued?.queuePosition).toBe(1);
    expect(queued?.progressMessage).toContain("queued");
    const cancelled = runtime.cancelTransaction(queued!.transactionId);
    releaseFirstExecution();
    const firstResult = await first;
    const waited = await runtime.waitTransaction(queued!.transactionId, 500);

    expect(cancelled.ok).toBe(true);
    expect(firstResult.ok).toBe(true);
    expect(waited.transaction?.status).toBe("cancelled");
    expect(executeCount).toBe(1);
  });

  it("returns queued progress for apply requests when another mutation is active", async () => {
    const workbookId = "workbook_transaction_busy_apply" as WorkbookId;
    const runtime = new RuntimeService({ persistState: false });
    const session = runtime.sessions.createSession();
    let releaseFirstExecution!: () => void;
    let executeCount = 0;
    const firstExecutionStarted = new Promise<void>((resolve) => {
      runtime.attachAddinClient(session.connectionId, {
        request: async (method: string, params: any) => {
          if (method === "workbook.snapshot_ranges") {
            return snapshotResponse(workbookId, params.ranges);
          }
          if (method === "operation.execute_batch") {
            executeCount += 1;
            if (executeCount === 1) {
              resolve();
              await new Promise<void>((release) => {
                releaseFirstExecution = release;
              });
            }
            return operationOk();
          }
          throw new Error(`Unexpected method ${method}`);
        }
      } as any);
    });

    const first = runtime.applyBatch({ workbookId, mode: "apply", operations: [writeValuesOperation(workbookId)] });
    await firstExecutionStarted;
    const second = await runtime.applyBatch({ workbookId, mode: "apply", operations: [writeFormulaOperation(workbookId)] });

    expect(second.ok).toBe(true);
    expect(second.transactionStatus).toBe("queued");
    expect(second.progressMessage).toContain("queued");
    releaseFirstExecution();
    await first;
    const waited = await runtime.waitTransaction(second.transactionId!, 500);

    expect(waited.completed).toBe(true);
    expect(waited.transaction?.status).toBe("applied");
  });

  it("waits for a transaction to reach terminal status", async () => {
    const workbookId = "workbook_transaction_wait" as WorkbookId;
    const runtime = runtimeWithExecutingAddin(workbookId);
    const applied = await runtime.applyBatch({ workbookId, mode: "apply", operations: [writeValuesOperation(workbookId)] });

    const waited = await runtime.waitTransaction(applied.transactionId!, 100);

    expect(waited.ok).toBe(true);
    expect(waited.completed).toBe(true);
    expect(waited.transaction?.status).toBe("applied");
  });

  it("submits a batch without waiting for Excel execution to finish", async () => {
    const workbookId = "workbook_transaction_submit" as WorkbookId;
    const runtime = new RuntimeService({ persistState: false });
    const session = runtime.sessions.createSession();
    let releaseExecution!: () => void;
    const executionStarted = new Promise<void>((resolve) => {
      runtime.attachAddinClient(session.connectionId, {
        request: async (method: string, params: any) => {
          if (method === "workbook.snapshot_ranges") {
            return snapshotResponse(workbookId, params.ranges);
          }
          if (method === "operation.execute_batch") {
            resolve();
            await new Promise<void>((release) => {
              releaseExecution = release;
            });
            return operationOk();
          }
          throw new Error(`Unexpected method ${method}`);
        }
      } as any);
    });

    const submitted = runtime.submitBatch({
      workbookId,
      mode: "apply",
      operations: [writeValuesOperation(workbookId)],
      retryStrategy: "split_style_entries",
      chunksTotal: 3,
      chunksCompleted: 1,
      progressMessage: "Queued style update chunk 2 of 3."
    });

    expect(submitted.ok).toBe(true);
    expect(submitted.status).toBe("queued");
    expect(submitted.transactionId).toBeTruthy();
    expect(submitted.transaction?.retryStrategy).toBe("split_style_entries");
    expect(submitted.transaction?.chunksTotal).toBe(3);
    expect(submitted.transaction?.chunksCompleted).toBe(1);
    await executionStarted;
    expect(runtime.getTransaction(submitted.transactionId!).transaction?.status).toBe("applying");
    releaseExecution();
    const waited = await runtime.waitTransaction(submitted.transactionId!, 500);

    expect(waited.completed).toBe(true);
    expect(waited.transaction?.status).toBe("applied");
  });

  it("preflights large style and matrix batches before execution", () => {
    const workbookId = "workbook_batch_preflight" as WorkbookId;
    const runtime = new RuntimeService({ persistState: false });
    const styleOperations = Array.from({ length: 30 }, (_, index) => writeStyleOperation(workbookId, `A${index + 1}:A${index + 1}`));
    const stylePreflight = runtime.preflightBatch({ workbookId, mode: "validate", operations: styleOperations });
    const values = Array.from({ length: 600 }, (_, index) => [index]);
    const matrixPreflight = runtime.preflightBatch({
      workbookId,
      mode: "validate",
      operations: [{
        kind: "range.write_values",
        operationId: "op_large_values" as OperationId,
        workbookId,
        destructiveLevel: "values",
        reason: "Write large values",
        target: { workbookId, sheetName: "Sheet1", address: "A1:A600" },
        values,
        preserveFormats: true
      }]
    });

    expect(stylePreflight.recommendedExecutionMode).toBe("chunked_submit");
    expect(stylePreflight.chunkPlan?.strategy).toBe("split_style_entries");
    expect(stylePreflight.chunkPlan?.chunksTotal).toBe(2);
    expect(matrixPreflight.recommendedExecutionMode).toBe("chunked_submit");
    expect(matrixPreflight.chunkPlan?.strategy).toBe("split_matrix_rows");
    expect(matrixPreflight.chunkPlan?.chunksTotal).toBe(2);
  });

  it("tracks chunked batch work as a parent job", async () => {
    const workbookId = "workbook_chunked_job" as WorkbookId;
    const runtime = runtimeWithExecutingAddin(workbookId);
    const operations = Array.from({ length: 30 }, (_, index) => writeStyleOperation(workbookId, `B${index + 1}:B${index + 1}`));

    const submitted = runtime.submitChunkedBatch({ workbookId, mode: "apply", operations }, { goal: "Apply report styles" });
    const waited = await runtime.waitJob(submitted.jobId, 1_000);

    expect(submitted.ok).toBe(true);
    expect(submitted.jobId).toBeTruthy();
    expect(submitted.transactionIds).toHaveLength(2);
    expect(waited.completed).toBe(true);
    expect(waited.job?.status).toBe("applied");
    expect(waited.job?.chunksCompleted).toBe(2);
  });

  it("cancels queued child transactions from a parent job", async () => {
    const workbookId = "workbook_chunked_job_cancel" as WorkbookId;
    const runtime = new RuntimeService({ persistState: false });
    const session = runtime.sessions.createSession();
    let releaseFirstExecution!: () => void;
    let executeCount = 0;
    const firstExecutionStarted = new Promise<void>((resolve) => {
      runtime.attachAddinClient(session.connectionId, {
        request: async (method: string, params: any) => {
          if (method === "workbook.snapshot_ranges") {
            return snapshotResponse(workbookId, params.ranges);
          }
          if (method === "operation.execute_batch") {
            executeCount += 1;
            if (executeCount === 1) {
              resolve();
              await new Promise<void>((release) => {
                releaseFirstExecution = release;
              });
            }
            return operationOk();
          }
          throw new Error(`Unexpected method ${method}`);
        }
      } as any);
    });
    const operations = Array.from({ length: 30 }, (_, index) => writeStyleOperation(workbookId, `C${index + 1}:C${index + 1}`));
    const submitted = runtime.submitChunkedBatch({ workbookId, mode: "apply", operations }, { goal: "Apply report styles" });
    await firstExecutionStarted;

    const cancelled = runtime.cancelJob(submitted.jobId);
    releaseFirstExecution();
    const waited = await runtime.waitJob(submitted.jobId, 1_000);

    expect(cancelled.ok).toBe(false);
    expect(cancelled.cancelledTransactions).toHaveLength(1);
    expect(waited.job?.status).toBe("partially_applied");
    expect(waited.job?.chunksCompleted).toBe(1);
    expect(executeCount).toBe(1);
  });

  it("retries timed-out style batches as smaller queued chunks", async () => {
    const workbookId = "workbook_transaction_style_retry" as WorkbookId;
    const runtime = new RuntimeService({ persistState: false });
    const session = runtime.sessions.createSession();
    let executeCount = 0;
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string, params: any) => {
        if (method === "workbook.snapshot_ranges") {
          return snapshotResponse(workbookId, params.ranges);
        }
        if (method === "operation.execute_batch") {
          executeCount += 1;
          if (executeCount === 1) {
            throw new Error("Timed out waiting for add-in method: operation.execute_batch");
          }
          return operationOk();
        }
        throw new Error(`Unexpected method ${method}`);
      }
    } as any);

    const result = await runtime.applyBatch({
      workbookId,
      mode: "apply",
      operations: [
        writeStyleOperation(workbookId, "A1:A1"),
        writeStyleOperation(workbookId, "A2:A2"),
        writeStyleOperation(workbookId, "A3:A3")
      ]
    });
    const retryIds = (result.data as any).retryTransactionIds as string[];

    expect(result.ok).toBe(true);
    expect(result.warnings.some((warning) => warning.code === "RETRYING_SMALLER_BATCH")).toBe(true);
    expect(retryIds).toHaveLength(2);
    for (const retryId of retryIds) {
      const waited = await runtime.waitTransaction(retryId as any, 500);
      expect(waited.transaction?.status).toBe("applied");
    }
    expect(executeCount).toBe(3);
  });
});
