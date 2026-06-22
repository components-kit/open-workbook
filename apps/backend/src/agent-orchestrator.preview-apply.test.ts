import { describe, expect, it } from "vitest";
import { AgentOrchestrator } from "./agent-orchestrator.js";
import { FakeAgentRuntime, createCachedMetadata, selectionInfo, sheets } from "./agent-orchestrator.test-support.js";

describe("AgentOrchestrator Preview Apply Safety", () => {
  it("requires structured values for write previews even when request text includes rows", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: `Add 2 mockup rows to the Report sheet at range Report!A3:B4.

Data rows:
["A", "B"]
["C", "D"]`,
        mode: "preview_update",
        intent: { action: "write_values" },
        target: { sheetName: "Report", range: "Report!A3:B4" }
      });

      expect(preview.status).toBe("NEEDS_INPUT");
      expect(preview.summary).toContain("Preview needs structured values");
      expect(preview.summary).toContain("rows embedded in request text");
      expect(preview.nextAction).toBe("ask_user");
      expect(preview.warnings).toContain("Mutation payloads must be supplied in the structured values field, such as values.values, values.rows, or values.patches.");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("reports missing structured values before target ambiguity for write previews", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Write rows into the report",
        mode: "preview_update",
        intent: { action: "write_values" }
      });

      expect(preview.status).toBe("NEEDS_INPUT");
      expect(preview.summary).toContain("Preview needs structured values");
      expect(preview.nextAction).toBe("ask_user");
      expect(preview.candidates?.length).toBeGreaterThan(0);
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("requires preview confirmation token before apply", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
  
      const preview = await agent.run({
        request: "Update Data B2",
        mode: "preview_update",
        target: { sheetName: "Data", range: "B2" },
        values: { values: [[999]] }
      });
      const missingToken = await agent.run({
        request: "Apply update",
        mode: "apply_update",
        operationId: preview.operationId
      });
  
      expect(preview.status).toBe("PREVIEW_READY");
      expect(preview.confirmationToken).toBeTruthy();
      expect(missingToken.status).toBe("NEEDS_INPUT");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("rejects wrong confirmation tokens without applying", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Update Data B2",
        mode: "preview_update",
        target: { sheetName: "Data", range: "B2" },
        values: { values: [[999]] }
      });
      const wrongToken = await agent.run({
        request: "Apply update",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: "confirm_wrong"
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect(wrongToken.status).toBe("NEEDS_INPUT");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("returns terminal apply output on repeated apply without rerunning the batch", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Update Data B2",
        mode: "preview_update",
        target: { sheetName: "Data", range: "B2" },
        values: { values: [[999]] }
      });
      const first = await agent.run({
        request: "Apply update",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });
      const second = await agent.run({
        request: "Apply update again",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });
      const status = await agent.run({
        request: "Check applied operation",
        mode: "operation_status",
        operationId: preview.operationId
      });
      const cancelApplied = await agent.run({
        request: "Cancel applied operation",
        mode: "cancel_operation",
        operationId: preview.operationId
      });

      expect(first.status).toBe("SUCCESS");
      expect(second).toMatchObject({ status: "SUCCESS", operationId: preview.operationId });
      expect(runtime.writeBatchCount).toBe(1);
      expect((status.answer as any).applyStatus).toBe("applied");
      expect(cancelApplied.status).toBe("VALIDATION_FAILED");
    });

  it("invalidates workbook context and stored result handles after successful apply", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_apply_invalidate");
      agent.metadataCache.set(metadata);

      const read = await agent.run({
        request: "Read first transaction rows",
        mode: "answer",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "read_values" },
        target: { tableName: "Transactions" },
        values: { rowLimit: 2 }
      });
      const resultUri = (read.answer as any).resultUri as string;
      const resultId = resultUri.split("/").pop()!;
      const preview = await agent.run({
        request: "Update Data B2",
        mode: "preview_update",
        workbookContextId: metadata.workbookContextId,
        target: { sheetName: "Data", range: "B2" },
        values: { values: [[999]] }
      });
      const applied = await agent.run({
        request: "Apply update",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(applied.status).toBe("SUCCESS");
      expect(applied.invalidatedContextIds).toContain(metadata.workbookContextId);
      expect(applied.invalidatedResourceUris).toContain(resultUri);
      expect(agent.getResultResource(resultId) as any).toMatchObject({ ok: false });
      expect(agent.getContextResource(metadata.workbookContextId) as any).toMatchObject({ ok: false });
    });

  it("reports and cancels previewed operations without workbook readiness", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Update Data B2",
        mode: "preview_update",
        target: { sheetName: "Data", range: "B2" },
        values: { values: [[999]] }
      });
      const status = await agent.run({
        request: "Check operation",
        mode: "operation_status",
        operationId: preview.operationId
      });
      const cancelled = await agent.run({
        request: "Cancel operation",
        mode: "cancel_operation",
        operationId: preview.operationId
      });
      const applyAfterCancel = await agent.run({
        request: "Apply cancelled operation",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(status.status).toBe("SUCCESS");
      expect((status.answer as any).applyStatus).toBe("previewed");
      expect(cancelled.status).toBe("SUCCESS");
      expect(applyAfterCancel.status).toBe("NOT_FOUND");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("previews and applies grouped range patches as one batch operation", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
  
      const preview = await agent.run({
        request: "Update related pricing zones",
        mode: "preview_update",
        values: {
          patches: [
            { target: { sheetName: "Data", range: "B2:C2" }, values: [[9000, 9000]], reason: "Zone A price" },
            { target: { sheetName: "Data", range: "B3:C3" }, values: [[6000, 6000]], reason: "Zone B price" }
          ]
        }
      });
      const applied = await agent.run({
        request: "Apply grouped pricing update",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });
  
      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe("multi_range_preview");
      expect((preview.answer as any).patchCount).toBe(2);
      expect(preview.summary).toContain("Apply this grouped preview once");
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.writeBatchCount).toBe(1);
      expect(runtime.lastBatchOperations).toHaveLength(1);
      expect(runtime.lastBatchOperations[0]?.kind).toBe("range.write_values_many");
      expect((runtime.lastBatchOperations[0] as any).entries.map((entry: any) => entry.target.address)).toEqual(["B2:C2", "B3:C3"]);
    });

  it("summarizes large value previews without reading every before-cell", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);
    const metadata = createCachedMetadata("wbctx_large_preview");
    agent.metadataCache.set(metadata);
    const values = Array.from({ length: 21 }, (_row, rowIndex) =>
      Array.from({ length: 10 }, (_cell, columnIndex) => `R${rowIndex + 1}C${columnIndex + 1}`)
    );
    const readCountBefore = runtime.readBatchCount;

    const preview = await agent.run({
      request: "Write large filled matrix",
      mode: "preview_update",
      workbookContextId: metadata.workbookContextId,
      target: { sheetName: "Report", range: "A1:J21" },
      values: { values }
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect(preview.changes).toHaveLength(1);
    expect((preview.changes?.[0]?.after as any).kind).toBe("range_write_summary");
    expect((preview.changes?.[0]?.after as any).cellCount).toBe(210);
    expect(runtime.readBatchCount).toBe(readCountBefore);
  });

  it("caps detailed grouped patch preview reads across small patches", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);
    const values = Array.from({ length: 1 }, () => Array.from({ length: 50 }, (_cell, index) => index));
    const preview = await agent.run({
      request: "Update related wide rows",
      mode: "preview_update",
      values: {
        patches: Array.from({ length: 5 }, (_patch, index) => ({
          target: { sheetName: "Data", range: `A${index + 1}:AX${index + 1}` },
          values,
          reason: `Patch ${index + 1}`
        }))
      }
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect(preview.changes?.length).toBeLessThanOrEqual(201);
  });

  it("applies a pending preview under the agent identity that created it", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      runtime.agentExecutionContext = { agentId: "agent_author", agentName: "Author", clientType: "mcp" };
      const preview = await agent.run({
        request: "Update Data B2",
        mode: "preview_update",
        target: { sheetName: "Data", range: "B2" },
        values: { values: [[999]] }
      });

      runtime.agentExecutionContext = { agentId: "agent_other", agentName: "Other", clientType: "mcp" };
      const applied = await agent.run({
        request: "Apply update",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchRequest?.agentId).toBe("agent_author");
    });

  it("blocks unsafe grouped patches before creating a pending batch", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
  
      const result = await agent.run({
        request: "Update related cells",
        mode: "preview_update",
        values: {
          patches: [
            { target: { sheetName: "Data", range: "B2:C2" }, values: [[1, 2]] },
            { target: { sheetName: "Data", range: "B3:C3" }, values: [["=SUM(A1:A2)", 3]] }
          ]
        }
      });
  
      expect(result.status).toBe("VALIDATION_FAILED");
      expect(result.summary).toContain("formula-like");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("treats apply results without explicit ok as success when validation passes", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.omitOkOnWrite = true;
      const agent = new AgentOrchestrator(runtime as any);
  
      const preview = await agent.run({
        request: "Update Data B2",
        mode: "preview_update",
        target: { sheetName: "Data", range: "B2" },
        values: { values: [[999]] }
      });
      const applied = await agent.run({
        request: "Apply update",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });
  
      expect(applied.status).toBe("SUCCESS");
      expect((applied.answer as any).ok).toBe(true);
      expect((applied.answer as any).validationOk).toBe(true);
    });

  it("uses parsed A1 overlap checks for formula-protected regions", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_formula_overlap");
      metadata.formulaRegions = [{ id: "formula:manual", sheetName: "Report", range: "A10:A20", formulaCount: 11 }];
      agent.metadataCache.set(metadata);
  
      const allowed = await agent.run({
        request: "Update report input",
        mode: "preview_update",
        workbookContextId: metadata.workbookContextId,
        target: { sheetName: "Report", range: "B1" },
        values: { values: [["ok"]] }
      });
      const formulaPreview = await agent.run({
        request: "Update formula area",
        mode: "preview_update",
        workbookContextId: metadata.workbookContextId,
        target: { sheetName: "Report", range: "A12" },
        values: { values: [["bad"]] }
      });
  
      expect(allowed.status).toBe("PREVIEW_READY");
      expect(formulaPreview.status).toBe("PREVIEW_READY");
      expect(formulaPreview.warnings).toContain("Target overlaps detected formula regions. Review carefully before applying.");
    });

  it("returns previews for scoped auto value edits by default", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
  
      const result = await agent.run({
        request: "Change Data B2 to 999",
        target: { sheetName: "Data", range: "B2" },
        values: { values: [[999]] }
      });
  
      expect(result.status).toBe("PREVIEW_READY");
      expect(result.mode).toBe("auto");
      expect(result.confirmationToken).toBeTruthy();
      expect(result.telemetry.autoApplied).toBeUndefined();
      expect(result.telemetry.safetyDecision).toBe("manual_review:auto_apply_disabled");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("previews small add requests with explicit values and range", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
  
      const result = await agent.run({
        request: "Add a notes block to Report",
        target: { sheetName: "Report", range: "B1:B2" },
        values: { values: [["Owner"], ["Finance"]] }
      });
  
      expect(result.status).toBe("PREVIEW_READY");
      expect(result.telemetry.safetyDecision).toBe("manual_review:auto_apply_disabled");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("previews header formatting as a style mutation instead of reading values", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
  
      const result = await agent.run({
        request: "Format the header row on the Data sheet.",
        target: { sheetName: "Data", range: "A1:D1" }
      });
  
      expect(result.status).toBe("PREVIEW_READY");
      expect((result.answer as any).kind).toBe("style_preview");
      expect(result.proof[0]?.range).toBe("A1:D1");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("previews formula writes through a formula-aware path", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
  
      const result = await agent.run({
        request: "Duplicate the formula from A12 down",
        mode: "preview_update",
        target: { sheetName: "Report", range: "A13:A14" },
        values: { values: [["=SUM(B1:B10)"], ["=SUM(B2:B11)"]] }
      });
  
      expect(result.status).toBe("PREVIEW_READY");
      expect((result.answer as any).kind).toBe("formula_preview");
      expect(result.nextAction).toBe("call_apply_update");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("previews and applies number-format writes through the agent surface", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Format Data C2:C4 as currency",
      mode: "preview_update",
      intent: { action: "write_number_formats" },
      target: { sheetName: "Data", range: "C2:C4" },
      values: { numberFormat: "$#,##0.00" }
    });
    const applied = await agent.run({
      request: "Apply number formats",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("number_format_preview");
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastWriteOperations[0]?.kind).toBe("range.write_number_formats");
    expect((runtime.lastWriteOperations[0] as any).numberFormat).toEqual([["$#,##0.00"], ["$#,##0.00"], ["$#,##0.00"]]);
  });

  it("previews clear-format and autofit-row range operations", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const clearFormats = await agent.run({
      request: "Clear formatting from Data A1:D1",
      mode: "preview_update",
      target: { sheetName: "Data", range: "A1:D1" }
    });
    const autofitRows = await agent.run({
      request: "Autofit row height for Data A1:D4",
      mode: "preview_update",
      target: { sheetName: "Data", range: "A1:D4" }
    });

    expect(clearFormats.status).toBe("PREVIEW_READY");
    expect((clearFormats.answer as any).kind).toBe("clear_formats_preview");
    expect(clearFormats.telemetry.operationRisk).toBe("safe_format");
    expect(autofitRows.status).toBe("PREVIEW_READY");
    expect((autofitRows.answer as any).dimension).toBe("rows");
      expect(autofitRows.telemetry.actionHandlerId).toBe("autofit_rows");
    });

  it("routes border-only removal to style-dimension clearing", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Remove all borders from selected cells",
      mode: "preview_update",
      target: { sheetName: "Data", range: "A3:B10" }
    });
    const applied = await agent.run({
      request: "Apply border cleanup",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("clear_style_dimensions_preview");
    expect((preview.answer as any).dimensions).toEqual(["borders"]);
    expect(preview.telemetry.safetyFingerprintOnly).toBe(true);
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastWriteOperations[0]?.kind).toBe("range.clear_style_dimensions");
    expect((runtime.lastWriteOperations[0] as any).dimensions).toEqual(["borders"]);
  });

  it("routes add-border requests to style writes with border payloads", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Add thin border to every cell in row 2",
      mode: "preview_update",
      target: { sheetName: "Data", range: "A2:D2" }
    });
    const applied = await agent.run({
      request: "Apply border style",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("style_preview");
    expect((preview.answer as any).style.borders).toEqual({ style: "continuous", weight: "thin" });
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastWriteOperations[0]?.kind).toBe("range.write_styles");
    expect((runtime.lastWriteOperations[0] as any).style.borders).toEqual({ style: "continuous", weight: "thin" });
  });

  it("previews and applies explicit range copy operations", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Copy range values",
      mode: "preview_update",
      intent: { action: "copy_range" },
      values: {
        source: { sheetName: "Data", range: "A1:B2" },
        destination: { sheetName: "Report", range: "A1:B2" },
        copyType: "values"
      }
    });
    const applied = await agent.run({
      request: "Apply copy",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("range_copy_preview");
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastWriteOperations[0]?.kind).toBe("range.copy");
    expect(runtime.lastBatchRequest?.operations[0]?.destructiveLevel).toBe("values");
    expect((runtime.lastWriteOperations[0] as any).source.address).toBe("A1:B2");
    expect(runtime.lastWriteOperations[0]?.target.address).toBe("A1:B2");
  });

  it("treats full range clear as structural risk", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Clear range including values and formats",
      mode: "preview_update",
      intent: { action: "clear_range" },
      target: { sheetName: "Data", range: "D4:E5" }
    });
    await agent.run({
      request: "Apply clear",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect(runtime.lastWriteOperations[0]?.kind).toBe("range.clear");
    expect(runtime.lastBatchRequest?.operations[0]?.destructiveLevel).toBe("structure");
  });

  it("uses conservative destructive levels for range copy-all and move operations", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const copyPreview = await agent.run({
      request: "Copy range including formulas and formats",
      mode: "preview_update",
      intent: { action: "copy_range" },
      values: {
        source: { sheetName: "Data", range: "A1:B2" },
        destination: { sheetName: "Report", range: "A1:B2" },
        copyType: "all"
      }
    });
    await agent.run({ request: "Apply copy", mode: "apply_update", operationId: copyPreview.operationId, confirmationToken: copyPreview.confirmationToken });
    expect(runtime.lastBatchRequest?.operations[0]?.destructiveLevel).toBe("structure");

    const movePreview = await agent.run({
      request: "Move range",
      mode: "preview_update",
      intent: { action: "move_range" },
      values: {
        source: { sheetName: "Data", range: "A1:B2" },
        destination: { sheetName: "Report", range: "C1:D2" }
      }
    });
    await agent.run({ request: "Apply move", mode: "apply_update", operationId: movePreview.operationId, confirmationToken: movePreview.confirmationToken });
    expect(runtime.lastBatchRequest?.operations[0]?.destructiveLevel).toBe("structure");
  });

  it("previews and applies safety artifact metadata mutations through the agent state machine", async () => {
    const cases: Array<{
      action: "refresh_snapshot" | "invalidate_snapshot" | "delete_snapshot" | "create_file_backup" | "restore_file_backup" | "prune_backups" | "pin_backup" | "unpin_backup" | "delete_backup";
      values: Record<string, unknown>;
      expectedKind: string;
      expectedRuntimeCall: string;
    }> = [
      { action: "refresh_snapshot", values: { snapshotId: "snapshot_agent_unit" }, expectedKind: "snapshot.refresh_preview", expectedRuntimeCall: "snapshot.refresh" },
      { action: "invalidate_snapshot", values: { snapshotId: "snapshot_agent_unit" }, expectedKind: "snapshot.invalidate_preview", expectedRuntimeCall: "snapshot.invalidate" },
      { action: "delete_snapshot", values: { snapshotId: "snapshot_agent_unit" }, expectedKind: "snapshot.delete_preview", expectedRuntimeCall: "snapshot.delete" },
      { action: "create_file_backup", values: { reason: "Before risky edit", pin: true }, expectedKind: "backup_create_file_preview", expectedRuntimeCall: "backup.create_file" },
      { action: "restore_file_backup", values: { backupId: "backup_file_unit", mode: "open-as-new" }, expectedKind: "backup_restore_file_preview", expectedRuntimeCall: "backup.restore_file" },
      { action: "prune_backups", values: { kind: "file-copy", maxBackupsPerWorkbook: 2 }, expectedKind: "backup_prune_preview", expectedRuntimeCall: "backup.prune" },
      { action: "pin_backup", values: { backupId: "backup_agent_unit" }, expectedKind: "backup.pin_preview", expectedRuntimeCall: "backup.pin" },
      { action: "unpin_backup", values: { backupId: "backup_agent_unit" }, expectedKind: "backup.unpin_preview", expectedRuntimeCall: "backup.unpin" },
      { action: "delete_backup", values: { backupId: "backup_agent_unit" }, expectedKind: "backup.delete_preview", expectedRuntimeCall: "backup.delete" }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: `Prepare ${testCase.action}`,
        mode: "preview_update",
        intent: { action: testCase.action },
        values: testCase.values
      });
      const applied = await agent.run({
        request: `Apply ${testCase.action}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe(testCase.expectedKind);
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.runtimeMethodCalls[testCase.expectedRuntimeCall]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    }
  });

  it("previews and applies workbook lifecycle mutations through backend orchestration", async () => {
    const localConfig = { version: 1, workbookId: "workbook_agent_unit", exportedAt: "2026-06-18T00:00:00.000Z", source: "open-workbook-local-config", templates: [], regions: [] };
    const cases: Array<{
      action: "restore_workbook_backup" | "import_local_config" | "embed_local_config" | "import_embedded_local_config" | "close_workbook";
      values: Record<string, unknown>;
      expectedKind: string;
      expectedRuntimeCall: string;
    }> = [
      { action: "restore_workbook_backup", values: { backupId: "backup_agent_unit" }, expectedKind: "workbook_restore_backup_preview", expectedRuntimeCall: "workbook.restore_backup" },
      { action: "import_local_config", values: { config: localConfig }, expectedKind: "workbook_import_local_config_preview", expectedRuntimeCall: "workbook.import_local_config" },
      { action: "embed_local_config", values: { includePermissions: true }, expectedKind: "workbook_embed_local_config_preview", expectedRuntimeCall: "workbook.embed_local_config" },
      { action: "import_embedded_local_config", values: { overwrite: true }, expectedKind: "workbook_import_embedded_local_config_preview", expectedRuntimeCall: "workbook.import_embedded_local_config" },
      { action: "close_workbook", values: { closeBehavior: "SkipSave" }, expectedKind: "workbook_close_preview", expectedRuntimeCall: "workbook.close" }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const preview = await agent.run({
        request: `Prepare ${testCase.action}`,
        mode: "preview_update",
        intent: { action: testCase.action },
        values: testCase.values
      });
      const applied = await agent.run({
        request: `Apply ${testCase.action}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe(testCase.expectedKind);
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.runtimeMethodCalls[testCase.expectedRuntimeCall]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    }
  });

  it("previews and applies formula operation intents through backend orchestration", async () => {
    const cases: Array<{
      action: "copy_formula_patterns" | "fill_formula_down" | "fill_formula_right" | "repair_formula_patterns" | "convert_formulas_to_values";
      expectedKind: string;
      expectedRuntimeCall: string;
      input: Parameters<AgentOrchestrator["run"]>[0];
    }> = [
      {
        action: "copy_formula_patterns",
        expectedKind: "formula_copy_patterns_preview",
        expectedRuntimeCall: "formula.copy_patterns",
        input: {
          request: "Copy formula patterns",
          mode: "preview_update",
          intent: { action: "copy_formula_patterns" },
          values: {
            source: { sheetName: "Report", range: "B2" },
            destination: { sheetName: "Report", range: "B3:B5" }
          }
        }
      },
      {
        action: "fill_formula_down",
        expectedKind: "formula_fill_down_preview",
        expectedRuntimeCall: "formula.fill_down",
        input: {
          request: "Fill formula down",
          mode: "preview_update",
          intent: { action: "fill_formula_down" },
          values: {
            source: { sheetName: "Report", range: "B2" },
            destination: { sheetName: "Report", range: "B3:B5" }
          }
        }
      },
      {
        action: "fill_formula_right",
        expectedKind: "formula_fill_right_preview",
        expectedRuntimeCall: "formula.fill_right",
        input: {
          request: "Fill formula right",
          mode: "preview_update",
          intent: { action: "fill_formula_right" },
          values: {
            source: { sheetName: "Report", range: "B2" },
            destination: { sheetName: "Report", range: "C2:E2" }
          }
        }
      },
      {
        action: "repair_formula_patterns",
        expectedKind: "formula_repair_patterns_preview",
        expectedRuntimeCall: "formula.repair_patterns",
        input: {
          request: "Repair formula patterns from template",
          mode: "preview_update",
          intent: { action: "repair_formula_patterns" },
          target: { sheetName: "Report" },
          values: { templateId: "template_unit" }
        }
      },
      {
        action: "convert_formulas_to_values",
        expectedKind: "formula_convert_to_values_preview",
        expectedRuntimeCall: "formula.convert_to_values",
        input: {
          request: "Convert formulas to values",
          mode: "preview_update",
          intent: { action: "convert_formulas_to_values" },
          target: { sheetName: "Report", range: "B2" }
        }
      }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run(testCase.input);
      const applied = await agent.run({
        request: `Apply ${testCase.action}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe(testCase.expectedKind);
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.runtimeMethodCalls[testCase.expectedRuntimeCall]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    }
  });

  it("previews and applies named item mutations through backend orchestration", async () => {
    const cases: Array<{
      action: "create_name" | "update_name" | "delete_name";
      expectedKind: string;
      expectedRuntimeCall: string;
      values: Record<string, unknown>;
      target?: Parameters<AgentOrchestrator["run"]>[0]["target"];
    }> = [
      { action: "create_name", expectedKind: "name_create_preview", expectedRuntimeCall: "names.create", values: { name: "NewInput", comment: "Unit input" }, target: { sheetName: "Report", range: "B2:B3" } },
      { action: "update_name", expectedKind: "name_update_preview", expectedRuntimeCall: "names.update", values: { name: "RevenueTotal", formula: "=Report!$B$2" } },
      { action: "delete_name", expectedKind: "name_delete_preview", expectedRuntimeCall: "names.delete", values: { name: "RevenueTotal" } }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: `Prepare ${testCase.action}`,
        mode: "preview_update",
        intent: { action: testCase.action },
        values: testCase.values,
        ...(testCase.target ? { target: testCase.target } : {})
      });
      const applied = await agent.run({
        request: `Apply ${testCase.action}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe(testCase.expectedKind);
      expect(applied.status).toBe("SUCCESS");
      const runtimeCallCount =
        runtime.runtimeMethodCalls[testCase.expectedRuntimeCall] ??
        runtime.tableMethodCalls.filter((call) => call.method === testCase.expectedRuntimeCall).length;
      expect(runtimeCallCount).toBe(1);
    }
  });

  it("previews and applies region mutations through backend orchestration", async () => {
    const cases: Array<{
      action: "register_region" | "clear_region_values" | "write_region_values" | "fill_region";
      expectedKind: string;
      expectedRuntimeCall: string;
      input: Parameters<AgentOrchestrator["run"]>[0];
    }> = [
      {
        action: "register_region",
        expectedKind: "region_register_preview",
        expectedRuntimeCall: "region.register",
        input: {
          request: "Register input region",
          mode: "preview_update",
          intent: { action: "register_region" },
          target: { sheetName: "Report", range: "B1:B3" },
          values: { regionName: "InputRegion", kind: "data", description: "Editable inputs" }
        }
      },
      {
        action: "clear_region_values",
        expectedKind: "region_clear_values_preview",
        expectedRuntimeCall: "region.clear_values",
        input: {
          request: "Clear region values",
          mode: "preview_update",
          intent: { action: "clear_region_values" },
          values: { regionName: "InputRegion" }
        }
      },
      {
        action: "write_region_values",
        expectedKind: "region_write_values_preview",
        expectedRuntimeCall: "region.write_values",
        input: {
          request: "Write region values",
          mode: "preview_update",
          intent: { action: "write_region_values" },
          values: { regionName: "InputRegion", values: [["Owner"], ["Finance"], ["Ready"]] }
        }
      },
      {
        action: "fill_region",
        expectedKind: "region_fill_preview",
        expectedRuntimeCall: "region.fill",
        input: {
          request: "Fill region values",
          mode: "preview_update",
          intent: { action: "fill_region" },
          values: { regionName: "InputRegion", rows: [["Owner"], ["Finance"], ["Ready"]], clearFirst: true }
        }
      }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run(testCase.input);
      const applied = await agent.run({
        request: `Apply ${testCase.action}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe(testCase.expectedKind);
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.runtimeMethodCalls[testCase.expectedRuntimeCall]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    }
  });

  it("previews and applies template mutations through backend orchestration", async () => {
    const cases: Array<{
      action: "register_template" | "unregister_template" | "clear_template_data_regions" | "fill_template_regions" | "repair_sheet_from_template";
      expectedKind: string;
      expectedRuntimeCall: string;
      expectedWriteOperationKind?: string;
      input: Parameters<AgentOrchestrator["run"]>[0];
    }> = [
      {
        action: "register_template",
        expectedKind: "template_register_preview",
        expectedRuntimeCall: "template.register",
        input: {
          request: "Register template",
          mode: "preview_update",
          intent: { action: "register_template" },
          target: { sheetName: "Report" },
          values: { name: "Report Template", dataRegions: ["B2:B3"] }
        }
      },
      {
        action: "unregister_template",
        expectedKind: "template_unregister_preview",
        expectedRuntimeCall: "template.unregister",
        input: {
          request: "Unregister template",
          mode: "preview_update",
          intent: { action: "unregister_template" },
          values: { templateId: "template_unit" }
        }
      },
      {
        action: "clear_template_data_regions",
        expectedKind: "template_clear_data_regions_preview",
        expectedRuntimeCall: "template.repair_sheet",
        input: {
          request: "Clear template data regions",
          mode: "preview_update",
          intent: { action: "clear_template_data_regions" },
          target: { sheetName: "Report" },
          values: { templateId: "template_unit" }
        }
      },
      {
        action: "fill_template_regions",
        expectedKind: "template_fill_regions_preview",
        expectedRuntimeCall: "batch.apply",
        expectedWriteOperationKind: "range.write_values",
        input: {
          request: "Fill template data regions",
          mode: "preview_update",
          intent: { action: "fill_template_regions" },
          target: { sheetName: "Report" },
          values: { templateId: "template_unit", regionValues: { "B2:B3": [["North"], ["South"]] } }
        }
      },
      {
        action: "repair_sheet_from_template",
        expectedKind: "template_repair_preview",
        expectedRuntimeCall: "template.repair_sheet",
        input: {
          request: "Repair sheet from template",
          mode: "preview_update",
          intent: { action: "repair_sheet_from_template" },
          target: { sheetName: "Report" },
          values: { templateId: "template_unit", repair: ["styles", "formulas"] }
        }
      }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const preview = await agent.run(testCase.input);
      const applied = await agent.run({
        request: `Apply ${testCase.action}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe(testCase.expectedKind);
      expect(applied.status).toBe("SUCCESS");
      if (testCase.expectedWriteOperationKind) {
        expect(runtime.writeBatchCount).toBe(1);
        expect(runtime.lastWriteOperations[0]?.kind).toBe(testCase.expectedWriteOperationKind);
      } else {
        const runtimeCallCount =
          runtime.runtimeMethodCalls[testCase.expectedRuntimeCall] ??
          runtime.tableMethodCalls.filter((call) => call.method === testCase.expectedRuntimeCall).length;
        expect(runtimeCallCount).toBe(1);
      }
    }
  });

  it("previews and applies style template mutations through backend orchestration", async () => {
    const cases: Array<{
      action: "copy_style_from_template" | "repair_style_consistency";
      expectedKind: string;
      expectedRuntimeCall: string;
      input: Parameters<AgentOrchestrator["run"]>[0];
    }> = [
      {
        action: "copy_style_from_template",
        expectedKind: "style_copy_preview",
        expectedRuntimeCall: "style.copy_dimensions",
        input: {
          request: "Copy style from template",
          mode: "preview_update",
          intent: { action: "copy_style_from_template" },
          values: {
            source: { sheetName: "Data", range: "A1:B4" },
            destination: { sheetName: "Report", range: "A1:B4" },
            dimensions: ["fills", "fonts"]
          }
        }
      },
      {
        action: "repair_style_consistency",
        expectedKind: "style_repair_consistency_preview",
        expectedRuntimeCall: "style.repair_consistency",
        input: {
          request: "Repair style consistency",
          mode: "preview_update",
          intent: { action: "repair_style_consistency" },
          target: { sheetName: "Report" },
          values: { templateId: "template_unit" }
        }
      }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const preview = await agent.run(testCase.input);
      const applied = await agent.run({
        request: `Apply ${testCase.action}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe(testCase.expectedKind);
      expect(applied.status).toBe("SUCCESS");
      const runtimeCallCount =
        runtime.runtimeMethodCalls[testCase.expectedRuntimeCall] ??
        runtime.tableMethodCalls.filter((call) => call.method === testCase.expectedRuntimeCall).length;
      expect(runtimeCallCount).toBe(1);
    }
  });

  it("previews grouped style copies as one apply operation", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Copy same style from source to targets",
      mode: "preview_update",
      intent: { action: "copy_style_from_template" },
      values: {
        styleCopies: [
          { source: { sheetName: "Data", range: "A1:B1" }, destination: { sheetName: "Report", range: "A1:B1" }, dimensions: ["fills", "fonts"] },
          { source: { sheetName: "Data", range: "A2:B2" }, destination: { sheetName: "Report", range: "C1:D1" }, dimensions: ["fills", "fonts"] }
        ]
      }
    });
    const applied = await agent.run({
      request: "Apply grouped style copies",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("style_copy_many_preview");
    expect(applied.status).toBe("SUCCESS");
    expect((applied.answer as any).backupIds).toEqual(["backup_style_many"]);
    expect((applied.answer as any).rollbackAvailable).toBe(true);
    expect(runtime.runtimeMethodCalls["style.copy_dimensions_many"]).toBe(1);
    expect(runtime.runtimeMethodCalls["style.copy_dimensions"]).toBeUndefined();
  });

  it("previews and applies a styled table replacement as one workflow", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const headers = ["Dated", "Loading Date", "Qty", "Attention", "From (FM)", "Description", "Received (Tons)", "Volume"];
    const row = ["20/6/26", "20/6/26", "5X40'HQ", "5X40'HQ", "SC88", "Loading at Sai Noi Factory", 32.5, "5X40'HQ UPGRADE"];
    const preview = await agent.run({
      request: "Rotate booking fields to headers and replace the old vertical table with same style",
      mode: "preview_update",
      intent: { action: "replace_range_with_styled_table" },
      target: { sheetName: "Report", range: "A1:H2" },
      values: {
        headers,
        row,
        clearRange: "A1:B25",
        headerStyleSource: { sheetName: "Data", range: "A1:D1" },
        bodyStyleSource: { sheetName: "Data", range: "A2:D2" },
        dimensions: ["fills", "fonts", "borders", "alignment"],
        autofit: true
      }
    });
    const applied = await agent.run({
      request: "Apply styled table replacement",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("replace_range_with_styled_table_preview");
    expect((preview.answer as any).clearCount).toBe(1);
    expect((preview.answer as any).styleCopyCount).toBe(4);
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.writeBatchCount).toBe(1);
    expect(runtime.lastWriteOperations.map((operation) => operation.kind)).toEqual(["range.clear", "range.write_values", "range.autofit_columns"]);
    expect((runtime.lastWriteOperations[0] as any).applyTo).toBe("all");
    expect(runtime.runtimeMethodCalls["style.copy_dimensions_many"]).toBe(1);
    expect(runtime.runtimeMethodCalls["style.copy_dimensions"]).toBeUndefined();
  });

  it("routes generic OCR field/value data to the styled table replacement workflow", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Convert extracted invoice field/value OCR data into a horizontal styled table",
      mode: "preview_update",
      target: { sheetName: "Report", range: "A1:D2" },
      values: {
        headers: ["Invoice", "Date", "Customer", "Total"],
        row: ["INV-100", "2026-06-20", "ACME", 2500],
        clearRange: "A1:D20",
        headerStyleSource: { sheetName: "Data", range: "A1:D1" },
        bodyStyleSource: { sheetName: "Data", range: "A2:D2" }
      }
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("replace_range_with_styled_table_preview");
    expect((preview.answer as any).styleCopyCount).toBe(2);
  });

  it("redirects fragmented style-copy previews to a grouped workflow", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);
    const metadata = createCachedMetadata("wbctx_fragmented_style");
    agent.metadataCache.set(metadata);

    const base = {
      mode: "preview_update" as const,
      workbookContextId: metadata.workbookContextId,
      values: {
        source: { sheetName: "Data", range: "A2:D2" }
      }
    };
    await agent.run({ ...base, request: "Copy body style to Booking A2:D2", values: { ...base.values, destination: { sheetName: "Report", range: "A2:D2" } } });
    const redirected = await agent.run({ ...base, request: "Copy body style to Booking E2:H2", values: { ...base.values, destination: { sheetName: "Report", range: "E2:H2" } } });

    expect(redirected.status).toBe("NEEDS_WORKFLOW_REDIRECT");
    expect(redirected.nextAction).toBe("call_preview_update");
    expect((redirected.answer as any).kind).toBe("workflow_redirect");
    expect((redirected.answer as any).suggestedIntentAction).toBe("copy_style_from_template");
    expect((redirected.answer as any).suggestedValues.styleCopies).toHaveLength(2);
    expect(redirected.warnings).toContain("Fragmented style-copy previews were redirected to a grouped workflow.");
  });

  it("redirects repeated value-only cleanup because stale formats can remain", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);
    const metadata = createCachedMetadata("wbctx_fragmented_clear");
    agent.metadataCache.set(metadata);

    await agent.run({
      request: "Clear leftover data in A3:X10",
      mode: "preview_update",
      workbookContextId: metadata.workbookContextId,
      intent: { action: "clear_values" },
      target: { sheetName: "Report", range: "A3:X10" }
    });
    const redirected = await agent.run({
      request: "Clear leftover data in A11:X18",
      mode: "preview_update",
      workbookContextId: metadata.workbookContextId,
      intent: { action: "clear_values" },
      target: { sheetName: "Report", range: "A11:X18" }
    });

    expect(redirected.status).toBe("NEEDS_WORKFLOW_REDIRECT");
    expect((redirected.answer as any).suggestedIntentAction).toBe("clear_range");
    expect((redirected.answer as any).suggestedValues.clearRange).toBe("A3:X18");
    expect(redirected.warnings).toContain("Fragmented value-only cleanup was redirected because it can leave stale formatting.");
  });

  it("redirects adjacent value-write previews to grouped patches", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);
    const metadata = createCachedMetadata("wbctx_fragmented_write_values");
    agent.metadataCache.set(metadata);

    await agent.run({
      request: "Write first extracted value",
      mode: "preview_update",
      workbookContextId: metadata.workbookContextId,
      target: { sheetName: "Report", range: "A1" },
      values: { values: [["A"]] }
    });
    const redirected = await agent.run({
      request: "Write second extracted value",
      mode: "preview_update",
      workbookContextId: metadata.workbookContextId,
      target: { sheetName: "Report", range: "B1" },
      values: { values: [["B"]] }
    });

    expect(redirected.status).toBe("NEEDS_WORKFLOW_REDIRECT");
    expect((redirected.answer as any).suggestedIntentAction).toBe("write_values");
    expect((redirected.answer as any).suggestedValues.patches).toEqual([
      { target: { sheetName: "Report", range: "A1" }, values: [["A"]] },
      { target: { sheetName: "Report", range: "B1" }, values: [["B"]] }
    ]);
    expect(redirected.warnings).toContain("Fragmented value-write previews were redirected to a grouped patch workflow.");
  });

  it("redirects repeated format previews to write_styles_many", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);
    const metadata = createCachedMetadata("wbctx_fragmented_format");
    agent.metadataCache.set(metadata);

    await agent.run({
      request: "Make A1 bold",
      mode: "preview_update",
      workbookContextId: metadata.workbookContextId,
      intent: { action: "format_range" },
      target: { sheetName: "Report", range: "A1" }
    });
    const redirected = await agent.run({
      request: "Make B1 bold",
      mode: "preview_update",
      workbookContextId: metadata.workbookContextId,
      intent: { action: "format_range" },
      target: { sheetName: "Report", range: "B1" }
    });

    expect(redirected.status).toBe("NEEDS_WORKFLOW_REDIRECT");
    expect((redirected.answer as any).suggestedIntentAction).toBe("write_styles_many");
    expect((redirected.answer as any).suggestedValues.entries).toHaveLength(2);
    expect(redirected.warnings).toContain("Fragmented format previews were redirected to a grouped style workflow.");
  });

  describe("OpenCode production regressions", () => {
    it("applies explicit black header styling without falling back to default header colors", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Set header row A1:X1 on Data sheet to black fill (#000000) with white font color (#FFFFFF), bold text, center aligned.",
        mode: "preview_update",
        intent: { action: "format_range" },
        target: { sheetName: "Data", range: "A1:X1" }
      });
      const applied = await agent.run({
        request: "Apply black header style.",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe("style_preview");
      expect((preview.answer as any).style).toMatchObject({
        fillColor: "#000000",
        fontColor: "#FFFFFF",
        fontBold: true,
        horizontalAlignment: "center"
      });
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations[0]).toMatchObject({
        kind: "range.write_styles",
        target: { sheetName: "Data", address: "A1:X1" },
        style: {
          fillColor: "#000000",
          fontColor: "#FFFFFF",
          fontBold: true,
          horizontalAlignment: "center"
        },
        preserveValues: true
      });
    });

    it("applies flattened multi-style entries as one write_styles_many batch", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Set yellow fill #FFFF00 on rows 2, 4, and 7 of Booking sheet for full row A:X.",
        mode: "preview_update",
        intent: { action: "write_styles_many" },
        target: { sheetName: "Booking" },
        values: {
          entries: [
            { range: "A2:X2", sheetName: "Booking", fillColor: "#FFFF00" },
            { range: "A4:X4", sheetName: "Booking", fillColor: "#FFFF00" },
            { range: "A7:X7", sheetName: "Booking", fillColor: "#FFFF00" }
          ]
        }
      });
      const applied = await agent.run({
        request: "Apply yellow row highlights.",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.writeBatchCount).toBe(1);
      expect(runtime.lastBatchOperations).toEqual([
        expect.objectContaining({
          kind: "range.write_styles_many",
          entries: [
            { target: expect.objectContaining({ sheetName: "Booking", address: "A2:X2" }), style: { fillColor: "#FFFF00" }, preserveValues: true },
            { target: expect.objectContaining({ sheetName: "Booking", address: "A4:X4" }), style: { fillColor: "#FFFF00" }, preserveValues: true },
            { target: expect.objectContaining({ sheetName: "Booking", address: "A7:X7" }), style: { fillColor: "#FFFF00" }, preserveValues: true }
          ]
        })
      ]);
    });

    it("applies inserted columns through the structural batch operation", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Add new columns next to Qty on Data.",
        mode: "preview_update",
        intent: { action: "insert_columns" },
        target: { sheetName: "Data", range: "C:D" }
      });
      const applied = await agent.run({
        request: "Apply inserted columns.",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe("range.insert_columns_preview");
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations[0]).toMatchObject({
        kind: "range.insert_columns",
        destructiveLevel: "structure",
        target: { sheetName: "Data", address: "C1:D4" }
      });
    });

    it("applies range column swaps with a normalized columnOrder payload", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Swap columns A and B in Data",
        mode: "preview_update",
        intent: { action: "reorder_range_columns" },
        target: { sheetName: "Data", range: "A1:B4" },
        values: { order: [2, 1] }
      });
      const applied = await agent.run({
        request: "Apply column swap.",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).columnOrder).toEqual([2, 1]);
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations[0]).toMatchObject({
        kind: "range.reorder_columns",
        destructiveLevel: "structure",
        target: { sheetName: "Data", address: "A1:B4" },
        columnOrder: [2, 1]
      });
    });

    it("applies table column swaps through the table host method", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Move Status before Date in the Transactions table.",
        mode: "preview_update",
        intent: { action: "reorder_table_columns" },
        target: { tableName: "Transactions" },
        values: { columns: ["Status", "Date", "Account", "Amount"] }
      });
      const applied = await agent.run({
        request: "Apply table column order.",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe("table_reorder_columns_preview");
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.tableMethodCalls).toContainEqual({
        method: "table.reorder_columns",
        request: expect.objectContaining({
          tableName: "Transactions",
          columnOrder: ["Status", "Date", "Account", "Amount"]
        })
      });
    });

    it("applies dropdown validation from agent-friendly options", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Add select list to Data D2:D4.",
        mode: "preview_update",
        intent: { action: "write_data_validation" },
        target: { sheetName: "Data", range: "D2:D4" },
        values: { options: ["20GP", "40GP", "40HQ"], inCellDropDown: true }
      });
      const applied = await agent.run({
        request: "Apply dropdown validation.",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).validation).toEqual({
        type: "list",
        source: ["20GP", "40GP", "40HQ"],
        inCellDropDown: true,
        ignoreBlanks: true
      });
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations[0]).toMatchObject({
        kind: "range.write_data_validation",
        target: { sheetName: "Data", address: "D2:D4" },
        validation: {
          type: "list",
          source: ["20GP", "40GP", "40HQ"],
          inCellDropDown: true,
          ignoreBlanks: true
        }
      });
    });

    it("applies formula conditional formatting from formula and style fields", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Add formula color on Data A2:D4 when column D is 40HQ.",
        mode: "preview_update",
        intent: { action: "write_conditional_formatting" },
        target: { sheetName: "Data", range: "A2:D4" },
        values: { formula: "=$D2=\"40HQ\"", style: { fillColor: "#FFFF00", fontColor: "#000000" } }
      });
      const applied = await agent.run({
        request: "Apply formula conditional formatting.",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe("write_conditional_formatting_preview");
      expect((preview.answer as any).kind).not.toBe("sheet_create_preview");
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations[0]).toMatchObject({
        kind: "range.write_conditional_formatting",
        target: { sheetName: "Data", address: "A2:D4" },
        rule: {
          type: "custom",
          formula: "=$D2=\"40HQ\"",
          style: { fillColor: "#FFFF00", fontColor: "#000000" }
        }
      });
    });
  });

  it("honors explicit black header colors instead of default header styling", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Set header row A1:X1 on Data sheet to black fill (#000000) with white font color (#FFFFFF), bold text, center aligned.",
      mode: "preview_update",
      intent: { action: "format_range" },
      target: { sheetName: "Data", range: "A1:X1" }
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).style).toMatchObject({
      fillColor: "#000000",
      fontColor: "#FFFFFF",
      fontBold: true
    });
  });

  it("honors explicit header fill hex when the request also names a color family", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Set HR A1:E1 to dark blue fill (#1F4E78), white font (#FFFFFF), bold text, center aligned.",
      mode: "preview_update",
      intent: { action: "format_range" },
      target: { sheetName: "Data", range: "A1:E1" }
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).style).toMatchObject({
      fillColor: "#1F4E78",
      fontColor: "#FFFFFF",
      fontBold: true,
      horizontalAlignment: "center"
    });
  });

  it("accepts flattened write_styles_many entries from agents", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Set yellow fill #FFFF00 on rows 2, 4, and 7 of Booking sheet for full row A:X.",
      mode: "preview_update",
      intent: { action: "write_styles_many" },
      target: { sheetName: "Booking" },
      values: {
        entries: [
          { range: "A2:X2", sheetName: "Booking", fillColor: "#FFFF00" },
          { range: "A4:X4", sheetName: "Booking", fillColor: "#FFFF00" },
          { range: "A7:X7", sheetName: "Booking", fillColor: "#FFFF00" }
        ]
      }
    });
    const applied = await agent.run({
      request: "Apply yellow highlights.",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastBatchOperations[0]).toMatchObject({
      kind: "range.write_styles_many",
      entries: [
        { target: { sheetName: "Booking", address: "A2:X2" }, style: { fillColor: "#FFFF00" } },
        { target: { sheetName: "Booking", address: "A4:X4" }, style: { fillColor: "#FFFF00" } },
        { target: { sheetName: "Booking", address: "A7:X7" }, style: { fillColor: "#FFFF00" } }
      ]
    });
  });

  it("previews dropdown data validation as a validation operation", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Set data validation on Data D2:D4 as a list dropdown. Allowed values: 20GP, 40GP, 40HQ.",
      mode: "preview_update",
      intent: { action: "write_data_validation" },
      target: { sheetName: "Data", range: "D2:D4" },
      values: { validation: { type: "List", formula1: "20GP,40GP,40HQ" } }
    });
    const applied = await agent.run({
      request: "Apply dropdown validation.",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("write_data_validation_preview");
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastBatchOperations[0]).toMatchObject({
      kind: "range.write_data_validation",
      target: { sheetName: "Data", address: "D2:D4" },
      validation: { type: "list", source: ["20GP", "40GP", "40HQ"] }
    });
  });

  it("previews formula conditional formatting instead of sheet creation", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Add conditional formatting rule on Data sheet range A2:D4. Formula =$D2=\"40HQ\" should fill the entire row with yellow color #FFFF00.",
      mode: "preview_update",
      target: { sheetName: "Data", range: "A2:D4" },
      values: { formula: "=$D2=\"40HQ\"", style: { fillColor: "#FFFF00" } }
    });
    const applied = await agent.run({
      request: "Apply conditional formatting.",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("write_conditional_formatting_preview");
    expect((preview.answer as any).kind).not.toBe("sheet_create_preview");
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastBatchOperations[0]).toMatchObject({
      kind: "range.write_conditional_formatting",
      target: { sheetName: "Data", address: "A2:D4" },
      rule: { type: "custom", formula: "=$D2=\"40HQ\"", style: { fillColor: "#FFFF00" } }
    });
  });

  it("previews plain range column reorder for swap requests", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Swap columns A and B in Data",
      mode: "preview_update",
      intent: { action: "reorder_range_columns" },
      target: { sheetName: "Data", range: "A1:B4" },
      values: { columnOrder: [2, 1] }
    });
    const applied = await agent.run({
      request: "Apply column swap.",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("range.reorder_columns_preview");
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastBatchOperations[0]).toMatchObject({
      kind: "range.reorder_columns",
      target: { sheetName: "Data", address: "A1:B4" },
      columnOrder: [2, 1]
    });
  });

  it("redirects repeated formula fills to one larger formula workflow", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);
    const metadata = createCachedMetadata("wbctx_fragmented_formula_fill");
    agent.metadataCache.set(metadata);

    const base = {
      request: "Fill formula down",
      mode: "preview_update" as const,
      workbookContextId: metadata.workbookContextId,
      intent: { action: "fill_formula_down" as const }
    };
    await agent.run({
      ...base,
      values: {
        source: { sheetName: "Report", range: "A1" },
        destination: { sheetName: "Report", range: "A2:A5" }
      }
    });
    const redirected = await agent.run({
      ...base,
      values: {
        source: { sheetName: "Report", range: "A1" },
        destination: { sheetName: "Report", range: "A6:A10" }
      }
    });

    expect(redirected.status).toBe("NEEDS_WORKFLOW_REDIRECT");
    expect((redirected.answer as any).suggestedIntentAction).toBe("fill_formula_down");
    expect((redirected.answer as any).suggestedValues.destination).toEqual({ sheetName: "Report", range: "A2:A10" });
    expect(redirected.warnings).toContain("Fragmented formula-fill-down previews were redirected to a grouped formula workflow.");
  });

  it("reports partial workflow failure with rollback evidence", async () => {
    const runtime = new FakeAgentRuntime();
    runtime.failStyleCopyOnCall = 2;
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Replace table and copy style",
      mode: "preview_update",
      intent: { action: "replace_range_with_styled_table" },
      target: { sheetName: "Report", range: "A1:B2" },
      values: {
        headers: ["Dated", "Qty"],
        row: ["20/6/26", "5X40'HQ"],
        clearRange: "A1:B20",
        headerStyleSource: { sheetName: "Data", range: "A1:B1" },
        bodyStyleSource: { sheetName: "Data", range: "A2:B2" },
        dimensions: ["fills", "fonts"]
      }
    });
    const applied = await agent.run({
      request: "Apply styled table replacement",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect(applied.status).toBe("VALIDATION_FAILED");
    expect((applied.answer as any).partialFailure).toBe(true);
    expect((applied.answer as any).rollbackAvailable).toBe(true);
    expect((applied.answer as any).backupIds).toContain("backup_style_many");
    expect((applied.answer as any).stepResults).toHaveLength(2);
    expect(applied.warnings).toContain("Style copy failed");
  });

  it("rejects mismatched direct style copy ranges before apply", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Copy same style from Data to Report",
      mode: "preview_update",
      intent: { action: "copy_style_from_template" },
      values: {
        source: { sheetName: "Data", range: "A1:D1" },
        destination: { sheetName: "Report", range: "A1:H1" },
        dimensions: ["fills", "fonts"]
      }
    });

    expect(preview.status).toBe("NEEDS_INPUT");
    expect(preview.summary).toContain("same dimensions");
    expect(runtime.runtimeMethodCalls["style.copy_dimensions"]).toBeUndefined();
  });

  it("returns cached apply results when the same operation is retried", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Write values",
      mode: "preview_update",
      intent: { action: "write_values" },
      target: { sheetName: "Report", range: "A1:B1" },
      values: { values: [["A", "B"]] }
    });
    const first = await agent.run({
      request: "Apply values",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });
    const second = await agent.run({
      request: "Apply values again",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(first.status).toBe("SUCCESS");
    expect(second.status).toBe("SUCCESS");
    expect(runtime.writeBatchCount).toBe(1);
  });

  it("previews and applies repair capabilities through backend orchestration", async () => {
    const cases: Array<{
      action: "repair_style_from_template" | "repair_formulas_from_template" | "repair_table_structure";
      expectedKind: string;
      expectedRuntimeCall: string;
      input: Parameters<AgentOrchestrator["run"]>[0];
    }> = [
      {
        action: "repair_style_from_template",
        expectedKind: "style_repair_consistency_preview",
        expectedRuntimeCall: "style.repair_consistency",
        input: {
          request: "Repair styles from template",
          mode: "preview_update",
          intent: { action: "repair_style_from_template" },
          target: { sheetName: "Report" },
          values: { templateId: "template_unit" }
        }
      },
      {
        action: "repair_formulas_from_template",
        expectedKind: "formula_repair_patterns_preview",
        expectedRuntimeCall: "formula.repair_patterns",
        input: {
          request: "Repair formulas from template",
          mode: "preview_update",
          intent: { action: "repair_formulas_from_template" },
          target: { sheetName: "Report" },
          values: { templateId: "template_unit" }
        }
      },
      {
        action: "repair_table_structure",
        expectedKind: "table_copy_structure_preview",
        expectedRuntimeCall: "table.copy_structure",
        input: {
          request: "Repair table structure",
          mode: "preview_update",
          intent: { action: "repair_table_structure" },
          target: { tableName: "Transactions" },
          values: { targetSheetName: "Report", targetAddress: "A1:D4" }
        }
      }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const preview = await agent.run(testCase.input);
      const applied = await agent.run({
        request: `Apply ${testCase.action}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe(testCase.expectedKind);
      expect(applied.status).toBe("SUCCESS");
      const runtimeCallCount =
        runtime.runtimeMethodCalls[testCase.expectedRuntimeCall] ??
        runtime.tableMethodCalls.filter((call) => call.method === testCase.expectedRuntimeCall).length;
      expect(runtimeCallCount).toBe(1);
    }
  });

  it("previews and applies cleaning mutations through backend orchestration", async () => {
    const cases: Array<{
      action: "normalize_headers" | "trim_whitespace" | "remove_duplicates" | "parse_dates" | "parse_numbers" | "standardize_currency" | "fill_missing_values" | "split_column" | "merge_columns";
      expectedRuntimeCall: string;
      values?: Record<string, unknown>;
    }> = [
      { action: "normalize_headers", expectedRuntimeCall: "clean.normalize_headers", values: { headerRowIndex: 0 } },
      { action: "trim_whitespace", expectedRuntimeCall: "clean.trim_whitespace" },
      { action: "remove_duplicates", expectedRuntimeCall: "clean.remove_duplicates", values: { hasHeader: true, keyColumns: [0, 1] } },
      { action: "parse_dates", expectedRuntimeCall: "clean.parse_dates" },
      { action: "parse_numbers", expectedRuntimeCall: "clean.parse_numbers" },
      { action: "standardize_currency", expectedRuntimeCall: "clean.standardize_currency" },
      { action: "fill_missing_values", expectedRuntimeCall: "clean.fill_missing_values", values: { strategy: "zero" } },
      { action: "split_column", expectedRuntimeCall: "clean.split_column", values: { columnIndex: 1, delimiter: "-", targetAddress: "E1:F4" } },
      { action: "merge_columns", expectedRuntimeCall: "clean.merge_columns", values: { columnIndexes: [1, 2], separator: " ", targetAddress: "E1:E4" } }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const preview = await agent.run({
        request: `Clean data with ${testCase.action}`,
        mode: "preview_update",
        intent: { action: testCase.action },
        target: { sheetName: "Data", range: "A1:D4" },
        values: testCase.values
      });
      const applied = await agent.run({
        request: `Apply ${testCase.action}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe("cleaning_preview");
      expect((preview.answer as any).action).toBe(testCase.action);
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.runtimeMethodCalls[testCase.expectedRuntimeCall]).toBe(1);
    }
  });

  it("previews formula recalculation as a workbook calculate operation", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Recalculate formulas",
      mode: "preview_update",
      intent: { action: "recalculate_formulas" }
    });
    const applied = await agent.run({
      request: "Apply formula recalc",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("formula.recalculate_preview");
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastBatchOperations[0]?.kind).toBe("workbook.calculate");
  });

  it("routes every covered range-core capability to its internal operation", async () => {
    const cases: Array<{
      capabilityName: string;
      expectedOperationKind: string;
      input: Parameters<AgentOrchestrator["run"]>[0];
    }> = [
      {
        capabilityName: "excel.range.write_values",
        expectedOperationKind: "range.write_values",
        input: {
          request: "Update Data B2",
          mode: "preview_update",
          target: { sheetName: "Data", range: "B2" },
          values: { values: [[123]] }
        }
      },
      {
        capabilityName: "excel.range.write_formulas",
        expectedOperationKind: "range.write_formulas",
        input: {
          request: "Write formula to Report A12",
          mode: "preview_update",
          intent: { action: "write_formulas" },
          target: { sheetName: "Report", range: "A12" },
          values: { values: [["=SUM(B1:B10)"]] }
        }
      },
      {
        capabilityName: "excel.range.write_number_formats",
        expectedOperationKind: "range.write_number_formats",
        input: {
          request: "Format Data C2 as currency",
          mode: "preview_update",
          intent: { action: "write_number_formats" },
          target: { sheetName: "Data", range: "C2" },
          values: { numberFormat: "$#,##0.00" }
        }
      },
      {
        capabilityName: "excel.range.write_styles",
        expectedOperationKind: "range.write_styles",
        input: {
          request: "Format the header row on Data",
          mode: "preview_update",
          intent: { action: "format_range" },
          target: { sheetName: "Data", range: "A1:D1" }
        }
      },
      {
        capabilityName: "excel.range.clear",
        expectedOperationKind: "range.clear",
        input: {
          request: "Clear the whole target",
          mode: "preview_update",
          intent: { action: "clear_range" },
          target: { sheetName: "Data", range: "D4" }
        }
      },
      {
        capabilityName: "excel.range.clear_values_keep_format",
        expectedOperationKind: "range.clear_values_keep_format",
        input: {
          request: "Clear values",
          mode: "preview_update",
          intent: { action: "clear_values" },
          target: { sheetName: "Data", range: "B3" }
        }
      },
      {
        capabilityName: "excel.range.clear_values",
        expectedOperationKind: "range.clear_values",
        input: {
          request: "Clear raw values",
          mode: "preview_update",
          intent: { action: "clear_values_raw" },
          target: { sheetName: "Data", range: "B3" }
        }
      },
      {
        capabilityName: "excel.range.clear_formats",
        expectedOperationKind: "range.clear_formats",
        input: {
          request: "Clear formatting",
          mode: "preview_update",
          intent: { action: "clear_formats" },
          target: { sheetName: "Data", range: "A1:D1" }
        }
      },
      {
        capabilityName: "excel.range.copy",
        expectedOperationKind: "range.copy",
        input: {
          request: "Copy range",
          mode: "preview_update",
          intent: { action: "copy_range" },
          values: {
            source: { sheetName: "Data", range: "A1:B2" },
            destination: { sheetName: "Report", range: "A1:B2" }
          }
        }
      },
      {
        capabilityName: "excel.range.write_styles_many",
        expectedOperationKind: "range.write_styles_many",
        input: {
          request: "Style multiple ranges",
          mode: "preview_update",
          intent: { action: "write_styles_many" },
          values: {
            entries: [
              { sheetName: "Data", range: "A1:D1", style: { bold: true } },
              { sheetName: "Report", range: "B2:C2", style: { fillColor: "#fff2cc" } }
            ]
          }
        }
      },
      {
        capabilityName: "excel.range.move",
        expectedOperationKind: "range.move",
        input: {
          request: "Move range",
          mode: "preview_update",
          intent: { action: "move_range" },
          values: {
            source: { sheetName: "Data", range: "A1:B2" },
            destination: { sheetName: "Report", range: "C1:D2" }
          }
        }
      },
      {
        capabilityName: "excel.range.autofit_columns",
        expectedOperationKind: "range.autofit_columns",
        input: {
          request: "Autofit columns",
          mode: "preview_update",
          intent: { action: "autofit" },
          target: { sheetName: "Data", range: "A:D" }
        }
      },
      {
        capabilityName: "excel.range.autofit_rows",
        expectedOperationKind: "range.autofit_rows",
        input: {
          request: "Autofit rows",
          mode: "preview_update",
          intent: { action: "autofit_rows" },
          target: { sheetName: "Data", range: "A1:D4" }
        }
      },
      {
        capabilityName: "excel.range.insert_rows",
        expectedOperationKind: "range.insert_rows",
        input: {
          request: "Insert rows",
          mode: "preview_update",
          intent: { action: "insert_rows" },
          target: { sheetName: "Data", range: "2:3" }
        }
      },
      {
        capabilityName: "excel.range.delete_rows",
        expectedOperationKind: "range.delete_rows",
        input: {
          request: "Delete rows",
          mode: "preview_update",
          intent: { action: "delete_rows" },
          target: { sheetName: "Data", range: "2:3" }
        }
      },
      {
        capabilityName: "excel.range.insert_columns",
        expectedOperationKind: "range.insert_columns",
        input: {
          request: "Insert columns",
          mode: "preview_update",
          intent: { action: "insert_columns" },
          target: { sheetName: "Data", range: "B:C" }
        }
      },
      {
        capabilityName: "excel.range.delete_columns",
        expectedOperationKind: "range.delete_columns",
        input: {
          request: "Delete columns",
          mode: "preview_update",
          intent: { action: "delete_columns" },
          target: { sheetName: "Data", range: "B:C" }
        }
      },
      {
        capabilityName: "excel.range.merge",
        expectedOperationKind: "range.merge",
        input: {
          request: "Merge range",
          mode: "preview_update",
          intent: { action: "merge_range" },
          target: { sheetName: "Report", range: "A1:B1" },
          values: { across: true }
        }
      },
      {
        capabilityName: "excel.range.unmerge",
        expectedOperationKind: "range.unmerge",
        input: {
          request: "Unmerge range",
          mode: "preview_update",
          intent: { action: "unmerge_range" },
          target: { sheetName: "Report", range: "A1:B1" }
        }
      }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const preview = await agent.run(testCase.input);

      expect(preview.status, testCase.capabilityName).toBe("PREVIEW_READY");

      const applied = await agent.run({
        request: `Apply ${testCase.capabilityName}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(applied.status, testCase.capabilityName).toBe("SUCCESS");
      expect(runtime.lastBatchOperations[0]?.kind, testCase.capabilityName).toBe(testCase.expectedOperationKind);
    }
  });

  it("routes every covered table-core mutation to its runtime method", async () => {
    const cases: Array<{
      capabilityName: string;
      expectedMethod: string;
      input: Parameters<AgentOrchestrator["run"]>[0];
    }> = [
      {
        capabilityName: "excel.table.append_rows",
        expectedMethod: "table.append_rows",
        input: {
          request: "Append transaction rows to the Transactions table",
          mode: "preview_update",
          intent: { action: "append_table_rows" },
          target: { tableName: "Transactions" },
          values: { rows: [["2026-06-01", "A-100", 123, "Open"]] }
        }
      },
      {
        capabilityName: "excel.table.update_rows",
        expectedMethod: "table.update_rows",
        input: {
          request: "Update transaction table row",
          mode: "preview_update",
          intent: { action: "update_table_rows" },
          target: { tableName: "Transactions" },
          values: { rows: [{ index: 0, values: ["2026-06-01", "A-100", 999, "Closed"] }] }
        }
      },
      {
        capabilityName: "excel.table.create",
        expectedMethod: "table.create",
        input: {
          request: "Create a table on Report B1:D4",
          mode: "preview_update",
          intent: { action: "create_table" },
          target: { sheetName: "Report", range: "B1:D4" },
          values: { tableName: "ReportInputs", hasHeaders: true, style: "TableStyleMedium2" }
        }
      },
      {
        capabilityName: "excel.table.resize",
        expectedMethod: "table.resize",
        input: {
          request: "Resize Transactions table",
          mode: "preview_update",
          intent: { action: "resize_table" },
          target: { tableName: "Transactions" },
          values: { address: "A1:D10" }
        }
      },
      {
        capabilityName: "excel.table.reorder_columns",
        expectedMethod: "table.reorder_columns",
        input: {
          request: "Reorder Transactions columns",
          mode: "preview_update",
          intent: { action: "reorder_table_columns" },
          target: { tableName: "Transactions" },
          values: { columnOrder: ["Status", "Date", "Account", "Amount"] }
        }
      },
      {
        capabilityName: "excel.table.clear_data_keep_formulas",
        expectedMethod: "table.clear_data_keep_formulas",
        input: {
          request: "Clear table data rows",
          mode: "preview_update",
          intent: { action: "clear_table_data" },
          target: { tableName: "Transactions" }
        }
      },
      {
        capabilityName: "excel.table.clear_filters",
        expectedMethod: "table.clear_filters",
        input: {
          request: "Clear table filters",
          mode: "preview_update",
          intent: { action: "clear_table_filters" },
          target: { tableName: "Transactions" }
        }
      },
      {
        capabilityName: "excel.table.apply_filters",
        expectedMethod: "table.apply_filters",
        input: {
          request: "Filter Transactions table",
          mode: "preview_update",
          intent: { action: "filter_range" },
          target: { tableName: "Transactions" },
          values: { filters: [{ column: "Status", criteria: { filterOn: "Values", values: ["Open"] } }] }
        }
      },
      {
        capabilityName: "excel.table.sort",
        expectedMethod: "table.sort",
        input: {
          request: "Sort Transactions by amount",
          mode: "preview_update",
          intent: { action: "sort_table" },
          target: { tableName: "Transactions" }
        }
      },
      {
        capabilityName: "excel.table.apply_view",
        expectedMethod: "table.apply_view",
        input: {
          request: "Filter Transactions to open and sort by amount descending",
          mode: "preview_update",
          intent: { action: "apply_table_view" },
          target: { tableName: "Transactions" },
          values: {
            filters: [{ column: "Status", criteria: { filterOn: "Values", values: ["Open"] } }],
            sort: { fields: [{ key: 2, ascending: false }] }
          }
        }
      },
      {
        capabilityName: "excel.table.set_total_row",
        expectedMethod: "table.set_total_row",
        input: {
          request: "Show total row on Transactions table",
          mode: "preview_update",
          intent: { action: "set_table_total_row" },
          target: { tableName: "Transactions" },
          values: { showTotals: true }
        }
      },
      {
        capabilityName: "excel.table.set_style",
        expectedMethod: "table.set_style",
        input: {
          request: "Set table style",
          mode: "preview_update",
          intent: { action: "set_table_style" },
          target: { tableName: "Transactions" },
          values: { style: "TableStyleMedium2" }
        }
      },
      {
        capabilityName: "excel.table.copy_structure",
        expectedMethod: "table.copy_structure",
        input: {
          request: "Copy Transactions table structure",
          mode: "preview_update",
          intent: { action: "copy_table_structure" },
          target: { tableName: "Transactions" },
          values: { targetSheetName: "Report", targetAddress: "A10:D12", newTableName: "TransactionsCopy" }
        }
      }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const preview = await agent.run(testCase.input);

      expect(preview.status, testCase.capabilityName).toBe("PREVIEW_READY");

      const applied = await agent.run({
        request: `Apply ${testCase.capabilityName}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(applied.status, testCase.capabilityName).toBe("SUCCESS");
      expect(runtime.tableMethodCalls.at(-1)?.method, testCase.capabilityName).toBe(testCase.expectedMethod);
    }
  });

  it("rejects incomplete table append rows before creating a preview", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Append a row to Transactions with one missing value.",
      mode: "preview_update",
      target: { tableName: "Transactions" },
      values: { rows: [["2026-01-15", "Globex", "Support"]] }
    });

    expect(preview.status).toBe("VALIDATION_FAILED");
    expect(preview.summary).toContain("has 3 value(s)");
    expect(preview.summary).toContain("4 column(s)");
    expect(preview.nextAction).toBe("ask_user");
    expect(runtime.tableMethodCalls.length).toBe(0);
  });

  it("routes every covered sheet-core mutation to its internal operation", async () => {
    const cases: Array<{
      capabilityName: string;
      expectedOperationKind: string;
      input: Parameters<AgentOrchestrator["run"]>[0];
    }> = [
      {
        capabilityName: "excel.sheet.create",
        expectedOperationKind: "sheet.create",
        input: {
          request: "Create a new sheet",
          mode: "preview_update",
          intent: { action: "create_sheet" },
          values: { sheetName: "New Report" }
        }
      },
      {
        capabilityName: "excel.sheet.copy",
        expectedOperationKind: "sheet.copy",
        input: {
          request: "Copy the Report sheet",
          mode: "preview_update",
          intent: { action: "copy_sheet" },
          target: { sheetName: "Report" },
          values: { newSheetName: "Report Copy" }
        }
      },
      {
        capabilityName: "excel.sheet.rename",
        expectedOperationKind: "sheet.rename",
        input: {
          request: "Rename the Report sheet",
          mode: "preview_update",
          intent: { action: "rename_sheet" },
          target: { sheetName: "Report" },
          values: { newSheetName: "Report Renamed" }
        }
      },
      {
        capabilityName: "excel.sheet.delete",
        expectedOperationKind: "sheet.delete",
        input: {
          request: "Delete the Report sheet",
          mode: "preview_update",
          intent: { action: "delete_sheet" },
          target: { sheetName: "Report" }
        }
      },
      {
        capabilityName: "excel.sheet.hide",
        expectedOperationKind: "sheet.hide",
        input: {
          request: "Hide the Report sheet",
          mode: "preview_update",
          intent: { action: "hide_sheet" },
          target: { sheetName: "Report" }
        }
      },
      {
        capabilityName: "excel.sheet.unhide",
        expectedOperationKind: "sheet.unhide",
        input: {
          request: "Unhide the Report sheet",
          mode: "preview_update",
          intent: { action: "unhide_sheet" },
          target: { sheetName: "Report" }
        }
      },
      {
        capabilityName: "excel.sheet.protect",
        expectedOperationKind: "sheet.protect",
        input: {
          request: "Protect the Report sheet",
          mode: "preview_update",
          intent: { action: "protect_sheet" },
          target: { sheetName: "Report" },
          values: { password: "unit" }
        }
      },
      {
        capabilityName: "excel.sheet.unprotect",
        expectedOperationKind: "sheet.unprotect",
        input: {
          request: "Unprotect the Report sheet",
          mode: "preview_update",
          intent: { action: "unprotect_sheet" },
          target: { sheetName: "Report" },
          values: { password: "unit" }
        }
      },
      {
        capabilityName: "excel.sheet.clear",
        expectedOperationKind: "sheet.clear",
        input: {
          request: "Clear the Report sheet",
          mode: "preview_update",
          intent: { action: "clear_sheet" },
          target: { sheetName: "Report" },
          values: { applyTo: "contents" }
        }
      },
      {
        capabilityName: "excel.sheet.set_tab_color",
        expectedOperationKind: "sheet.set_tab_color",
        input: {
          request: "Set the Report sheet tab color",
          mode: "preview_update",
          intent: { action: "set_sheet_tab_color" },
          target: { sheetName: "Report" },
          values: { color: "#4472C4" }
        }
      }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const preview = await agent.run(testCase.input);

      expect(preview.status, testCase.capabilityName).toBe("PREVIEW_READY");

      const applied = await agent.run({
        request: `Apply ${testCase.capabilityName}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(applied.status, testCase.capabilityName).toBe("SUCCESS");
      expect(runtime.lastBatchOperations[0]?.kind, testCase.capabilityName).toBe(testCase.expectedOperationKind);
    }
  });

  it("routes covered workbook snapshot and backup actions to runtime methods", async () => {
    const cases: Array<{
      capabilityName: string;
      intentAction: "create_snapshot" | "create_backup";
      expectedMethod: string;
    }> = [
      { capabilityName: "excel.workbook.snapshot", intentAction: "create_snapshot", expectedMethod: "workbook.snapshot" },
      { capabilityName: "excel.workbook.create_backup", intentAction: "create_backup", expectedMethod: "workbook.create_backup" }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const preview = await agent.run({
        request: `Create ${testCase.capabilityName}`,
        mode: "preview_update",
        intent: { action: testCase.intentAction },
        target: { sheetName: "Data", range: "A1:D4" },
        values: { reason: `Unit ${testCase.capabilityName}` }
      });
      const applied = await agent.run({
        request: `Apply ${testCase.capabilityName}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status, testCase.capabilityName).toBe("PREVIEW_READY");
      expect(applied.status, testCase.capabilityName).toBe("SUCCESS");
      expect(runtime.runtimeMethodCalls[testCase.expectedMethod], testCase.capabilityName).toBe(1);
      expect(runtime.writeBatchCount, testCase.capabilityName).toBe(0);
    }
  });

  it("previews duplicate sheet template cleanup as one pending operation", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
  
      const result = await agent.run({
        request: "Can you duplicate Report sheet, remove data and keep only template?",
        target: { sheetName: "Report" },
        values: { dataRegions: ["B2:B3"] }
      });
      const applied = await agent.run({
        request: "Apply template cleanup",
        mode: "apply_update",
        operationId: result.operationId,
        confirmationToken: result.confirmationToken
      });
  
      expect(result.status).toBe("PREVIEW_READY");
      expect((result.answer as any).kind).toBe("template_cleanup_preview");
      expect((result.answer as any).operationKind).toBe("sheet.copy_clean_data_regions");
      expect((result.answer as any).dataRegions).toEqual(["B2:B3"]);
      expect((result.answer as any).sourceSheetName).toBe("Report");
      expect(result.nextAction).toBe("call_apply_update");
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.writeBatchCount).toBe(1);
      expect(runtime.lastBatchOperations).toHaveLength(1);
      expect(runtime.lastBatchOperations[0]?.kind).toBe("sheet.copy_clean_data_regions");
    });

  it("previews and applies table appends through the agent surface", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
  
      const preview = await agent.run({
        request: "Append transaction rows to the Transactions table",
        mode: "preview_update",
        target: { tableName: "Transactions" },
        values: { rows: [["2026-06-01", 204, "71-4653", "Company gas top-up"]] }
      });
      const applied = await agent.run({
        request: "Apply append",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });
  
      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe("table_append_preview");
      expect((preview.answer as any).tableName).toBe("Transactions");
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.appendTableRowCount).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("does not auto-apply ambiguous natural-language updates", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
  
      const result = await agent.run({
        request: "Update financial 2026",
        values: { values: [[999]] }
      });
  
      expect(result.status).toBe("AMBIGUOUS_TARGET");
      expect(result.nextAction).toBe("call_with_target");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("does not auto-apply formula-sensitive requests as value writes", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
  
      const result = await agent.run({
        request: "Fix formula in Report A12",
        target: { sheetName: "Report", range: "A12" },
        values: { values: [[100]] }
      });
  
      expect(result.status).toBe("NEEDS_INPUT");
      expect(result.nextAction).toBe("manual_review");
      expect(result.telemetry.safetyDecision).toBe("manual_review:advanced_workflow");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("blocks formula-like values in generic value previews", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
  
      const result = await agent.run({
        request: "Change Data B2",
        mode: "preview_update",
        target: { sheetName: "Data", range: "B2" },
        values: { values: [["=SUM(A1:A2)"]] }
      });
  
      expect(result.status).toBe("VALIDATION_FAILED");
      expect(result.nextAction).toBe("manual_review");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("does not auto-apply sparse broad writes", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
  
      const result = await agent.run({
        request: "Change Report B1:C4",
        target: { sheetName: "Report", range: "B1:C4" },
        values: { values: [[1]] }
      });
  
      expect(result.status).toBe("VALIDATION_FAILED");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("reports explicit route metadata for auto mutation routing", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
  
      const result = await agent.run({
        request: "Update Data B2",
        target: { sheetName: "Data", range: "B2" },
        values: { values: [[123]] }
      });
  
      expect(result.status).toBe("PREVIEW_READY");
      expect(result.telemetry.routeMode).toBe("preview_update");
      expect(result.telemetry.routeMatchedRule).toBe("mutation.keyword");
      expect(result.telemetry.routeConfidence).toBeGreaterThan(0);
      expect(result.telemetry.operationRisk).toBe("small_value_write");
      expect(result.telemetry.targetFingerprintStatus).toBe("matched");
    });

  it("carries operation risk telemetry through apply", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
  
      const preview = await agent.run({
        request: "Format header row on Data",
        mode: "preview_update",
        target: { sheetName: "Data", range: "A1:D1" }
      });
      const applied = await agent.run({
        request: "Apply style update",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });
  
      expect(preview.status).toBe("PREVIEW_READY");
      expect(preview.telemetry.operationRisk).toBe("safe_format");
      expect(applied.status).toBe("SUCCESS");
      expect(applied.telemetry.operationRisk).toBe("safe_format");
      expect((applied.answer as any).operationRisk).toBe("safe_format");
    });

  it("routes natural autofilter and border requests to range handlers instead of sheet creation", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const filter = await agent.run({
        request: "Add autofilter to the header row of Booking sheet range A1:X7.",
        mode: "preview_update",
        target: { sheetName: "Data", range: "A1:D4" }
      });
      const borders = await agent.run({
        request: "Add borders to Booking sheet range A1:X7. Set EdgeBottom and EdgeRight as thin continuous black borders.",
        mode: "preview_update",
        target: { sheetName: "Data", range: "A1:D4" }
      });
      const explicitFormat = await new AgentOrchestrator(new FakeAgentRuntime() as any).run({
        request: "format_range action: Apply borders to range Data!A1:D4. Set EdgeBottom and EdgeRight borders.",
        mode: "preview_update",
        target: { sheetName: "Data", range: "A1:D4" }
      });

      expect(filter.status).toBe("PREVIEW_READY");
      expect((filter.answer as any).kind).toBe("filter_preview");
      expect(filter.telemetry.actionHandlerId).toBe("filter_range");
      expect(filter.telemetry.operationRisk).toBe("safe_format");
      expect((filter.answer as any).kind).not.toBe("sheet_create_preview");

      expect(borders.status).toBe("PREVIEW_READY");
      expect((borders.answer as any).kind).toBe("style_preview");
      expect(borders.telemetry.actionHandlerId).toBe("format_range");
      expect((borders.answer as any).kind).not.toBe("sheet_create_preview");

      expect(explicitFormat.status).toBe("PREVIEW_READY");
      expect((explicitFormat.answer as any).kind).toBe("style_preview");
      expect(explicitFormat.telemetry.actionHandlerId).toBe("format_range");
    });

  it("keeps parse_dates patch targets exact instead of broadening to the whole sheet", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Parse text dates in Data columns A and C rows 2-4.",
        mode: "preview_update",
        intent: { action: "parse_dates" },
        target: { sheetName: "Data", range: "A1:D4" },
        values: {
          patches: [
            { target: { sheetName: "Data", range: "A2:A4" }, numberFormat: "dd/mm/yyyy" },
            { target: { sheetName: "Data", range: "C2:C4" }, numberFormat: "dd/mm/yyyy" }
          ]
        }
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe("cleaning_preview");
      expect((preview.answer as any).grouped).toBe(true);
      expect((preview.answer as any).requests.map((request: any) => request.address)).toEqual(["A2:A4", "C2:C4"]);
      expect(preview.proof.map((proof) => proof.range)).toEqual(["A2:A4", "C2:C4"]);
      expect(preview.changes?.map((change) => change.range)).toEqual(["A2:A4", "C2:C4"]);

      const applied = await agent.run({
        request: "Apply exact date parsing.",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(applied.status).toBe("SUCCESS");
      expect(runtime.runtimeMethodCalls["clean.parse_dates"]).toBe(2);
    });

  it("normalizes agent-friendly table filter shorthands before apply", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      agent.metadataCache.set(createCachedMetadata("wbctx_filter_alias"));

      const preview = await agent.run({
        request: "Filter the Transactions table to Status Open",
        mode: "preview_update",
        workbookContextId: "wbctx_filter_alias",
        intent: { action: "filter_range" },
        target: { tableName: "Transactions" },
        values: { filters: [{ column: "Status", filterType: "text", value: "Open" }] }
      });
      const applied = await agent.run({
        request: "Apply filter",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.tableMethodCalls.at(-1)).toEqual({
        method: "table.apply_filters",
        request: {
          workbookId: "workbook_agent_unit",
          tableName: "Transactions",
          filters: [{ column: "Status", criteria: { filterOn: "Values", values: ["Open"] } }]
        }
      });
    });

  it("normalizes criterion table filters before apply", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      agent.metadataCache.set(createCachedMetadata("wbctx_filter_criterion"));

      const preview = await agent.run({
        request: "Filter the Transactions table to Status Open",
        mode: "preview_update",
        workbookContextId: "wbctx_filter_criterion",
        intent: { action: "filter_range" },
        target: { tableName: "Transactions" },
        values: { filters: [{ column: "Status", criterion: "Open" }] }
      });
      const applied = await agent.run({
        request: "Apply filter",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(applied.status).toBe("SUCCESS");
      expect(runtime.tableMethodCalls.at(-1)?.request.filters[0].criteria).toEqual({ filterOn: "Values", values: ["Open"] });
    });

  it("accepts rowStart and rowEnd aliases for compact table reads", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      agent.metadataCache.set(createCachedMetadata("wbctx_table_page_alias"));

      const result = await agent.run({
        request: "Read rows 2 through 2 from Transactions",
        mode: "answer",
        workbookContextId: "wbctx_table_page_alias",
        target: { tableName: "Transactions" },
        values: { rowStart: 2, rowEnd: 2 }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).rowOffset).toBe(1);
      expect((result.answer as any).rowLimit).toBe(1);
      expect((result.answer as any).values).toBeUndefined();
      const resultId = String((result.answer as any).resultUri).split("/").pop()!;
      expect((agent.getResultResource(resultId) as any).answer.values).toEqual([["2026-06-02", "A-101", 456, "Closed"]]);
    });

  it("normalizes structured table sort values", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      agent.metadataCache.set(createCachedMetadata("wbctx_sort_alias"));

      const preview = await agent.run({
        request: "Sort Transactions by status descending",
        mode: "preview_update",
        workbookContextId: "wbctx_sort_alias",
        intent: { action: "sort_table" },
        target: { tableName: "Transactions" },
        values: { sortBy: "Status", direction: "descending" }
      });
      const applied = await agent.run({
        request: "Apply sort",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).sortField).toBe("Status");
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.tableMethodCalls.at(-1)?.request.fields).toEqual([{ key: 3, ascending: false }]);
    });
});
