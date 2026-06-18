import { describe, expect, it } from "vitest";
import { AgentOrchestrator } from "./agent-orchestrator.js";
import { FakeAgentRuntime, createCachedMetadata, selectionInfo, sheets } from "./agent-orchestrator.test-support.js";

describe("AgentOrchestrator Preview Apply Safety", () => {
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
      expect(runtime.lastWriteOperations.map((operation) => operation.target.address)).toEqual(["B2:C2", "B3:C3"]);
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

  it("previews duplicate sheet template cleanup as one pending operation", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
  
      const result = await agent.run({
        request: "Can you duplicate Report sheet, remove data and keep only template?",
        target: { sheetName: "Report" }
      });
  
      expect(result.status).toBe("PREVIEW_READY");
      expect((result.answer as any).kind).toBe("template_cleanup_preview");
      expect((result.answer as any).sourceSheetName).toBe("Report");
      expect(result.nextAction).toBe("call_apply_update");
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
});
