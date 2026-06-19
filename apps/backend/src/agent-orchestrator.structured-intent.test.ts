import { describe, expect, it } from "vitest";
import { AgentOrchestrator } from "./agent-orchestrator.js";
import { FakeAgentRuntime, createCachedMetadata, selectionInfo, sheets } from "./agent-orchestrator.test-support.js";

describe("AgentOrchestrator Structured Intent", () => {
  it("uses caller structured intent to route simple auto requests without backend LLM parsing", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Do this to the selected target",
        intent: { action: "format_range", confidence: 0.92, reason: "Caller parsed a header-formatting request." },
        target: { sheetName: "Data", range: "A1:D1" }
      });

      expect(result.status).toBe("PREVIEW_READY");
      expect((result.answer as any).kind).toBe("style_preview");
      expect(result.telemetry.routeMatchedRule).toBe("caller_intent.action");
      expect(result.telemetry.intentSource).toBe("caller_structured");
      expect(result.telemetry.intentAction).toBe("format_range");
      expect(result.telemetry.intentAccepted).toBe(true);
      expect(result.telemetry.operationRisk).toBe("safe_format");
      expect(result.telemetry.actionHandlerId).toBe("format_range");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("uses caller structured intent for schema reads when request text is minimal", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Inspect this",
        intent: { action: "read_schema", confidence: 0.88 },
        target: { tableName: "Transactions" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("table_schema");
      expect(result.telemetry.routeMode).toBe("answer");
      expect(result.telemetry.intentAction).toBe("read_schema");
      expect(result.telemetry.internalReadCount).toBe(0);
    });

  it("uses caller structured intent for workbook answer actions", async () => {
      const cases: Array<{ action: "list_open_workbooks" | "get_workbook_info" | "refresh_workbook_snapshot" | "get_workbook_snapshot" | "detect_external_changes" | "export_local_config" | "read_embedded_local_config"; expectedKind: string; runtimeCall: string; values?: Record<string, unknown> }> = [
        { action: "list_open_workbooks", expectedKind: "open_workbooks", runtimeCall: "workbook.list_open_workbooks" },
        { action: "get_workbook_info", expectedKind: "workbook_info", runtimeCall: "workbook.get_workbook_info" },
        { action: "refresh_workbook_snapshot", expectedKind: "workbook_snapshot_refresh", runtimeCall: "workbook.refresh_snapshot", values: { snapshotId: "snapshot_agent_unit" } },
        { action: "get_workbook_snapshot", expectedKind: "workbook_snapshot", runtimeCall: "workbook.get_snapshot", values: { snapshotId: "snapshot_agent_unit" } },
        { action: "detect_external_changes", expectedKind: "workbook_external_changes", runtimeCall: "workbook.detect_external_changes", values: { snapshotId: "snapshot_agent_unit" } },
        { action: "export_local_config", expectedKind: "workbook_local_config_export", runtimeCall: "workbook.export_local_config" },
        { action: "read_embedded_local_config", expectedKind: "workbook_embedded_local_config", runtimeCall: "workbook.read_embedded_local_config" }
      ];

      for (const testCase of cases) {
        const runtime = new FakeAgentRuntime();
        const agent = new AgentOrchestrator(runtime as any);
        const result = await agent.run({
          request: "Inspect workbook",
          intent: { action: testCase.action, confidence: 0.91 },
          values: testCase.values
        });

        expect(result.status).toBe("SUCCESS");
        expect((result.answer as any).kind).toBe(testCase.expectedKind);
        expect(result.telemetry.intentAction).toBe(testCase.action);
        expect(result.telemetry.routeMode).toBe("answer");
        expect(runtime.runtimeMethodCalls[testCase.runtimeCall]).toBe(1);
        expect(runtime.writeBatchCount).toBe(0);
      }
    });

  it("uses caller structured intent for range metadata answer actions", async () => {
      const cases: Array<{ action: "read_hyperlinks" | "read_comments" | "read_notes" | "read_merged_cells" | "read_data_validation" | "read_conditional_formatting" | "search_range" | "find_blank_cells" | "find_range_errors"; runtimeCall: string; values?: Record<string, unknown> }> = [
        { action: "read_hyperlinks", runtimeCall: "range.read_hyperlinks" },
        { action: "read_comments", runtimeCall: "range.read_comments" },
        { action: "read_notes", runtimeCall: "range.read_notes" },
        { action: "read_merged_cells", runtimeCall: "range.read_merged_cells" },
        { action: "read_data_validation", runtimeCall: "range.read_data_validation" },
        { action: "read_conditional_formatting", runtimeCall: "range.read_conditional_formatting" },
        { action: "search_range", runtimeCall: "range.search", values: { text: "Open" } },
        { action: "find_blank_cells", runtimeCall: "range.find_blank_cells" },
        { action: "find_range_errors", runtimeCall: "range.find_errors" }
      ];

      for (const testCase of cases) {
        const runtime = new FakeAgentRuntime();
        const agent = new AgentOrchestrator(runtime as any);
        const result = await agent.run({
          request: "Inspect range metadata",
          mode: "answer",
          intent: { action: testCase.action, confidence: 0.91 },
          target: { sheetName: "Data", range: "A1:D4" },
          values: testCase.values
        });

        expect(result.status).toBe("SUCCESS");
        expect((result.answer as any).kind).toBe("range_metadata");
        expect(result.telemetry.intentAction).toBe(testCase.action);
        expect(result.telemetry.routeMode).toBe("answer");
        expect(runtime.runtimeMethodCalls[testCase.runtimeCall]).toBe(1);
        expect(runtime.writeBatchCount).toBe(0);
      }
    });

  it("uses caller structured intent for compact range reads", async () => {
      const cases: Array<"read_range_compact" | "get_range_summary"> = ["read_range_compact", "get_range_summary"];

      for (const action of cases) {
        const runtime = new FakeAgentRuntime();
        const agent = new AgentOrchestrator(runtime as any);
        const result = await agent.run({
          request: "Read this range",
          mode: "answer",
          intent: { action, confidence: 0.91 },
          target: { sheetName: "Data", range: "A1:D4" }
        });

        expect(result.status).toBe("SUCCESS");
        expect((result.answer as any).kind).toBe(action === "get_range_summary" ? "range_summary" : "range_compact");
        expect(result.telemetry.intentAction).toBe(action);
        expect(result.telemetry.routeMode).toBe("answer");
        expect(runtime.readBatchCount).toBeGreaterThan(0);
        expect(runtime.writeBatchCount).toBe(0);
      }
    });

  it("uses caller structured intent for workflow plan reports", async () => {
      const cases: Array<{ action: "prepare_session" | "create_formula_sheet" | "create_template_report" | "create_pivot_chart_summary" | "preview_risky_edit" | "inspect_analyze" | "rollback_validate"; mutatesWorkbook: boolean }> = [
        { action: "prepare_session", mutatesWorkbook: false },
        { action: "create_formula_sheet", mutatesWorkbook: true },
        { action: "create_template_report", mutatesWorkbook: true },
        { action: "create_pivot_chart_summary", mutatesWorkbook: true },
        { action: "preview_risky_edit", mutatesWorkbook: true },
        { action: "inspect_analyze", mutatesWorkbook: false },
        { action: "rollback_validate", mutatesWorkbook: true }
      ];

      for (const testCase of cases) {
        const runtime = new FakeAgentRuntime();
        const agent = new AgentOrchestrator(runtime as any);
        const result = await agent.run({
          request: "Plan the workflow",
          mode: "answer",
          intent: { action: testCase.action, confidence: 0.91 },
          target: { sheetName: "Data", range: "A1:D4" }
        });

        expect(result.status).toBe("SUCCESS");
        expect((result.answer as any).kind).toBe("workflow_plan");
        expect((result.answer as any).action).toBe(testCase.action);
        expect((result.answer as any).mutatesWorkbook).toBe(testCase.mutatesWorkbook);
        expect((result.answer as any).requiredCapabilities.length).toBeGreaterThan(0);
        expect(result.telemetry.intentAction).toBe(testCase.action);
        expect(result.telemetry.routeMode).toBe("answer");
        expect(result.nextAction).toBe(testCase.mutatesWorkbook ? "manual_review" : "answer_now");
        expect(runtime.writeBatchCount).toBe(0);
      }
    });

  it("uses caller structured intent for formula previews without relying on request wording", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Put this in the cell",
        mode: "preview_update",
        intent: { action: "write_formulas", confidence: 0.9 },
        target: { sheetName: "Report", range: "B2" },
        values: { values: [["=SUM(Data!C2:C4)"]] }
      });

      expect(result.status).toBe("PREVIEW_READY");
      expect((result.answer as any).kind).toBe("formula_preview");
      expect(result.telemetry.intentAction).toBe("write_formulas");
      expect(result.telemetry.operationRisk).toBe("formula_write");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("uses caller structured intent for formula pattern reads", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Inspect this formula area",
        mode: "answer",
        intent: { action: "read_formula_patterns", confidence: 0.91 },
        target: { sheetName: "Report", range: "B2" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("formula_patterns");
      expect((result.answer as any).patterns.formulaCount).toBe(1);
      expect(result.proof[0]).toMatchObject({ sheetName: "Report", range: "B2" });
      expect(result.telemetry.intentAction).toBe("read_formula_patterns");
      expect(result.telemetry.internalReadCount).toBe(1);
      expect(runtime.runtimeMethodCalls["formula.read_patterns"]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("uses caller structured intent for style fingerprint reads", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Inspect styling",
        mode: "answer",
        intent: { action: "read_style_fingerprint", confidence: 0.91 },
        target: { sheetName: "Report", range: "A1:B4" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("style_fingerprint");
      expect(result.proof[0]).toMatchObject({ sheetName: "Report", range: "A1:B4" });
      expect(result.telemetry.intentAction).toBe("read_style_fingerprint");
      expect(result.telemetry.internalReadCount).toBe(1);
      expect(runtime.runtimeMethodCalls["style.get_fingerprint"]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("uses caller structured intent for style fingerprint comparisons", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Compare styling",
        mode: "answer",
        intent: { action: "compare_style_fingerprint", confidence: 0.91 },
        values: {
          source: { sheetName: "Data", range: "A1:B4" },
          destination: { sheetName: "Report", range: "A1:B4" },
          dimensions: ["fills", "fonts"]
        }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("style_compare");
      expect(result.proof[0]).toMatchObject({ sheetName: "Data", range: "A1:B4" });
      expect(result.proof[1]).toMatchObject({ sheetName: "Report", range: "A1:B4" });
      expect(result.telemetry.intentAction).toBe("compare_style_fingerprint");
      expect(result.telemetry.internalReadCount).toBe(2);
      expect(runtime.runtimeMethodCalls["style.compare_fingerprint"]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("uses caller structured intent for theme capability reports", async () => {
      const cases: Array<{ action: "get_theme" | "apply_theme"; runtimeCall: string; values?: Record<string, unknown> }> = [
        { action: "get_theme", runtimeCall: "style.get_theme" },
        { action: "apply_theme", runtimeCall: "style.apply_theme", values: { theme: { name: "Corporate" } } }
      ];

      for (const testCase of cases) {
        const runtime = new FakeAgentRuntime();
        const agent = new AgentOrchestrator(runtime as any);

        const result = await agent.run({
          request: "Handle workbook theme",
          intent: { action: testCase.action, confidence: 0.91 },
          ...(testCase.values ? { values: testCase.values } : {})
        });

        expect(result.status).toBe("VALIDATION_FAILED");
        expect((result.answer as any).kind).toBe("theme_capability_report");
        expect(result.telemetry.intentAction).toBe(testCase.action);
        expect(result.telemetry.routeMode).toBe("answer");
        expect(result.nextAction).toBe("manual_review");
        expect(runtime.runtimeMethodCalls[testCase.runtimeCall]).toBe(1);
        expect(runtime.writeBatchCount).toBe(0);
      }
    });

  it("uses caller structured intent for cleaning inspections", async () => {
      const cases: Array<{ action: "detect_header_row" | "detect_outliers" | "fuzzy_match"; runtimeCall: string; values?: Record<string, unknown> }> = [
        { action: "detect_header_row", runtimeCall: "clean.detect_header_row" },
        { action: "detect_outliers", runtimeCall: "clean.detect_outliers", values: { columnIndex: 2, threshold: 2 } },
        { action: "fuzzy_match", runtimeCall: "clean.fuzzy_match", values: { lookupValues: ["Paid", "Open"] } }
      ];

      for (const testCase of cases) {
        const runtime = new FakeAgentRuntime();
        const agent = new AgentOrchestrator(runtime as any);

        const result = await agent.run({
          request: "Inspect data quality",
          mode: "answer",
          intent: { action: testCase.action, confidence: 0.91 },
          target: { sheetName: "Data", range: "A1:D4" },
          values: testCase.values
        });

        expect(result.status).toBe("SUCCESS");
        expect((result.answer as any).kind).toBe("cleaning_report");
        expect(result.proof[0]).toMatchObject({ sheetName: "Data", range: "A1:D4" });
        expect(result.telemetry.intentAction).toBe(testCase.action);
        expect(result.telemetry.internalReadCount).toBe(1);
        expect(runtime.runtimeMethodCalls[testCase.runtimeCall]).toBe(1);
        expect(runtime.writeBatchCount).toBe(0);
      }
    });

  it("uses caller structured intent for unsupported repair reports", async () => {
      const cases: Array<{ action: "repair_filters_from_template" | "repair_print_layout" | "repair_named_ranges" | "repair_formula_errors" | "repair_merged_cells"; runtimeCall: string }> = [
        { action: "repair_filters_from_template", runtimeCall: "repair.filters_from_template" },
        { action: "repair_print_layout", runtimeCall: "repair.print_layout" },
        { action: "repair_named_ranges", runtimeCall: "repair.named_ranges" },
        { action: "repair_formula_errors", runtimeCall: "repair.formula_errors" },
        { action: "repair_merged_cells", runtimeCall: "repair.merged_cells" }
      ];

      for (const testCase of cases) {
        const runtime = new FakeAgentRuntime();
        const agent = new AgentOrchestrator(runtime as any);

        const result = await agent.run({
          request: "Check repair availability",
          mode: "answer",
          intent: { action: testCase.action, confidence: 0.91 },
          target: { sheetName: "Report" },
          values: { templateId: "template_unit" }
        });

        expect(result.status).toBe("VALIDATION_FAILED");
        expect((result.answer as any).kind).toBe("repair_report");
        expect(result.telemetry.intentAction).toBe(testCase.action);
        expect(result.telemetry.internalReadCount).toBe(1);
        expect(runtime.runtimeMethodCalls[testCase.runtimeCall]).toBe(1);
      }
    });

  it("uses caller structured intent for formula dependency graph reads", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Trace dependencies for this formula",
        mode: "answer",
        intent: { action: "get_formula_dependency_graph", confidence: 0.9 },
        target: { sheetName: "Report", range: "B2" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("formula_dependency_graph");
      expect((result.answer as any).graph.edges[0].to.address).toBe("C2:C4");
      expect(result.telemetry.intentAction).toBe("get_formula_dependency_graph");
      expect(result.telemetry.internalReadCount).toBe(1);
      expect(runtime.runtimeMethodCalls["formula.get_dependency_graph"]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("uses caller structured intent for formula precedent tracing", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Trace precedents for this formula",
        mode: "answer",
        intent: { action: "trace_formula_precedents", confidence: 0.9 },
        target: { sheetName: "Report", range: "B2" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("trace_formula_precedents");
      expect((result.answer as any).result.edges[0].kind).toBe("precedent");
      expect(result.telemetry.intentAction).toBe("trace_formula_precedents");
      expect(result.telemetry.internalReadCount).toBe(1);
      expect(runtime.runtimeMethodCalls["formula.trace_precedents"]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("uses caller structured intent for formula dependent tracing", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Trace dependents for this input range",
        mode: "answer",
        intent: { action: "trace_formula_dependents", confidence: 0.9 },
        target: { sheetName: "Data", range: "C2:C4" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("trace_formula_dependents");
      expect((result.answer as any).result.edges[0].kind).toBe("dependent");
      expect(result.telemetry.intentAction).toBe("trace_formula_dependents");
      expect(result.telemetry.internalReadCount).toBe(1);
      expect(runtime.runtimeMethodCalls["formula.trace_dependents"]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("uses caller structured intent for formula validation", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Validate formulas in this cell",
        mode: "validate",
        intent: { action: "validate_formula_range", confidence: 0.9 },
        target: { sheetName: "Report", range: "B2" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("validate_formula_range");
      expect(result.telemetry.intentAction).toBe("validate_formula_range");
      expect(result.telemetry.validationStatus).toBe("passed");
      expect(runtime.runtimeMethodCalls["validate.formulas"]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("uses caller structured intent for formula error scans", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Find formula errors",
        mode: "answer",
        intent: { action: "find_formula_errors", confidence: 0.9 },
        target: { sheetName: "Report", range: "B2" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("formula_errors");
      expect(result.telemetry.intentAction).toBe("find_formula_errors");
      expect(result.telemetry.internalReadCount).toBe(1);
      expect(runtime.runtimeMethodCalls["validate.formulas"]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("uses caller structured intent for formula explanation", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Explain this formula",
        mode: "answer",
        intent: { action: "explain_formula", confidence: 0.9 },
        target: { sheetName: "Report", range: "B2" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("formula_explain");
      expect((result.answer as any).result.summary.functions).toContain("SUM");
      expect(result.telemetry.intentAction).toBe("explain_formula");
      expect(result.telemetry.internalReadCount).toBe(1);
      expect(runtime.runtimeMethodCalls["formula.read_patterns"]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("uses caller structured intent for named item reads", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read this named item",
        mode: "answer",
        intent: { action: "read_named_item", confidence: 0.9 },
        target: { candidateId: "name:RevenueTotal" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("named_item");
      expect((result.answer as any).result.name.name).toBe("RevenueTotal");
      expect(result.proof[0]).toMatchObject({ sheetName: "Report", range: "B2" });
      expect(result.telemetry.intentAction).toBe("read_named_item");
      expect(result.telemetry.internalReadCount).toBe(1);
      expect(runtime.runtimeMethodCalls["names.get"]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("uses caller structured intent for registered region reads", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read this registered region",
        mode: "answer",
        intent: { action: "read_region", confidence: 0.9 },
        target: { candidateId: "name:InputRegion" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("region");
      expect((result.answer as any).result.region.name).toBe("InputRegion");
      expect(result.proof[0]).toMatchObject({ sheetName: "Report", range: "B1:B3" });
      expect(result.telemetry.intentAction).toBe("read_region");
      expect(result.telemetry.internalReadCount).toBe(1);
      expect(runtime.runtimeMethodCalls["region.get"]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("uses caller structured intent for snapshot lists", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Show snapshots",
        mode: "answer",
        intent: { action: "list_snapshots", confidence: 0.9 }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("snapshot_list");
      expect((result.answer as any).result.snapshots[0].payloadSummary.rangeSnapshotCount).toBe(1);
      expect((result.answer as any).result.snapshots[0].payload).toBeUndefined();
      expect(result.telemetry.intentAction).toBe("list_snapshots");
      expect(result.telemetry.internalReadCount).toBe(1);
      expect(runtime.runtimeMethodCalls["snapshot.list"]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("uses caller structured intent for compact snapshot reads", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read snapshot metadata",
        mode: "answer",
        intent: { action: "read_snapshot", confidence: 0.9 },
        values: { snapshotId: "snapshot_agent_unit" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("snapshot");
      expect((result.answer as any).snapshot.snapshotId).toBe("snapshot_agent_unit");
      expect((result.answer as any).snapshot.payloadSummary.cellCount).toBe(16);
      expect((result.answer as any).snapshot.payload).toBeUndefined();
      expect(result.proof[0]).toMatchObject({ sheetName: "Data", range: "A1:D4" });
      expect(result.telemetry.intentAction).toBe("read_snapshot");
      expect(result.telemetry.internalReadCount).toBe(1);
      expect(runtime.runtimeMethodCalls["snapshot.get_compact"]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("uses caller structured intent for compact snapshot comparisons", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Compare these snapshots",
        mode: "answer",
        intent: { action: "compare_snapshots", confidence: 0.9 },
        values: { leftSnapshotId: "snapshot_before", rightSnapshotId: "snapshot_after" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("snapshot_compare");
      expect((result.answer as any).result.diff.cellsChanged).toBe(4);
      expect(result.proof[0]).toMatchObject({ sheetName: "Data", range: "A1:D4" });
      expect(result.telemetry.intentAction).toBe("compare_snapshots");
      expect(result.telemetry.internalReadCount).toBe(1);
      expect(runtime.runtimeMethodCalls["snapshot.compare_compact"]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("uses caller structured intent for template metadata reads", async () => {
      const cases: Array<{
        action: "detect_templates" | "list_templates" | "read_template" | "infer_template_regions";
        expectedRuntimeCall: string;
        input?: Partial<Parameters<AgentOrchestrator["run"]>[0]>;
      }> = [
        { action: "detect_templates", expectedRuntimeCall: "template.detect" },
        { action: "list_templates", expectedRuntimeCall: "template.list" },
        { action: "read_template", expectedRuntimeCall: "template.get", input: { values: { templateId: "template_unit" } } },
        { action: "infer_template_regions", expectedRuntimeCall: "template.infer_regions", input: { values: { templateId: "template_unit" } } }
      ];

      for (const testCase of cases) {
        const runtime = new FakeAgentRuntime();
        const agent = new AgentOrchestrator(runtime as any);
        const result = await agent.run({
          request: `Run ${testCase.action}`,
          mode: "answer",
          intent: { action: testCase.action },
          ...(testCase.input ?? {})
        });

        expect(result.status).toBe("SUCCESS");
        expect(result.telemetry.intentAction).toBe(testCase.action);
        expect(runtime.runtimeMethodCalls[testCase.expectedRuntimeCall]).toBe(1);
        expect(runtime.writeBatchCount).toBe(0);
      }
    });

  it("uses caller structured intent for backup lists", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Show backups",
        mode: "answer",
        intent: { action: "list_backups", confidence: 0.9 }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("backup_list");
      expect((result.answer as any).result.backups[0].backup.backupId).toBe("backup_agent_unit");
      expect(result.telemetry.intentAction).toBe("list_backups");
      expect(result.telemetry.internalReadCount).toBe(1);
      expect(runtime.runtimeMethodCalls["backup.list"]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("uses caller structured intent for backup reads", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Read backup metadata",
        mode: "answer",
        intent: { action: "read_backup", confidence: 0.9 },
        values: { backupId: "backup_agent_unit" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("backup");
      expect((result.answer as any).result.backup.backupId).toBe("backup_agent_unit");
      expect(result.telemetry.intentAction).toBe("read_backup");
      expect(result.telemetry.internalReadCount).toBe(1);
      expect(runtime.runtimeMethodCalls["backup.get"]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("uses caller structured intent for backup verification", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Verify backup",
        mode: "answer",
        intent: { action: "verify_backup", confidence: 0.9 },
        values: { backupId: "backup_agent_unit" }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("backup_verify");
      expect((result.answer as any).result.manifest.restoreStatus).toBe("available");
      expect(result.telemetry.intentAction).toBe("verify_backup");
      expect(result.telemetry.internalReadCount).toBe(1);
      expect(runtime.runtimeMethodCalls["backup.verify"]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("uses caller structured intent for validation-specific runtime checks", async () => {
      const cases: Array<{
        action:
          | "validate_workbook"
          | "validate_compact"
          | "validate_sheet"
          | "validate_formula_against_template"
          | "validate_template_consistency"
          | "validate_sheet_against_template"
          | "validate_formulas"
          | "validate_styles"
          | "validate_tables"
          | "validate_table_against_template"
          | "validate_filters"
          | "validate_print_layout"
          | "validate_no_broken_references"
          | "validate_no_formula_errors"
          | "validate_no_unintended_changes";
        expectedRuntimeCall: string;
        input?: Partial<Parameters<AgentOrchestrator["run"]>[0]>;
      }> = [
        { action: "validate_workbook", expectedRuntimeCall: "validate.workbook" },
        { action: "validate_compact", expectedRuntimeCall: "validate.workbook" },
        { action: "validate_sheet", expectedRuntimeCall: "validate.sheet", input: { target: { sheetName: "Data" } } },
        { action: "validate_formula_against_template", expectedRuntimeCall: "formula.validate_against_template", input: { target: { sheetName: "Report", range: "B2:B20" }, values: { templateId: "template_unit" } } },
        { action: "validate_template_consistency", expectedRuntimeCall: "validate.template_consistency", input: { target: { sheetName: "Report" }, values: { templateId: "template_unit" } } },
        { action: "validate_sheet_against_template", expectedRuntimeCall: "template.validate_sheet", input: { target: { sheetName: "Report" }, values: { templateId: "template_unit" } } },
        { action: "validate_formulas", expectedRuntimeCall: "validate.formulas", input: { target: { sheetName: "Report", range: "B2" } } },
        { action: "validate_styles", expectedRuntimeCall: "validate.styles", input: { target: { sheetName: "Report" } } },
        { action: "validate_tables", expectedRuntimeCall: "validate.tables", input: { target: { tableName: "Transactions" } } },
        { action: "validate_table_against_template", expectedRuntimeCall: "table.validate_against_template", input: { target: { tableName: "Transactions" }, values: { templateId: "template_unit" } } },
        { action: "validate_filters", expectedRuntimeCall: "validate.filters", input: { target: { tableName: "Transactions" } } },
        { action: "validate_print_layout", expectedRuntimeCall: "validate.print_layout", input: { values: { templateId: "template_unit", targetSheetName: "Report" } } },
        { action: "validate_no_broken_references", expectedRuntimeCall: "validate.no_broken_references", input: { target: { sheetName: "Report", range: "A1:B20" } } },
        { action: "validate_no_formula_errors", expectedRuntimeCall: "validate.no_formula_errors", input: { target: { sheetName: "Report", range: "B2" } } },
        { action: "validate_no_unintended_changes", expectedRuntimeCall: "validate.no_unintended_changes", input: { values: { leftSnapshotId: "snapshot_before", rightSnapshotId: "snapshot_after" } } }
      ];

      for (const testCase of cases) {
        const runtime = new FakeAgentRuntime();
        const agent = new AgentOrchestrator(runtime as any);

        const result = await agent.run({
          request: `Run ${testCase.action}`,
          mode: "validate",
          intent: { action: testCase.action, confidence: 0.9 },
          ...(testCase.input ?? {})
        });

        expect(result.status).toBe("SUCCESS");
        expect((result.answer as any).kind).toBe(testCase.action);
        expect(result.telemetry.intentAction).toBe(testCase.action);
        expect(result.telemetry.validationStatus).toBe("passed");
        expect(result.telemetry.internalReadCount).toBe(1);
        expect(runtime.runtimeMethodCalls[testCase.expectedRuntimeCall]).toBe(1);
        expect(runtime.writeBatchCount).toBe(0);
      }
    });

  it("uses multilingual caller intent and target hints for value reads", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "ช่วยอ่านยอดรายได้ในชีตเดือนมิถุนายน",
        mode: "answer",
        intent: { action: "read_values", confidence: 0.93, targetHints: ["Financials - June 2026", "June financial sheet", "ชีตเดือนมิถุนายน"] }
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.proof[0]?.sheetName).toBe("Financials - June 2026");
      expect(result.telemetry.intentAction).toBe("read_values");
      expect(result.telemetry.targetHintCount).toBe(3);
      expect(result.telemetry.targetHintUsed).toBe(true);
    });

  it("uses multilingual caller intent for style previews with explicit targets", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "ช่วยจัดรูปแบบแถวหัวตารางให้เด่นขึ้น",
        intent: { action: "format_range", confidence: 0.91, reason: "Caller normalized Thai formatting request." },
        target: { sheetName: "Data", range: "A1:D1" }
      });

      expect(result.status).toBe("PREVIEW_READY");
      expect((result.answer as any).kind).toBe("style_preview");
      expect(result.telemetry.intentAction).toBe("format_range");
      expect(result.telemetry.actionHandlerId).toBe("format_range");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("uses multilingual caller intent for workbook-level operations", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "ช่วยบันทึกไฟล์นี้",
        intent: { action: "save", confidence: 0.95 }
      });

      expect(result.status).toBe("PREVIEW_READY");
      expect((result.answer as any).kind).toBe("workbook.save_preview");
      expect(result.telemetry.intentAction).toBe("save");
      expect(result.telemetry.actionHandlerId).toBe("save_workbook");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("rejects unsupported structured intent actions before previewing mutations", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Do something unsupported",
        intent: { action: "make_everything_magic", confidence: 0.9 } as any,
        target: { sheetName: "Data", range: "A1:D1" }
      });

      expect(result.status).toBe("VALIDATION_FAILED");
      expect((result.answer as any).kind).toBe("intent_rejected");
      expect(result.telemetry.intentSource).toBe("mixed");
      expect(result.telemetry.intentAccepted).toBe(false);
      expect(result.telemetry.intentRejectedReason).toBe("unsupported_intent_action");
      expect(runtime.writeBatchCount).toBe(0);
    });
});
