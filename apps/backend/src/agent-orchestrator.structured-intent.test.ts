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
