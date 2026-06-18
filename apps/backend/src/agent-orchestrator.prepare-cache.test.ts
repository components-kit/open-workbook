import { describe, expect, it } from "vitest";
import { AgentOrchestrator } from "./agent-orchestrator.js";
import { FakeAgentRuntime, createCachedMetadata, selectionInfo, sheets } from "./agent-orchestrator.test-support.js";

describe("AgentOrchestrator Prepare Cache", () => {
  it("reuses prepared metadata and reports cache telemetry", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
  
      const first = await agent.run({ request: "Prepare workbook", mode: "prepare" });
      const readCallsAfterFirstPrepare = runtime.readBatchCount;
      const second = await agent.run({ request: "Prepare workbook again", mode: "prepare", workbookContextId: first.workbookContextId });
  
      expect(first.status).toBe("SUCCESS");
      expect(first.telemetry.cacheHit).toBe(false);
      expect(first.telemetry.metadataCacheStatus).toBe("miss");
      expect(readCallsAfterFirstPrepare).toBe(0);
      expect(second.telemetry.cacheHit).toBe(true);
      expect(second.telemetry.metadataCacheStatus).toBe("hit");
      expect(runtime.readBatchCount).toBe(readCallsAfterFirstPrepare);
    });

  it("answers vague workbook file reviews from complete structure metadata without sheet sampling", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
  
      const result = await agent.run({ request: "Can you look into transactions.xlsx?" });
  
      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("workbook_overview");
      expect((result.answer as any).sheets.map((sheet: any) => sheet.name)).toEqual(sheets.map((sheet) => sheet.name));
      expect(result.telemetry.internalReadCount).toBe(0);
      expect(runtime.readBatchCount).toBe(0);
    });

  it("upgrades prepared structure metadata when a later answer needs sheet samples", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
  
      const prepared = await agent.run({ request: "Prepare workbook", mode: "prepare" });
      const answered = await agent.run({
        request: "Read Apr 2026 invoice rows",
        mode: "answer",
        workbookContextId: prepared.workbookContextId
      });
  
      expect(prepared.status).toBe("SUCCESS");
      expect(answered.status).toBe("SUCCESS");
      expect(answered.proof[0]?.sheetName).toBe("Apr 2026");
      expect(answered.proof[0]?.range).toBe("O1:AE244");
      expect(runtime.readBatchCount).toBeGreaterThan(0);
    });

  it("bounds candidates and reports token savings when a response budget is provided", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
  
      const result = await agent.run({
        request: "Find amount status region report data total formula",
        mode: "find",
        budget: { maxExamples: 2, maxPayloadBytes: 850 }
      });
  
      expect(result.status).toBe("SUCCESS");
      expect(result.candidates?.length).toBeLessThanOrEqual(2);
      expect(result.proof.length).toBeLessThanOrEqual(2);
      expect(result.telemetry.candidateCount).toBe(result.candidates?.length ?? 0);
      expect(result.telemetry.payloadBytes).toBeLessThanOrEqual(1000);
      expect(result.telemetry.estimatedTokensSaved).toBeGreaterThanOrEqual(0);
    });
});
