import { describe, expect, it } from "vitest";
import { AgentOrchestrator } from "./agent-orchestrator.js";
import { FakeAgentRuntime, createCachedMetadata, selectionInfo, sheets, workbookId } from "./agent-orchestrator.test-support.js";

describe("AgentOrchestrator Prepare Cache", () => {
  it("reports runtime status through the agent status mode", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.agentExecutionContext = { agentId: "agent_status", agentName: "Status Agent", clientType: "mcp" };
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({ request: "Check Excel connection", mode: "status" });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).activeAddinConnected).toBe(true);
      expect((result.answer as any).collaboration).toMatchObject({
        agentCount: 1,
        activeTaskCount: 0,
        activeLockCount: 0,
        activeTransactionCount: 0,
        openConflictCount: 0
      });
      expect((result.answer as any).collaboration.agents[0]).toMatchObject({ agentId: "agent_status", agentName: "Status Agent" });
      expect(result.resourceLinks[0]?.uri).toBe("excel://runtime/status");
      expect(result.telemetry.metadataCacheStatus).toBe("not_applicable");
      expect(runtime.runtimeMethodCalls["runtime.get_connection_readiness"]).toBe(1);
      expect(runtime.runtimeMethodCalls["runtime.get_status"]).toBe(1);
      expect(runtime.runtimeMethodCalls["runtime.get_collaboration_status"]).toBe(1);
    });

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

  it("invalidates sampled metadata when workbook content version changes", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const first = await agent.run({ request: "Find sections and blocks", mode: "find" });
      const readsAfterFirst = runtime.readBatchCount;
      runtime.workbookContentVersion += 1;
      const second = await agent.run({
        request: "Find sections and blocks again",
        mode: "find",
        workbookContextId: first.workbookContextId
      });

      expect(first.telemetry.metadataDetailLevel).toBe("sampled");
      expect(second.telemetry.cacheHit).toBe(false);
      expect(second.telemetry.metadataCacheStatus).toBe("miss");
      expect(second.telemetry.metadataFreshnessReason).toBe("built sampled metadata");
      expect(runtime.readBatchCount).toBeGreaterThan(readsAfterFirst);
    });

  it("reuses sampled metadata after unrelated changes when target range has no journal overlap", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.getWorkbookChangeJournal = (input: any) => ({
        ok: true,
        workbookId: input.workbookId,
        currentVersion: runtime.workbookContentVersion,
        sinceVersion: input.sinceVersion,
        overlapStatus: "no_overlap",
        entries: []
      });
      const agent = new AgentOrchestrator(runtime as any);

      const first = await agent.run({
        request: "Find sections and blocks",
        mode: "find",
        target: { sheetName: "Data", range: "A1:D4" }
      });
      const readsAfterFirst = runtime.readBatchCount;
      runtime.workbookContentVersion += 1;
      const second = await agent.run({
        request: "Find sections and blocks again",
        mode: "find",
        workbookContextId: first.workbookContextId,
        target: { sheetName: "Data", range: "A1:D4" }
      });

      expect(first.telemetry.metadataDetailLevel).toBe("sampled");
      expect(second.telemetry.cacheHit).toBe(true);
      expect(second.telemetry.metadataCacheStatus).toBe("hit");
      expect(second.telemetry.metadataFreshnessReason).toBe("cached metadata is fresh for target; no overlapping changes since context");
      expect(runtime.readBatchCount).toBe(readsAfterFirst);
    });

  it("reuses workbook context from continuation metadata", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const prepared = await agent.run({ request: "Prepare workbook", mode: "prepare" });
      const answered = await agent.run({
        request: "How many sheets are in this workbook?",
        mode: "answer",
        continuation: prepared.continuation
      });

      expect(prepared.continuation).toMatchObject({ workbookContextId: prepared.workbookContextId, responseMode: "brief" });
      expect(answered.status).toBe("SUCCESS");
      expect(answered.workbookContextId).toBe(prepared.workbookContextId);
      expect(answered.telemetry.cacheHit).toBe(true);
      expect(answered.telemetry.internalReadCount).toBe(0);
      expect((answered.answer as any).sheetCount).toBe(sheets.length);
    });

  it("answers copied workbook context handles without rebuilding workbook metadata", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const prepared = await agent.run({ request: "Prepare workbook", mode: "prepare" });
      const callsAfterPrepare = { ...runtime.runtimeMethodCalls };
      const handled = await agent.run({
        request: `continue using excel://agent/contexts/${prepared.workbookContextId}`,
        mode: "answer",
        responseMode: "verbose"
      });

      expect(handled.status).toBe("SUCCESS");
      expect(handled.workbookContextId).toBe(prepared.workbookContextId);
      expect((handled.answer as any).ok).toBe(true);
      expect((handled.answer as any).workbook.sheetCount).toBe(sheets.length);
      expect(runtime.runtimeMethodCalls["workbook.get_workbook_map"]).toBe(callsAfterPrepare["workbook.get_workbook_map"]);
      expect(runtime.readBatchCount).toBe(0);
    });

  it("prepares workbook discovery metadata through internal workbook, sheet, table, name, and region capabilities", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.collaborationStatus = {
        ok: true,
        workbookId,
        agents: [{ agentId: "agent_prepare", agentName: "Prepare Agent", clientType: "mcp", status: "active" }],
        tasks: [{ taskId: "task_prepare", agentId: "agent_prepare", workbookId, status: "running", title: "Prepare workbook" }],
        locks: [{ lockId: "lock_data", workbookId, ownerAgentId: "agent_prepare", status: "active", scopes: [{ type: "range", sheetName: "Data", address: "A1:D4" }] }],
        transactions: [{ transactionId: "txn_prepare", workbookId, agentId: "agent_prepare", status: "queued", operationCount: 1 }],
        conflicts: [{ conflictId: "conflict_prepare", workbookId, status: "open", reason: "overlapping range" }],
        events: [{ eventId: "event_prepare", workbookId, type: "task.created", createdAt: "2026-06-18T00:00:00.000Z" }]
      };
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({ request: "Prepare workbook", mode: "prepare" });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).sheetCount).toBe(sheets.length);
      expect((result.answer as any).tableCount).toBe(1);
      expect((result.answer as any).namedRangeCount).toBe(2);
      expect((result.answer as any).sheets.map((sheet: any) => sheet.name)).toEqual(sheets.map((sheet) => sheet.name));
      expect((result.answer as any).collaboration).toMatchObject({
        agentCount: 1,
        activeTaskCount: 1,
        activeLockCount: 1,
        activeTransactionCount: 1,
        openConflictCount: 1
      });
      expect((result.answer as any).collaboration.tasks[0]).toMatchObject({ taskId: "task_prepare", agentId: "agent_prepare", status: "running" });
      expect((result.answer as any).collaboration.locks[0]).toMatchObject({ lockId: "lock_data", ownerAgentId: "agent_prepare" });
      expect(runtime.runtimeMethodCalls["workbook.get_workbook_map"]).toBe(1);
      expect(runtime.runtimeMethodCalls["runtime.get_active_context"]).toBe(1);
      expect(runtime.runtimeMethodCalls["runtime.get_selection"]).toBe(1);
      expect(runtime.runtimeMethodCalls["table.list"]).toBe(1);
      expect(runtime.runtimeMethodCalls["table.get_info"]).toBe(1);
      expect(runtime.runtimeMethodCalls["names.list"]).toBe(1);
      expect(runtime.runtimeMethodCalls["region.list"]).toBe(1);
      expect(runtime.runtimeMethodCalls["runtime.get_collaboration_status"]).toBe(1);
      expect(runtime.readBatchCount).toBe(0);
    });

  it("returns a semantic workbook index from cached metadata without reading cell values", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Create workbook index",
        mode: "answer",
        detailLevel: "semantic_index",
        responseMode: "verbose",
        budget: { maxExamples: 20 }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("semantic_workbook_index");
      expect((result.answer as any).entries.map((entry: any) => entry.role)).toContain("transaction_sheet");
      expect(result.candidates?.some((candidate) => candidate.semanticRole === "data_table")).toBe(true);
      expect(result.resourceLinks.map((link) => link.uri)).toContain(`excel://agent/contexts/${result.workbookContextId}/semantic-index`);
      expect(result.telemetry.workflowRoute).toBe("semantic_index.find");
      expect(result.telemetry.metadataPolicy).toBe("sampled_allowed");
      expect(result.telemetry.readPolicy).toBe("metadata_only");
      expect(result.telemetry.fullReadUsed).toBe(false);
      expect(runtime.readBatchCount).toBe(0);
    });

  it("exposes semantic index through the workbook context resource", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const prepared = await agent.run({ request: "Prepare workbook", mode: "prepare" });
      const resource = agent.getSemanticIndexResource(String(prepared.workbookContextId)) as any;

      expect(resource.ok).toBe(true);
      expect(resource.semanticIndex.kind).toBe("semantic_workbook_index");
      expect(resource.semanticIndex.entries.some((entry: any) => entry.role === "data_table")).toBe(true);
    });

  it("answers vague workbook file reviews from complete structure metadata without sheet sampling", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
  
      const result = await agent.run({ request: "Can you look into transactions.xlsx?", responseMode: "verbose" });
  
      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("workbook_overview");
      expect((result.answer as any).semanticIndex.kind).toBe("semantic_workbook_index");
      expect((result.answer as any).sheets.map((sheet: any) => sheet.name)).toEqual(sheets.map((sheet) => sheet.name));
      expect(result.telemetry.workflowRoute).toBe("workbook.summary");
      expect(result.telemetry.metadataPolicy).toBe("structure_only");
      expect(result.telemetry.readPolicy).toBe("metadata_only");
      expect(result.telemetry.internalReadCount).toBe(0);
      expect(runtime.readBatchCount).toBe(0);
    });

  it("answers named-range inventory from cached workbook metadata", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({ request: "List named ranges in this workbook", mode: "answer" });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).kind).toBe("workbook_overview");
      expect((result.answer as any).namedRangeCount).toBe(2);
      expect((result.answer as any).namedRanges.map((name: any) => name.name)).toEqual(["RevenueTotal", "InputRegion"]);
      expect(result.telemetry.internalReadCount).toBe(0);
      expect(runtime.readBatchCount).toBe(0);
    });

  it("finds named ranges and registered regions as lookup candidates through bounded metadata sampling", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({ request: "Find the input region and revenue total", mode: "find", budget: { maxExamples: 10 } });

      expect(result.status).toBe("SUCCESS");
      expect(result.candidates?.map((candidate) => candidate.id)).toContain("name:InputRegion");
      expect(result.candidates?.map((candidate) => candidate.id)).toContain("name:RevenueTotal");
      expect(result.candidates?.find((candidate) => candidate.id === "name:InputRegion")?.range).toBe("B1:B3");
      expect(result.candidates?.find((candidate) => candidate.id === "name:RevenueTotal")?.sheetName).toBe("Report");
      expect(result.candidates?.some((candidate) => candidate.semanticRole === "named_region")).toBe(true);
      expect(result.telemetry.internalReadCount).toBe(0);
      expect(runtime.readBatchCount).toBe(sheets.length);
    });

  it("resolves a returned named-region candidate id for targeted live reads", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_named_region");
      metadata.namedRanges = [{ name: "InputRegion", sheetName: "Report", range: "B1:B3" }];
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "Read the selected named region",
        mode: "answer",
        workbookContextId: metadata.workbookContextId,
        target: { candidateId: "name:InputRegion" }
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.proof[0]).toMatchObject({ sheetName: "Report", range: "B1:B3" });
      expect((result.answer as any).kind).toBe("range_profile");
      expect(runtime.readBatchCount).toBe(1);
    });

  it("runs workbook validation through the agent validate mode", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({ request: "Validate workbook before saving", mode: "validate" });

      expect(result.status).toBe("SUCCESS");
      expect(result.summary).toBe("Workbook validation completed.");
      expect((result.answer as any).ok).toBe(true);
      expect(result.telemetry.validationStatus).toBe("passed");
      expect(runtime.runtimeMethodCalls["validate.workbook"]).toBe(1);
    });

  it("reports workbook validation failures without applying changes", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.validationResult = {
        ok: false,
        issues: [{ severity: "error", category: "formula", code: "FORMULA_ERROR", message: "Formula error found." }]
      };
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({ request: "Validate workbook before saving", mode: "validate" });

      expect(result.status).toBe("VALIDATION_FAILED");
      expect(result.nextAction).toBe("manual_review");
      expect((result.answer as any).issues[0].code).toBe("FORMULA_ERROR");
      expect(result.telemetry.validationStatus).toBe("failed");
      expect(runtime.writeBatchCount).toBe(0);
      expect(runtime.runtimeMethodCalls["validate.workbook"]).toBe(1);
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
      expect(result.telemetry.payloadBytes).toBeLessThanOrEqual(1200);
      expect(result.telemetry.estimatedTokensSaved).toBeGreaterThanOrEqual(0);
    });
});
