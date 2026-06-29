import { describe, expect, it } from "vitest";
import { AgentOrchestrator } from "./agent-orchestrator.js";
import { FakeAgentRuntime, createCachedMetadata } from "./agent-orchestrator.test-support.js";

describe("agent context and query regression evals", () => {
  it("keeps overview answers cache-first and self-contained", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const result = await agent.run({
      request: "What is this workbook about?",
      mode: "answer",
      detailLevel: "workbook_summary"
    });

    expect(result.status).toBe("SUCCESS");
    expect(result.nextAction).toBe("answer_now");
    expect(result.maxRecommendedFollowupCalls).toBe(0);
    expect(result.contextUsed).toMatchObject({
      strategy: "overview",
      levelUsed: 2,
      stagesUsed: expect.arrayContaining(["metadata", "schema"]),
      included: expect.arrayContaining(["metadata", "schema"])
    });
    expect(result.telemetry.internalReadCount).toBe(0);
    expect(runtime.readBatchCount).toBe(0);
  });

  it("reports stale-only context refresh plans without invalidating the whole context", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);
    const metadata = createCachedMetadata("wbctx_eval_refresh");
    agent.metadataCache.set(metadata);
    agent.metadataCache.markFacetsStale(metadata.workbookContextId, ["values"], ["Data!B2"]);

    const result = await agent.run({
      request: "Analyze Data values",
      mode: "answer",
      workbookContextId: metadata.workbookContextId,
      target: { sheetName: "Data" },
      context: { strategy: "analysis", include: ["schema", "values"] }
    });

    expect(result.contextUsed).toMatchObject({
      cachedFacetsUsed: expect.arrayContaining(["schema", "headers"]),
      staleFacets: ["values"],
      facetsToRefresh: ["values"],
      freshnessRequiresRead: true
    });
    expect(agent.metadataCache.getByContextId(metadata.workbookContextId)).toBeDefined();
  });

  it("updates cache optimistically after exact value patches", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);
    const metadata = createCachedMetadata("wbctx_eval_optimistic");
    agent.metadataCache.set(metadata);

    const preview = await agent.run({
      request: "Update Data B2",
      mode: "preview_update",
      workbookContextId: metadata.workbookContextId,
      target: { sheetName: "Data", range: "B2" },
      values: { patches: [{ target: { sheetName: "Data", range: "B2" }, values: [[999]] }] }
    });
    const applied = await agent.run({
      request: "Apply update",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(applied.status).toBe("SUCCESS");
    expect(applied.invalidatedContextIds).not.toContain(metadata.workbookContextId);
    expect(agent.metadataCache.getOptimisticValue(metadata.workbookContextId, "Data", "B2")).toMatchObject({ value: 999 });
    expect(agent.metadataCache.getContextState(metadata.workbookContextId)?.freshness.staleFacets).not.toContain("values");
  });

  it("uses query_rows for lookup and returns patch-only suggested operations", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);
    const metadata = createCachedMetadata("wbctx_eval_query_rows");
    agent.metadataCache.set(metadata);

    const result = await agent.run({
      request: "Find open rows and prepare to mark them reviewed",
      mode: "answer",
      workbookContextId: metadata.workbookContextId,
      intent: { action: "query_rows" },
      target: { sheetName: "Data", tableName: "Transactions" },
      values: {
        where: [{ column: "Status", op: "=", value: "Open" }],
        return: ["Date", "Status"],
        updateColumn: "Status",
        updateValue: "Reviewed",
        limit: 10
      },
      responseMode: "verbose"
    });

    expect(result.status).toBe("SUCCESS");
    expect((result.answer as any)).toMatchObject({
      kind: "query_rows_result",
      matchedRows: 2,
      rowAddresses: ["Data!2:2", "Data!4:4"]
    });
    expect(result.suggestedOperation).toMatchObject({
      mode: "preview_update",
      intent: { action: "write_values" },
      values: {
        patches: [
          { target: { sheetName: "Data", range: "D2" }, values: [["Reviewed"]] },
          { target: { sheetName: "Data", range: "D4" }, values: [["Reviewed"]] }
        ]
      }
    });
    expect(runtime.writeBatchCount).toBe(0);
  });
});
