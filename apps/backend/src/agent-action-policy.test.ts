import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { AGENT_ACTION_REGISTRY, assessAgentUpdateRisk, riskForOperationKind } from "./agent-action-policy.js";

const DYNAMIC_POLICY_KINDS = new Set(["batch", "style.copy_dimensions_many", "workflow.replace_styled_table"]);

describe("agent action policy contracts", () => {
  it("covers every pending agent action with apply routing and risk policy", () => {
    const storeSource = readFileSync(new URL("./agent-operation-store.ts", import.meta.url), "utf8");
    const orchestratorSource = readFileSync(new URL("./agent-orchestrator.ts", import.meta.url), "utf8");
    const pendingKinds = pendingAgentActionKinds(storeSource);
    const applyCases = applyPendingActionCases(orchestratorSource);
    const policyKinds = new Set(AGENT_ACTION_REGISTRY.map((definition) => definition.kind));

    expect([...pendingKinds].filter((kind) => !applyCases.has(kind)).sort()).toEqual([]);
    expect([...pendingKinds].filter((kind) => !DYNAMIC_POLICY_KINDS.has(kind) && !policyKinds.has(kind)).sort()).toEqual([]);
    expect(riskForOperationKind("style.copy_dimensions_many")).toBe("safe_format");
    expect(riskForOperationKind("workflow.replace_styled_table")).toBe("destructive");
  });

  it("assesses cache and safety risk separately for common mutations", () => {
    expect(assessAgentUpdateRisk({
      kind: "batch",
      operations: [{
        kind: "range.write_values_many",
        operationId: "op",
        workbookId: "wb" as any,
        destructiveLevel: "values",
        reason: "unit",
        entries: [{ target: { workbookId: "wb" as any, sheetName: "Data", address: "D2" }, values: [["Closed"]] }]
      }]
    })).toMatchObject({
      cacheRisk: "low",
      safetyRisk: "low",
      cacheAction: "update_cache",
      invalidatedFacets: expect.arrayContaining(["values", "aggregates", "formulaResults"]),
      preservedFacets: expect.arrayContaining(["schema", "headers", "fieldContext", "validation"]),
      requiresRefreshBeforeNextMutation: false
    });

    expect(assessAgentUpdateRisk({
      kind: "table.append_rows",
      request: { workbookId: "wb" as any, tableName: "Transactions", rows: [[1, 2, 3]] }
    })).toMatchObject({
      cacheRisk: "medium",
      safetyRisk: "medium",
      cacheAction: "partial_invalidate",
      invalidatedFacets: expect.arrayContaining(["tableDimensions", "aggregates", "rowPositions"])
    });

    expect(assessAgentUpdateRisk({
      kind: "table.reorder_columns",
      request: { workbookId: "wb" as any, tableName: "Transactions", columns: ["Status", "Date"] }
    })).toMatchObject({
      cacheRisk: "high",
      cacheAction: "rebuild_context",
      requiresRefreshBeforeNextMutation: true
    });
  });
});

function pendingAgentActionKinds(source: string): Set<string> {
  const unionSource = source.slice(source.indexOf("export type PendingAgentAction"), source.indexOf("export type AgentCleanMutationAction"));
  return new Set([...unionSource.matchAll(/kind:\s+"([^"]+)"/g)].map((match) => match[1]!));
}

function applyPendingActionCases(source: string): Set<string> {
  const methodSource = source.slice(source.indexOf("private applyPendingActionInContext"), source.indexOf("private async applyStyleCopyRequests"));
  return new Set([...methodSource.matchAll(/case\s+"([^"]+)"/g)].map((match) => match[1]!));
}
