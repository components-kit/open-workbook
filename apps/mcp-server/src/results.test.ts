import { describe, expect, it } from "vitest";
import type { AgentRunOutput } from "@components-kit/open-workbook-protocol";
import { agentJsonResult } from "./results.js";

describe("MCP result rendering", () => {
  it("keeps text compact while preserving structured content and resources", () => {
    const output: AgentRunOutput = {
      status: "SUCCESS",
      mode: "answer",
      workbookContextId: "wbctx_1",
      summary: "Answered range profile from cached metadata.",
      answer: {
        kind: "range_profile",
        sparseRows: Array.from({ length: 20 }, (_value, index) => ({
          row: index + 1,
          cells: [{ address: `A${index + 1}`, value: `value ${index + 1}` }]
        }))
      },
      proof: [{ sheetName: "Sheet1", range: "A1:B10" }],
      resourceLinks: [
        {
          uri: "excel://agent/results/agentres_1",
          name: "agent result",
          description: "Stored agent answer detail.",
          mimeType: "application/json"
        }
      ],
      continuation: {
        workbookContextId: "wbctx_1",
        resultUri: "excel://agent/results/agentres_1",
        fullResultUri: "excel://agent/results/agentres_1?view=full",
        responseMode: "brief"
      },
      nextAction: "answer_now",
      taskOutcome: "final_answer",
      finalAnswer: "Answered range profile from cached metadata.",
      agentInstruction: "Answer the user now from finalAnswer, proof, and inline structuredContent; do not call workbook tools again for this task.",
      maxRecommendedFollowupCalls: 0,
      warnings: [],
      telemetry: {
        internalCallCount: 1,
        payloadBytes: 1234,
        estimatedTokens: 309,
        elapsedMs: 2,
        cacheHit: true,
        routeReasons: ["debug reason that should not be in brief MCP structuredContent"],
        workflowReasons: ["workflow debug reason that should not be in brief MCP structuredContent"],
        semanticIndexStatus: "built"
      }
    };

    const result = agentJsonResult(output);
    const text = result.content[0]?.text ?? "";

    expect(text.length).toBeLessThan(700);
    expect(text).toContain("SUCCESS answer");
    expect(text).toContain("taskOutcome: final_answer");
    expect(text).toContain("maxRecommendedFollowupCalls: 0");
    expect(text).toContain("do not call workbook tools again");
    expect(text).toContain("data: compact summary inline; fullResultUri available");
    expect(text).toContain("resultUri: excel://agent/results/agentres_1");
    expect(text).toContain("call excel.agent.run with fullResultUri");
    expect(text).toContain("do not use webfetch");
    expect(text).not.toContain("sparseRows");
    expect((result.structuredContent.answer as any).sparseRows).toBeUndefined();
    expect((result.structuredContent.answer as any).fullResultUri).toBe("excel://agent/results/agentres_1?view=full");
    expect((result.structuredContent.telemetry as any).routeReasons).toBeUndefined();
    expect((result.structuredContent.telemetry as any).workflowReasons).toBeUndefined();
    expect((result.structuredContent.telemetry as any).semanticIndexStatus).toBeUndefined();
    expect(result.resources).toEqual([
      {
        uri: "excel://agent/results/agentres_1",
        name: "agent result",
        description: "Stored agent answer detail.",
        mimeType: "application/json"
      }
    ]);
  });
});
