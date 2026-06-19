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
      warnings: [],
      telemetry: {
        internalCallCount: 1,
        payloadBytes: 1234,
        estimatedTokens: 309,
        elapsedMs: 2,
        cacheHit: true
      }
    };

    const result = agentJsonResult(output);
    const text = result.content[0]?.text ?? "";

    expect(text.length).toBeLessThan(500);
    expect(text).toContain("SUCCESS answer");
    expect(text).toContain("resultUri: excel://agent/results/agentres_1");
    expect(text).not.toContain("sparseRows");
    expect(result.structuredContent.answer).toEqual(output.answer);
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
