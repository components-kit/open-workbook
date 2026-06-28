import { describe, expect, it } from "vitest";
import type { AgentRunOutput } from "@components-kit/open-workbook-protocol";
import { agentJsonResult } from "./results.js";

describe("MCP result rendering", () => {
  it("marks workbook summaries as complete so clients do not chase full resources", () => {
    const output: AgentRunOutput = {
      status: "SUCCESS",
      mode: "answer",
      workbookContextId: "wbctx_summary",
      summary: "Returned workbook summary for RFQ.xlsx from cached metadata.",
      answer: {
        kind: "workbook_summary",
        source: "cached_metadata",
        workbook: { workbookId: "RFQ.xlsx", name: "RFQ.xlsx", sheetCount: 2 },
        sheetCount: 2,
        tableCount: 0,
        namedRangeCount: 0,
        sheets: [
          { name: "Vendor Propose", usedRange: "B2:K10" },
          { name: "Term condition", usedRange: "A3:F21" }
        ]
      },
      proof: [{ sheetName: "Vendor Propose", range: "B2:K10", label: "used range" }],
      resourceLinks: [],
      nextAction: "answer_now",
      taskOutcome: "final_answer",
      finalAnswer: "Workbook summary: 2 sheets, 0 tables.",
      agentInstruction: "This workbook summary is complete for an overview request. Answer now from cached metadata; do not fetch fullResultUri, chunk-read sheets, or call low-level MCP resources unless the user asks for all raw rows or exact cell values.",
      maxRecommendedFollowupCalls: 0,
      warnings: [],
      telemetry: {
        internalCallCount: 1,
        payloadBytes: 900,
        estimatedTokens: 225,
        elapsedMs: 2,
        cacheHit: true
      }
    };

    const result = agentJsonResult(output);
    const text = result.content[0]?.text ?? "";

    expect(text).toContain("taskOutcome: final_answer");
    expect(text).toContain("maxRecommendedFollowupCalls: 0");
    expect(text).toContain("summary is complete from cached metadata");
    expect(text).toContain("do not fetch fullResultUri");
    expect(text).not.toContain("excel.agent.run continuation.fullResultUri");
    expect((result.structuredContent.answer as any).kind).toBe("workbook_summary");
  });

  it("marks data validation summaries as complete inline for dropdown inspection", () => {
    const output: AgentRunOutput = {
      status: "SUCCESS",
      mode: "answer",
      workbookContextId: "wbctx_1",
      summary: "Read data validation for May 2026!E:E.",
      answer: {
        kind: "data_validation_summary",
        source: "runtime_range_metadata",
        method: "range.read_data_validation",
        sheetName: "May 2026",
        range: "E:E",
        ruleCount: 1,
        type: "list",
        inCellDropDown: true,
        options: ["driver_wage_remaining", "owner_cash_topup"],
        optionCount: 2,
        sourceComplete: true,
        fieldContext: [{
          field: "Transaction Type",
          range: "E:E",
          hasValidation: true,
          allowedValues: ["driver_wage_remaining", "owner_cash_topup"],
          validation: { type: "list", sourceType: "inline", optionCount: 2 }
        }],
        guidance: "Use this inline validation summary to answer dropdown option questions."
      },
      proof: [{ sheetName: "May 2026", range: "E:E", label: "range metadata" }],
      resourceLinks: [],
      nextAction: "answer_now",
      taskOutcome: "final_answer",
      finalAnswer: "Read dropdown validation for May 2026!E:E.",
      agentInstruction: "Answer from this data_validation_summary. Dropdown validation metadata/options are complete inline for the requested range; do not fetch fullResultUri.",
      maxRecommendedFollowupCalls: 0,
      warnings: [],
      telemetry: {
        internalCallCount: 1,
        payloadBytes: 1000,
        estimatedTokens: 250,
        elapsedMs: 2,
        cacheHit: true
      }
    };

    const result = agentJsonResult(output);
    const text = result.content[0]?.text ?? "";

    expect(text).toContain("data: validation metadata/options complete inline");
    expect(text).toContain("do not fetch full detail");
    expect(text).not.toContain("exact rows/raw values need");
    expect((result.structuredContent.answer as any).kind).toBe("data_validation_summary");
    expect((result.structuredContent.answer as any).fieldContext[0]).toMatchObject({
      field: "Transaction Type",
      validation: { sourceType: "inline" }
    });
  });

  it("keeps text compact while preserving structured content and resource links", () => {
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
    expect(text).toContain("data: compact summary inline; exact rows/raw values need excel.agent.run");
    expect(text).toContain("resultUri: excel://agent/results/agentres_1");
    expect(text).toContain("excel.agent.run continuation.fullResultUri");
    expect(text).toContain("not web");
    expect(text).toContain("never Webfetch/browser");
    expect(text).not.toContain("sparseRows");
    expect((result.structuredContent.answer as any).sparseRows).toBeUndefined();
    expect((result.structuredContent.answer as any).fullResultUri).toBe("excel://agent/results/agentres_1?view=full");
    expect((result.structuredContent.telemetry as any).routeReasons).toBeUndefined();
    expect((result.structuredContent.telemetry as any).workflowReasons).toBeUndefined();
    expect((result.structuredContent.telemetry as any).semanticIndexStatus).toBeUndefined();
    expect(result.structuredContent.resourceLinks).toEqual([
      {
        uri: "excel://agent/results/agentres_1",
        name: "agent result",
        description: "Stored agent answer detail.",
        mimeType: "application/json"
      }
    ]);
  });

  it("preserves similar-row evidence when MCP brief budget compacts the result", () => {
    const output: AgentRunOutput = {
      status: "SUCCESS",
      mode: "answer",
      workbookContextId: "wbctx_1",
      summary: "Found 5 similar historical row(s) across related workbook sheets.",
      answer: {
        kind: "similar_rows",
        source: { sheetName: "May 2026", range: "A20:K20" },
        sourceMode: "exact_source_row",
        predicates: [{ label: "Direction is Inflow", value: "inflow" }, { label: "Amount equals 10000", value: 10000 }],
        comparedRanges: Array.from({ length: 20 }, (_value, index) => ({ sheetName: `Sheet ${index}`, range: "A1:K244", reason: "related sheet" })),
        rows: Array.from({ length: 8 }, (_value, index) => ({
          sheetName: "Apr 2026",
          sheetRowNumber: index + 2,
          range: `A${index + 2}:K${index + 2}`,
          columns: [
            { letter: "A", name: "Transaction Date", value: "2026-04-16" },
            { letter: "D", name: "Description", value: "Owner adding fund" },
            { letter: "E", name: "Transaction Type", value: "owner_fund_added" },
            { letter: "F", name: "Direction", value: "Inflow" },
            { letter: "G", name: "Cash Amount", value: 10000 },
            { letter: "K", name: "Transfer From / To", value: "From X1183 MR. PRACH YOTHAPRA++" }
          ],
          matchedSignals: ["Direction is Inflow", "Amount equals 10000", "Transfer contains PRACH"],
          matchedColumns: [],
          whyMatched: "related sheet; Direction is Inflow; Amount equals 10000",
          filler: "x".repeat(1200)
        }))
      },
      proof: [{ sheetName: "Apr 2026", range: "A2:K2", label: "similar row" }],
      resourceLinks: [{
        uri: "excel://agent/results/agentres_1",
        name: "agent result",
        description: "Stored agent answer detail.",
        mimeType: "application/json"
      }],
      continuation: {
        workbookContextId: "wbctx_1",
        resultUri: "excel://agent/results/agentres_1",
        fullResultUri: "excel://agent/results/agentres_1?view=full",
        responseMode: "brief"
      },
      nextAction: "answer_now",
      taskOutcome: "final_answer",
      finalAnswer: "Found similar rows.",
      agentInstruction: "Answer the user now from finalAnswer, proof, and inline structuredContent; do not call workbook tools again for this task.",
      maxRecommendedFollowupCalls: 0,
      warnings: ["Agent response was compacted to satisfy the requested payload/token budget."],
      telemetry: {
        internalCallCount: 1,
        payloadBytes: 9000,
        estimatedTokens: 2250,
        elapsedMs: 2,
        cacheHit: false
      }
    };

    const result = agentJsonResult(output);
    const answer = result.structuredContent.answer as any;

    expect(answer.kind).toBe("similar_rows");
    expect(answer.rows).toHaveLength(5);
    expect(answer.rows[0].columns).toEqual(expect.arrayContaining([
      expect.objectContaining({ name: "Transaction Type", value: "owner_fund_added" }),
      expect.objectContaining({ name: "Transfer From / To", value: "From X1183 MR. PRACH YOTHAPRA++" })
    ]));
    expect(JSON.stringify(answer)).not.toContain("filler");
    expect(result.content[0]?.text).toContain("data: 5 exact rows inline");
  });

  it("preserves formula proof in brief structured content", () => {
    const output: AgentRunOutput = {
      status: "SUCCESS",
      mode: "answer",
      workbookContextId: "wbctx_1",
      summary: "Read formula proof for Report!B2.",
      answer: {
        kind: "formula_read",
        sheetName: "Report",
        range: "B2:B2",
        formulaCount: 1,
        hardcodedCount: 0,
        blankCount: 0,
        values: [[600]],
        text: [["600"]],
        formulas: [["=SUM(Data!C2:C4)"]],
        numberFormat: [["#,##0"]],
        cells: [{
          cell: "B2",
          value: 600,
          text: "600",
          formula: "=SUM(Data!C2:C4)",
          formulaR1C1: "=SUM(Data!R2C3:R4C3)",
          formulaStatus: "formula"
        }]
      },
      proof: [{ sheetName: "Report", range: "B2:B2" }],
      resourceLinks: [],
      continuation: {
        workbookContextId: "wbctx_1",
        responseMode: "brief"
      },
      nextAction: "answer_now",
      taskOutcome: "final_answer",
      finalAnswer: "Report!B2 contains =SUM(Data!C2:C4).",
      agentInstruction: "Answer the user now from formula proof; do not infer formulas from displayed values alone.",
      maxRecommendedFollowupCalls: 0,
      warnings: [],
      telemetry: {
        internalCallCount: 1,
        payloadBytes: 1800,
        estimatedTokens: 450,
        elapsedMs: 2,
        cacheHit: false
      }
    };

    const result = agentJsonResult(output);
    const answer = result.structuredContent.answer as any;

    expect(answer.kind).toBe("formula_read");
    expect(answer.formulas).toEqual([["=SUM(Data!C2:C4)"]]);
    expect(answer.text).toEqual([["600"]]);
    expect(answer.numberFormat).toEqual([["#,##0"]]);
    expect(answer.cells[0]).toMatchObject({
      cell: "B2",
      formula: "=SUM(Data!C2:C4)",
      formulaStatus: "formula"
    });
    expect(result.content[0]?.text).toContain("taskOutcome: final_answer");
  });
});
