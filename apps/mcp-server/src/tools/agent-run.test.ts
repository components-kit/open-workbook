import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { AGENT_DETAIL_LEVELS, AGENT_INTENT_ACTIONS } from "@components-kit/open-workbook-protocol";
import { agentRunInputSchema } from "./agent-run.js";

describe("excel.agent.run MCP schema", () => {
  it("accepts operation lifecycle statuses returned by the backend", () => {
    const source = readFileSync(new URL("./agent-run.ts", import.meta.url), "utf8");

    expect(source).toContain("AGENT_RUN_STATUSES");
    expect(source).toContain("AGENT_RUN_MODES");
    expect(source).toContain("AGENT_DETAIL_LEVELS");
  });

  it("mentions every public agent intent action through the protocol import", () => {
    expect(AGENT_INTENT_ACTIONS).toContain("replace_range_with_styled_table");
    expect(AGENT_INTENT_ACTIONS).toContain("read_style_summary");
    expect(AGENT_INTENT_ACTIONS).toContain("format_diagnostics");
    expect(AGENT_DETAIL_LEVELS).toContain("full_table");
    expect(AGENT_DETAIL_LEVELS).toContain("semantic_index");
  });

  it("exposes semantic and workflow telemetry fields in the output schema", () => {
    const source = readFileSync(new URL("./agent-run.ts", import.meta.url), "utf8");

    expect(source).toContain("workflowRoute");
    expect(source).toContain("semanticIndexStatus");
    expect(source).toContain("semanticCandidateUsed");
    expect(source).toContain("metadataPolicy");
    expect(source).toContain("readPolicy");
  });

  it("allows cache invalidation fields returned after successful applies", () => {
    const source = readFileSync(new URL("./agent-run.ts", import.meta.url), "utf8");

    expect(source).toContain("invalidatedContextIds");
    expect(source).toContain("invalidatedResourceUris");
  });

  it("normalizes JSON-string structured fields from lenient MCP clients", () => {
    const schema = agentRunInputSchema();

    expect((schema.target as any).parse("{\"sheetName\":\"Booking\",\"range\":\"A1:X7\"}")).toEqual({ sheetName: "Booking", range: "A1:X7" });
    expect((schema.intent as any).parse("{\"action\":\"read_values\",\"targetHints\":[\"Booking\"]}")).toEqual({ action: "read_values", targetHints: ["Booking"] });
    expect((schema.continuation as any).parse("{\"workbookContextId\":\"wbctx_1\",\"fullResultUri\":\"excel://agent/results/agentres_1?view=full\"}")).toEqual({
      workbookContextId: "wbctx_1",
      fullResultUri: "excel://agent/results/agentres_1?view=full"
    });
    expect((schema.values as any).parse({
      patches: [
        {
          target: "{\"sheetName\":\"Booking\",\"range\":\"A1:B2\"}",
          values: [[1, 2]]
        }
      ]
    })).toEqual({
      patches: [
        {
          target: { sheetName: "Booking", range: "A1:B2" },
          values: [[1, 2]]
        }
      ]
    });
  });

  it("accepts common structured update payloads without requiring agents to send schemas", () => {
    const schema = agentRunInputSchema();

    expect((schema.values as any).parse({
      values: [["Reviewed"]],
      style: { fillColor: "#1F4E78", fontColor: "#FFFFFF", fontBold: true },
      options: ["Open", "Reviewed", "Closed"],
      validation: { type: "list", source: ["Open", "Reviewed", "Closed"], inCellDropDown: true },
      rule: { type: "custom", formula: "=$E2=\"Open\"", style: { fillColor: "#FFFF00" } },
      columnOrder: [2, 1],
      numberFormat: "dd/mm/yyyy"
    })).toMatchObject({
      values: [["Reviewed"]],
      style: { fillColor: "#1F4E78", fontColor: "#FFFFFF", fontBold: true },
      options: ["Open", "Reviewed", "Closed"],
      validation: { type: "list", source: ["Open", "Reviewed", "Closed"], inCellDropDown: true },
      rule: { type: "custom", formula: "=$E2=\"Open\"", style: { fillColor: "#FFFF00" } },
      columnOrder: [2, 1],
      numberFormat: "dd/mm/yyyy"
    });
    expect((schema.values as any).parse({ numberFormats: [["dd/mm/yyyy"]] })).toMatchObject({ numberFormats: [["dd/mm/yyyy"]] });
    expect((schema.values as any).parse({
      rows: [
        { index: 1, values: ["2026-01-04", "Northwind", "Support", 525, "Closed"] }
      ]
    })).toMatchObject({
      rows: [
        { index: 1, values: ["2026-01-04", "Northwind", "Support", 525, "Closed"] }
      ]
    });
  });

  it("rejects malformed common update payloads with field-level schema errors", () => {
    const schema = agentRunInputSchema();

    expect(() => (schema.values as any).parse({ style: { fillColor: 42 } })).toThrow(/fillColor/i);
    expect(() => (schema.values as any).parse({ options: ["Open", 7] })).toThrow(/options/i);
    expect(() => (schema.values as any).parse({ validation: { source: ["Open", 7] } })).toThrow(/source/i);
    expect(() => (schema.values as any).parse({ rule: { formula: 42, style: { fillColor: "#FFFF00" } } })).toThrow(/formula/i);
    expect(() => (schema.values as any).parse({ columnOrder: [2, { bad: true }] })).toThrow(/columnOrder/i);
    expect(() => (schema.values as any).parse({ numberFormat: [["dd/mm/yyyy", 42]] })).toThrow(/numberFormat/i);
  });
});
