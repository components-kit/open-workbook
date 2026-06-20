import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { AGENT_DETAIL_LEVELS, AGENT_INTENT_ACTIONS } from "@components-kit/open-workbook-protocol";

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
});
