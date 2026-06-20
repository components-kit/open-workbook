import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

describe("MCP resources", () => {
  it("registers the agent semantic workbook index resource", () => {
    const source = readFileSync(new URL("./resources.ts", import.meta.url), "utf8");

    expect(source).toContain("excel://agent/contexts/{workbook_context_id}/semantic-index");
    expect(source).toContain("getAgentSemanticIndexResource");
    expect(source).toContain("excel://agent/results/{result_id}");
    expect(source).toContain("excel://compact/{resource_id}");
    expect(source).toContain("getCompactResource");
  });
});
