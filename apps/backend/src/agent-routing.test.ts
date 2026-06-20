import { describe, expect, it } from "vitest";
import { routeAgentRequest } from "./agent-routing.js";

describe("agent workflow routing", () => {
  it.each([
    ["What is this workbook?", "auto", undefined, "workbook.summary", "structure_only", "metadata_only"],
    ["Where is the receipt template?", "auto", undefined, "semantic_index.find", "sampled_allowed", "metadata_only"],
    ["Look at this sheet", "auto", undefined, "sheet.summary", "structure_only", "metadata_only"],
    ["What is the current styling?", "auto", undefined, "style.inspect", "sampled_allowed", "targeted_read"],
    ["Why is this date format wrong?", "auto", undefined, "format.diagnostics", "sampled_allowed", "targeted_read"],
    ["Read Transactions table sample", "auto", undefined, "table.sample", "sampled_required", "targeted_read"],
    ["Update invoice status", "auto", undefined, "mutation.preview", "sampled_allowed", "preview_only"],
    ["Apply update", "apply_update", undefined, "mutation.apply", "structure_only", "apply_only"],
    ["Validate workbook", "validate", undefined, "validation.run", "sampled_allowed", "targeted_read"]
  ] as const)("routes %s to %s", (request, mode, intent, workflowRoute, metadataPolicy, readPolicy) => {
    const route = routeAgentRequest(request, mode, intent);

    expect(route.workflowRoute).toBe(workflowRoute);
    expect(route.metadataPolicy).toBe(metadataPolicy);
    expect(route.readPolicy).toBe(readPolicy);
    expect(route.workflowConfidence).toBeGreaterThan(0);
  });

  it.each([
    ["read_style_summary", "style.inspect", "targeted_read"],
    ["format_diagnostics", "format.diagnostics", "targeted_read"],
    ["find_target", "semantic_index.find", "metadata_only"],
    ["write_values", "mutation.preview", "preview_only"]
  ] as const)("prefers structured intent action %s for workflow routing", (action, workflowRoute, readPolicy) => {
    const route = routeAgentRequest("Thai or mixed language request", "auto", {
      source: "caller_structured",
      action,
      accepted: true,
      confidence: 0.91
    });

    expect(route.workflowRoute).toBe(workflowRoute);
    expect(route.readPolicy).toBe(readPolicy);
    expect(route.workflowConfidence).toBe(0.91);
  });
});
