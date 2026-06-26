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
    ["write_values", "mutation.preview", "preview_only"],
    ["improve_visual_readability", "mutation.preview", "preview_only"]
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

  it.each([
    "Okay, what about the styling?",
    "Read the style summary of the Booking sheet",
    "Inspect current fonts, colors, borders, alignment, fills, and number formats"
  ])("routes style inspection wording to answer mode: %s", (request) => {
    const route = routeAgentRequest(request, "auto");

    expect(route.mode).toBe("answer");
    expect(route.matchedRule).toBe("read_inspection.keyword");
    expect(route.workflowRoute).toBe("style.inspect");
    expect(route.readPolicy).toBe("targeted_read");
  });

  it.each([
    "Read columns P, T, U, V from Booking sheet rows 2-7 to check if dates are now proper date serials with 4-digit year format.",
    "Verify the displayed date values and formats in P2:V7",
    "Check cells P2:V7 and confirm the date serials"
  ])("routes read/verify wording with format terms to answer mode: %s", (request) => {
    const route = routeAgentRequest(request, "auto");

    expect(route.mode).toBe("answer");
    expect(route.matchedRule).toBe("read_inspection.keyword");
    expect(route.workflowRoute).toBe("range.read");
    expect(route.readPolicy).toBe("targeted_read");
  });

  it.each([
    "Change the styling of A1:X7",
    "Set the fill color on Booking",
    "Format A1:X7 with borders"
  ])("keeps explicit style writes on preview mode: %s", (request) => {
    const route = routeAgentRequest(request, "auto");

    expect(route.mode).toBe("preview_update");
    expect(route.workflowRoute).toBe("mutation.preview");
  });
});
