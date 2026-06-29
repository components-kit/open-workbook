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
    expect(route.contextDecision.strategy).toBeTruthy();
    expect(route.contextDecision.include.length).toBeGreaterThan(0);
  });

  it.each([
    ["read_style_summary", "style.inspect", "targeted_read"],
    ["format_diagnostics", "format.diagnostics", "targeted_read"],
    ["find_target", "semantic_index.find", "metadata_only"],
    ["query_rows", "rows.query", "targeted_read"],
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

  it("routes lookup-style row requests to read-only query workflow", () => {
    const route = routeAgentRequest("Show rows where Status = Unpaid", "auto");

    expect(route.mode).toBe("answer");
    expect(route.workflowRoute).toBe("rows.query");
    expect(route.readPolicy).toBe("targeted_read");
    expect(route.contextDecision).toMatchObject({
      strategy: "focused",
      level: 3
    });
    expect(route.contextDecision.include).toEqual(expect.arrayContaining(["schema", "field_context", "values"]));
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

  it.each([
    ["What is this sheet about?", "answer", "overview", "active_sheet", 2, ["metadata", "schema"]],
    ["Look here and explain this area", "answer", "focused", "active_selection", 4, ["values", "field_context", "validation"]],
    ["Analyze this sheet for trends", "answer", "analysis", "workbook", 5, ["values", "formulas"]],
    ["Check why this dropdown is broken", "answer", "audit", "workbook", 4, ["validation", "filters", "formulas"]],
    ["Find payment status column", "find", "overview", "workbook", 2, ["schema", "tables", "regions"]],
    ["Update this category", "preview_update", "focused", "workbook", 3, ["field_context", "validation"]]
  ] as const)("infers context policy for %s", (request, mode, strategy, scope, level, include) => {
    const route = routeAgentRequest(request, mode);

    expect(route.contextDecision).toMatchObject({ strategy, scope, level, source: "inferred" });
    expect(route.contextDecision.include).toEqual(expect.arrayContaining([...include]));
    expect(route.contextDecision.plannedStages.length).toBeGreaterThan(0);
    expect(route.contextDecision.stopWhen).toBeTruthy();
  });

  it("uses explicit target as focused context scope for mutation previews", () => {
    const route = routeAgentRequest("Update status", "preview_update", undefined, undefined, { sheetName: "Data", range: "B2" });

    expect(route.contextDecision).toMatchObject({
      strategy: "focused",
      scope: "target"
    });
    expect(route.contextDecision.include).toEqual(expect.arrayContaining(["field_context", "validation"]));
  });

  it("respects caller context policy and fills missing fields from defaults", () => {
    const route = routeAgentRequest("Check dropdown issue", "answer", undefined, {
      strategy: "audit",
      level: 4,
      include: ["validation"]
    });

    expect(route.contextDecision).toMatchObject({
      strategy: "audit",
      scope: "workbook",
      level: 4,
      include: ["validation"],
      source: "caller"
    });
  });

  it("lets callers cap or request an explicit context level", () => {
    const route = routeAgentRequest("Analyze this sheet for trends", "answer", undefined, {
      level: 2,
      strategy: "analysis"
    });

    expect(route.contextDecision).toMatchObject({
      strategy: "analysis",
      level: 2,
      plannedStages: ["metadata", "schema", "semantic_index"],
      source: "caller"
    });
  });
});
