import { describe, expect, it } from "vitest";
import { getInternalCapabilityCatalog, getInternalCapabilityCatalogSummary, getPublicAgentToolCatalog, InternalCapabilityCatalog, PublicAgentToolCatalog } from "./tools.js";
import { PromptCatalog } from "./prompts.js";
import { ResourceCatalog } from "./resources.js";

describe("public tool and internal capability catalogs", () => {
  it("contains the full internal backend capability inventory", () => {
    expect(InternalCapabilityCatalog.length).toBeGreaterThan(200);
    expect(InternalCapabilityCatalog.some((capability) => capability.name === "excel.agent.run")).toBe(true);
    expect(InternalCapabilityCatalog.some((capability) => capability.name === "excel.runtime.get_capabilities")).toBe(true);
    expect(InternalCapabilityCatalog.some((capability) => capability.name === "excel.workbook.get_workbook_map")).toBe(true);
    expect(InternalCapabilityCatalog.some((capability) => capability.name === "excel.workbook.get_summary")).toBe(true);
    expect(InternalCapabilityCatalog.some((capability) => capability.name === "excel.range.read_compact")).toBe(true);
    expect(InternalCapabilityCatalog.some((capability) => capability.name === "excel.table.read_compact")).toBe(true);
    expect(InternalCapabilityCatalog.some((capability) => capability.name === "excel.lookup.search_workbook")).toBe(true);
    expect(InternalCapabilityCatalog.some((capability) => capability.name === "excel.lookup.inspect_match")).toBe(true);
    expect(InternalCapabilityCatalog.some((capability) => capability.name === "excel.validate.compact")).toBe(true);
    expect(InternalCapabilityCatalog.some((capability) => capability.name === "excel.snapshot.get_compact")).toBe(true);
    expect(InternalCapabilityCatalog.some((capability) => capability.name === "excel.workflow.prepare_session")).toBe(true);
    expect(InternalCapabilityCatalog.some((capability) => capability.name === "excel.workflow.create_formula_sheet")).toBe(true);
    expect(InternalCapabilityCatalog.some((capability) => capability.name === "excel.workflow.create_template_report")).toBe(true);
    expect(InternalCapabilityCatalog.some((capability) => capability.name === "excel.workflow.create_pivot_chart_summary")).toBe(true);
    expect(InternalCapabilityCatalog.some((capability) => capability.name === "excel.workflow.repair_formula_errors")).toBe(true);
    expect(InternalCapabilityCatalog.some((capability) => capability.name === "excel.workflow.preview_risky_edit")).toBe(true);
    expect(InternalCapabilityCatalog.some((capability) => capability.name === "excel.workflow.inspect_analyze")).toBe(true);
    expect(InternalCapabilityCatalog.some((capability) => capability.name === "excel.workflow.rollback_validate")).toBe(true);
  });

  it("exposes only the public agent tool through the MCP catalog", () => {
    const exposed = getPublicAgentToolCatalog();
    expect(exposed.map((tool) => tool.name)).toEqual(["excel.agent.run"]);
    expect(exposed.every((tool) => tool.status === "stable")).toBe(true);
    expect(PublicAgentToolCatalog.map((tool) => tool.name)).toEqual(["excel.agent.run"]);
  });

  it("keeps compact and workflow capabilities in the internal backend catalog", () => {
    const exposed = getInternalCapabilityCatalog();
    const exposedNames = new Set(exposed.map((tool) => tool.name));
    expect(InternalCapabilityCatalog.length).toBeGreaterThan(PublicAgentToolCatalog.length);
    expect(InternalCapabilityCatalog.some((capability) => capability.name === "excel.range.write_values")).toBe(true);
    expect(exposed.every((tool) => tool.status === "stable")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.runtime.get_status")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.table.reorder_columns")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.table.append_rows")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.table.apply_filters")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.table.sort")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.range.find_errors")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.workbook.get_summary")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.workbook.get_used_range_summary")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.sheet.get_summary")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.range.get_summary")).toBe(true);
    expect(exposed.find((tool) => tool.name === "excel.range.get_summary")?.requiredCapabilities).toContain("range.read");
    expect(exposed.some((tool) => tool.name === "excel.range.read_compact")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.table.get_schema")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.table.read_compact")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.lookup.search_workbook")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.lookup.resolve_range")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.lookup.inspect_match")).toBe(true);
    expect(exposed.find((tool) => tool.name === "excel.lookup.search_workbook")?.mutatesWorkbook).toBe(false);
    expect(exposed.find((tool) => tool.name === "excel.lookup.search_workbook")?.requiredCapabilities).toContain("range.read");
    expect(exposed.some((tool) => tool.name === "excel.validate.compact")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.snapshot.get_compact")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.snapshot.compare_compact")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.validate.no_unintended_changes")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.runtime.get_capabilities")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.runtime.get_selection")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.workflow.prepare_session")).toBe(true);
    expect(exposed.find((tool) => tool.name === "excel.workflow.prepare_session")?.mutatesWorkbook).toBe(false);
    expect(exposed.some((tool) => tool.name === "excel.workflow.create_formula_sheet")).toBe(true);
    expect(exposed.find((tool) => tool.name === "excel.workflow.create_formula_sheet")?.requiresConfirmation).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.workflow.create_template_report")).toBe(true);
    expect(exposed.find((tool) => tool.name === "excel.workflow.create_template_report")?.requiresConfirmation).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.workflow.create_pivot_chart_summary")).toBe(true);
    expect(exposed.find((tool) => tool.name === "excel.workflow.create_pivot_chart_summary")?.requiresConfirmation).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.workflow.repair_formula_errors")).toBe(true);
    expect(exposed.find((tool) => tool.name === "excel.workflow.repair_formula_errors")?.requiresConfirmation).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.workflow.preview_risky_edit")).toBe(true);
    expect(exposed.find((tool) => tool.name === "excel.workflow.preview_risky_edit")?.requiresConfirmation).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.workflow.inspect_analyze")).toBe(true);
    expect(exposed.find((tool) => tool.name === "excel.workflow.inspect_analyze")?.mutatesWorkbook).toBe(false);
    expect(exposed.some((tool) => tool.name === "excel.workflow.rollback_validate")).toBe(true);
    expect(exposed.find((tool) => tool.name === "excel.workflow.rollback_validate")?.requiresConfirmation).toBe(true);
    expect(exposed.find((tool) => tool.name === "excel.runtime.get_selection")?.status).toBe("stable");
    const catalogNames = new Set(InternalCapabilityCatalog.map((tool) => tool.name));
    for (const omitted of [
      "excel.range.read_values",
      "excel.range.read_formulas",
      "excel.range.read_number_formats",
      "excel.range.read_display_text",
      "excel.range.read_styles",
      "excel.range.read_full",
      "excel.table.read",
      "excel.snapshot.get",
      "excel.snapshot.compare",
      "excel.diff.get_details",
      "excel.diff.export_json",
      "excel.diff.export_html",
      "excel.filter.get_filters",
      "excel.filter.apply",
      "excel.filter.clear",
      "excel.filter.preserve_from_template",
      "excel.filter.validate",
      "excel.sort.apply",
      "excel.sort.clear",
      "excel.sort.preserve_from_template"
    ]) {
      expect(exposedNames.has(omitted)).toBe(false);
      expect(catalogNames.has(omitted)).toBe(false);
    }
  });

  it("summarizes public tools separately from internal backend capabilities", () => {
    const exposed = getPublicAgentToolCatalog({ includePreview: true });
    const internalSummary = getInternalCapabilityCatalogSummary({ includePreview: true });
    expect(exposed.map((tool) => tool.name)).toEqual(["excel.agent.run"]);
    expect(internalSummary.total).toBe(InternalCapabilityCatalog.length);
    expect(internalSummary.exposed).toBe(0);
    expect(internalSummary.capabilities.length).toBe(InternalCapabilityCatalog.length);
    expect(internalSummary.capabilities.some((tool) => tool.name === "excel.runtime.get_capabilities")).toBe(true);
    expect(InternalCapabilityCatalog.some((tool) => tool.status === "planned")).toBe(false);
  });

  it("tracks resources and prompts as catalog entries", () => {
    expect(ResourceCatalog.some((resource) => resource.uriTemplate === "excel://runtime/status")).toBe(true);
    expect(ResourceCatalog.some((resource) => resource.uriTemplate === "excel://compact/{resource_id}")).toBe(true);
    expect(ResourceCatalog.some((resource) => resource.uriTemplate === "excel://agent/contexts/{workbook_context_id}/semantic-index")).toBe(true);
    expect(ResourceCatalog.some((resource) => resource.uriTemplate === "excel://agent/results/{result_id}")).toBe(true);
    expect(ResourceCatalog.every((resource) => resource.status === "stable")).toBe(true);
    expect(PromptCatalog.some((prompt) => prompt.name === "excel.prompts.create_next_month_sheet" && prompt.status === "stable")).toBe(true);
    expect(PromptCatalog.some((prompt) => prompt.name === "excel.prompts.field_value_image_to_styled_table" && prompt.status === "stable")).toBe(true);
    expect(PromptCatalog.some((prompt) => prompt.name === "excel.prompts.booking_image_to_styled_table" && prompt.status === "stable")).toBe(true);
  });
});
