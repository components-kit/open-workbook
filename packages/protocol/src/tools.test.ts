import { describe, expect, it } from "vitest";
import { getExposedToolCatalog, getInternalCapabilityCatalog, getInternalCapabilityCatalogSummary, ToolCatalog } from "./tools.js";
import { PromptCatalog } from "./prompts.js";
import { ResourceCatalog } from "./resources.js";

describe("tool catalog", () => {
  it("contains the full requested tool skeleton", () => {
    expect(ToolCatalog.length).toBeGreaterThan(200);
    expect(ToolCatalog.some((tool) => tool.name === "excel.agent.run")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.name === "excel.runtime.get_capabilities")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.name === "excel.workbook.get_workbook_map")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.name === "excel.workbook.get_summary")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.name === "excel.range.read_compact")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.name === "excel.table.read_compact")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.name === "excel.lookup.search_workbook")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.name === "excel.lookup.inspect_match")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.name === "excel.compact.get_resource")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.name === "excel.compact.gc_resources")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.name === "excel.compact.context_stats")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.name === "excel.validate.compact")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.name === "excel.snapshot.get_compact")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.name === "excel.diff.get_compact")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.name === "excel.workflow.prepare_session")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.name === "excel.workflow.create_formula_sheet")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.name === "excel.workflow.create_template_report")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.name === "excel.workflow.create_pivot_chart_summary")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.name === "excel.workflow.repair_formula_errors")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.name === "excel.workflow.preview_risky_edit")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.name === "excel.workflow.inspect_analyze")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.name === "excel.workflow.rollback_validate")).toBe(true);
  });

  it("exposes only the public agent tool through the MCP catalog", () => {
    const exposed = getExposedToolCatalog();
    expect(exposed.map((tool) => tool.name)).toEqual(["excel.agent.run"]);
    expect(exposed.every((tool) => tool.status === "stable")).toBe(true);
  });

  it("keeps compact and workflow capabilities in the internal backend catalog", () => {
    const exposed = getInternalCapabilityCatalog();
    const exposedNames = new Set(exposed.map((tool) => tool.name));
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
    expect(exposed.some((tool) => tool.name === "excel.lookup.find_headers")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.lookup.find_tables_by_columns")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.lookup.find_entity")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.lookup.resolve_range")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.lookup.inspect_match")).toBe(true);
    expect(exposed.find((tool) => tool.name === "excel.lookup.search_workbook")?.mutatesWorkbook).toBe(false);
    expect(exposed.find((tool) => tool.name === "excel.lookup.search_workbook")?.requiredCapabilities).toContain("range.read");
    expect(exposed.some((tool) => tool.name === "excel.compact.get_resource")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.compact.clear_cache")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.compact.gc_resources")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.compact.context_stats")).toBe(true);
    expect(exposed.find((tool) => tool.name === "excel.compact.clear_cache")?.mutatesWorkbook).toBe(false);
    expect(exposed.some((tool) => tool.name === "excel.validate.compact")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.snapshot.get_compact")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.snapshot.compare_compact")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.diff.get_compact")).toBe(true);
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
    const catalogNames = new Set(ToolCatalog.map((tool) => tool.name));
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
    const exposed = getExposedToolCatalog({ includePreview: true });
    const internalSummary = getInternalCapabilityCatalogSummary({ includePreview: true });
    expect(exposed.map((tool) => tool.name)).toEqual(["excel.agent.run"]);
    expect(internalSummary.total).toBe(ToolCatalog.length);
    expect(internalSummary.exposed).toBe(0);
    expect(internalSummary.capabilities.length).toBe(ToolCatalog.length);
    expect(internalSummary.capabilities.some((tool) => tool.name === "excel.runtime.get_capabilities")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.status === "planned")).toBe(false);
  });

  it("tracks resources and prompts as catalog entries", () => {
    expect(ResourceCatalog.some((resource) => resource.uriTemplate === "excel://runtime/status")).toBe(true);
    expect(ResourceCatalog.some((resource) => resource.uriTemplate === "excel://compact/{resource_id}")).toBe(true);
    expect(ResourceCatalog.every((resource) => resource.status === "stable")).toBe(true);
    expect(PromptCatalog.some((prompt) => prompt.name === "excel.prompts.create_next_month_sheet" && prompt.status === "stable")).toBe(true);
  });
});
