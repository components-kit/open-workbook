import { describe, expect, it } from "vitest";
import { getExposedToolCatalog, ToolCatalog } from "./tools.js";
import { PromptCatalog } from "./prompts.js";
import { ResourceCatalog } from "./resources.js";

describe("tool catalog", () => {
  it("contains the full requested tool skeleton", () => {
    expect(ToolCatalog.length).toBeGreaterThan(200);
    expect(ToolCatalog.some((tool) => tool.name === "excel.runtime.get_capabilities")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.name === "excel.workbook.get_workbook_map")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.name === "excel.reconcile.bank_statement")).toBe(true);
    expect(ToolCatalog.some((tool) => tool.name === "excel.ocr.extract_invoice" && tool.status === "unsupported")).toBe(true);
  });

  it("exposes only stable tools by default", () => {
    const exposed = getExposedToolCatalog();
    expect(exposed.every((tool) => tool.status === "stable")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.runtime.get_status")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.template.validate_sheet_against_template")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.style.repair_consistency")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.style.compare_fingerprint")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.style.copy_fills")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.style.copy_number_formats")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.style.copy_data_validation")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.formula.repair_patterns")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.formula.read_patterns")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.formula.copy_patterns")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.formula.fill_down")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.formula.convert_to_values")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.formula.explain")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.table.append_rows")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.filter.apply")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.sort.apply")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.range.read_data_validation")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.range.find_errors")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.workbook.close")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.workbook.export_copy")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.validate.workbook")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.validate.no_unintended_changes")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.repair.style_from_template")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.repair.table_structure")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.names.create")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.region.fill")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.permissions.lock_regions")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.clean.trim_whitespace")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.clean.fuzzy_match")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.pivot.create")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.pivot.delete")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.chart.create")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.runtime.get_capabilities")).toBe(false);
    expect(exposed.some((tool) => tool.name === "excel.ocr.read_file")).toBe(false);
  });

  it("can expose preview tools without exposing planned or unsupported tools", () => {
    const exposed = getExposedToolCatalog({ includePreview: true });
    expect(exposed.some((tool) => tool.name === "excel.runtime.get_capabilities")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.workbook.get_workbook_map" && tool.status === "stable")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.pivot.create" && tool.status === "stable")).toBe(true);
    expect(exposed.some((tool) => tool.name === "excel.ocr.read_file")).toBe(false);
  });

  it("tracks resources and prompts as catalog entries", () => {
    expect(ResourceCatalog.some((resource) => resource.uriTemplate === "excel://runtime/status")).toBe(true);
    expect(ResourceCatalog.every((resource) => resource.status === "stable")).toBe(true);
    expect(PromptCatalog.some((prompt) => prompt.name === "excel.prompts.create_next_month_sheet" && prompt.status === "stable")).toBe(true);
    expect(PromptCatalog.some((prompt) => prompt.name === "excel.prompts.reconcile_statement" && prompt.status === "unsupported")).toBe(true);
  });
});
