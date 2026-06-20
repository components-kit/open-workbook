import { getPublicAgentToolCatalog } from "@components-kit/open-workbook-protocol";
import { describe, expect, it } from "vitest";
import { AGENT_ACTION_HANDLERS } from "./agent-action-handlers.js";
import {
  EXCEL_CAPABILITY_GROUPS,
  getExcelCapability,
  getExcelCapabilityAgentStatus,
  getExcelCapabilityGroup,
  getExcelCapabilitySummary,
  listExcelCapabilityCoverage,
  listExcelCapabilities,
  listExcelCapabilityGroups,
  summarizeExcelCapabilityCoverage
} from "./excel-capabilities.js";

describe("excel capabilities", () => {
  it("keeps Excel operations available as internal backend capabilities", () => {
    const capabilities = listExcelCapabilities();

    expect(capabilities.length).toBe(301);
    expect(getExcelCapability("excel.agent.run")).toBeTruthy();
    expect(getExcelCapability("excel.range.write_values")?.mutatesWorkbook).toBe(true);
    expect(getExcelCapability("excel.workflow.inspect_analyze")?.mutatesWorkbook).toBe(false);
  });

  it("does not advertise executor-limited operations as stable or covered", () => {
    for (const capabilityName of ["excel.range.write_hyperlinks", "excel.range.write_comments", "excel.sheet.move"]) {
      const capability = getExcelCapability(capabilityName);
      if (!capability) {
        expect(listExcelCapabilityCoverage().some((entry) => entry.capability.name === capabilityName), capabilityName).toBe(false);
        continue;
      }
      expect(capability.status, capabilityName).toBe("planned");
      expect(listExcelCapabilityCoverage().find((entry) => entry.capability.name === capabilityName)?.planningStatus, capabilityName).toBe("defer");
    }
  });

  it("keeps internal capabilities separate from the public MCP tool surface", () => {
    const summary = getExcelCapabilitySummary();
    const exposed = getPublicAgentToolCatalog();

    expect(exposed.map((tool) => tool.name)).toEqual(["excel.agent.run"]);
    expect(summary.total).toBe(301);
    expect(summary.exposed).toBe(0);
    expect(summary.capabilities.some((capability) => capability.name === "excel.range.read_compact")).toBe(true);
  });

  it("keeps catalog metadata complete enough for grouped coverage planning", () => {
    const capabilities = listExcelCapabilities();
    const names = new Set<string>();

    for (const capability of capabilities) {
      expect(capability.name).toMatch(/^excel\.[a-z_]+\.[a-z0-9_]+$/);
      expect(names.has(capability.name)).toBe(false);
      names.add(capability.name);
      expect(capability.title.length).toBeGreaterThan(0);
      expect(capability.description.length).toBeGreaterThan(0);
      expect(capability.namespace.length).toBeGreaterThan(0);
      expect(capability.status).toMatch(/^(stable|preview|planned|unsupported)$/);
      expect(capability.destructiveLevel).toMatch(/^(none|values|format|structure|workbook)$/);
      expect(capability.requiresConfirmation).toBe(capability.mutatesWorkbook);
      expect(Array.isArray(capability.requiredCapabilities)).toBe(true);
    }

    expect(names.size).toBe(capabilities.length);
  });

  it("assigns every internal capability to one stable backend capability group", () => {
    const capabilities = listExcelCapabilities();
    const grouped = listExcelCapabilityGroups();
    const groupedCount = grouped.reduce((total, group) => total + group.capabilities.length, 0);
    const groupedNames = new Set(grouped.flatMap((group) => group.capabilities.map((capability) => capability.name)));

    expect(groupedCount).toBe(capabilities.length);
    expect(groupedNames.size).toBe(capabilities.length);
    expect(EXCEL_CAPABILITY_GROUPS.every((group) => group.label.length > 0 && group.description.length > 0)).toBe(true);
    expect(getExcelCapabilityGroup("excel.range.write_values")).toBe("range");
    expect(getExcelCapabilityGroup("excel.workflow.inspect_analyze")).toBe("workflow");
    expect(getExcelCapabilityGroup("excel.snapshot.get_compact")).toBe("snapshot");
    expect(getExcelCapabilityGroup("excel.backup.prune")).toBe("backup");
    expect(getExcelCapabilityGroup("excel.permissions.set_scope")).toBe("permissions");
  });

  it("records current agent status without expanding orchestration coverage", () => {
    const grouped = listExcelCapabilityGroups();
    const agentActionHandlerCount = grouped.reduce((total, group) => total + group.agentActionHandlers, 0);

    expect(getExcelCapabilityAgentStatus("excel.agent.run")).toBe("agent_entrypoint");
    expect(getExcelCapabilityAgentStatus("excel.range.write_values")).toBe("agent_action_handler");
    expect(getExcelCapabilityAgentStatus("excel.table.get_schema")).toBe("agent_action_handler");
    expect(getExcelCapabilityAgentStatus("excel.range.read_compact")).toBe("agent_action_handler");
    expect(agentActionHandlerCount).toBeGreaterThan(new Set(AGENT_ACTION_HANDLERS.map((handler) => handler.capabilityName)).size);
    expect(grouped.reduce((total, group) => total + group.agentEntrypoint, 0)).toBe(1);
  });

  it("assigns every internal capability a coverage planning status", () => {
    const capabilities = listExcelCapabilities();
    const coverage = listExcelCapabilityCoverage();
    const summary = summarizeExcelCapabilityCoverage();

    expect(coverage.length).toBe(capabilities.length);
    expect(summary.total).toBe(capabilities.length);
    expect(summary.byPlanningStatus.covered).toBe(capabilities.length);
    expect(summary.byPlanningStatus.future_orchestration_candidate).toBe(0);
    expect(summary.byPlanningStatus.needs_unit_contract).toBe(0);
    expect(summary.byPlanningStatus.host_limited).toBe(0);
    expect(summary.byGroup.reduce((total, group) => total + group.total, 0)).toBe(capabilities.length);
    expect(coverage.find((entry) => entry.capability.name === "excel.range.write_values")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.range.copy")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.range.move")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.range.autofit_rows")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.range.read_compact")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.range.get_summary")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.range.read_hyperlinks")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.range.read_comments")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.range.read_notes")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.range.read_merged_cells")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.range.read_data_validation")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.range.read_conditional_formatting")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.range.search")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.range.find_blank_cells")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.range.find_errors")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.range.write_styles_many")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.range.clear_values")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.range.insert_rows")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.range.delete_rows")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.range.insert_columns")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.range.delete_columns")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.range.merge")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.range.unmerge")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workflow.prepare_session")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workflow.create_formula_sheet")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workflow.create_template_report")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workflow.create_pivot_chart_summary")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workflow.repair_formula_errors")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workflow.preview_risky_edit")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workflow.inspect_analyze")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workflow.rollback_validate")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workbook.list_open_workbooks")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workbook.get_workbook_info")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workbook.get_workbook_map")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workbook.get_summary")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workbook.snapshot")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workbook.refresh_snapshot")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workbook.get_snapshot")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workbook.detect_external_changes")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workbook.create_backup")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workbook.restore_backup")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workbook.export_local_config")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workbook.import_local_config")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workbook.embed_local_config")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workbook.read_embedded_local_config")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workbook.import_embedded_local_config")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workbook.close")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.snapshot.create")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.snapshot.refresh")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.snapshot.list")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.snapshot.get_compact")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.snapshot.compare_compact")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.snapshot.invalidate")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.snapshot.delete")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.backup.create_file")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.backup.list")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.backup.get")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.backup.verify")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.backup.restore_file")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.backup.prune")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.backup.delete")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.backup.pin")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.backup.unpin")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.sheet.list")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.sheet.get_used_range")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.lookup.search_workbook")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.lookup.resolve_range")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.lookup.inspect_match")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.table.append_rows")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.table.update_rows")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.table.get_schema")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.table.read_compact")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.table.create")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.table.resize")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.table.reorder_columns")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.table.set_style")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.table.copy_structure")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.table.validate_against_template")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.lookup.find_headers")).toBeUndefined();
    expect(coverage.find((entry) => entry.capability.name === "excel.lookup.find_tables_by_columns")).toBeUndefined();
    expect(coverage.find((entry) => entry.capability.name === "excel.lookup.find_entity")).toBeUndefined();
    expect(coverage.find((entry) => entry.capability.name === "excel.table.preserve_filters")).toBeUndefined();
    expect(coverage.find((entry) => entry.capability.name === "excel.sheet.create")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.sheet.copy")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.sheet.rename")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.sheet.delete")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.sheet.hide")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.sheet.unhide")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.sheet.protect")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.sheet.unprotect")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.sheet.clear")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.sheet.set_tab_color")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.names.list")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.names.create")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.names.update")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.names.delete")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.region.detect")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.region.register")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.region.list")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.region.get")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.region.clear_values")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.region.write_values")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.region.fill")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.template.detect_templates")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.template.register")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.template.unregister")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.template.get")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.template.list")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.template.infer_regions")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.template.clear_data_regions")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.template.fill_regions")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.template.validate_sheet_against_template")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.template.repair_sheet_from_template")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.style.get_fingerprint")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.style.compare_fingerprint")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.style.copy_from_template")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.style.apply_style")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.style.validate_consistency")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.style.repair_consistency")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.style.get_theme")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.style.apply_theme")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.style.copy_column_widths")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.style.copy_row_heights")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.style.copy_borders")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.style.copy_fills")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.style.copy_fonts")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.style.copy_alignment")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.style.copy_number_formats")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.style.copy_conditional_formatting")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.style.copy_data_validation")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.clean.detect_header_row")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.clean.normalize_headers")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.clean.trim_whitespace")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.clean.remove_duplicates")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.clean.parse_dates")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.clean.parse_numbers")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.clean.standardize_currency")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.clean.fill_missing_values")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.clean.split_column")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.clean.merge_columns")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.clean.detect_outliers")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.clean.fuzzy_match")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.repair.style_from_template")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.repair.formulas_from_template")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.repair.filters_from_template")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.repair.table_structure")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.repair.print_layout")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.repair.named_ranges")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.repair.formula_errors")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.repair.merged_cells")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.validate.workbook")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.validate.compact")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.validate.sheet")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.validate.template_consistency")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.validate.formulas")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.validate.styles")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.validate.tables")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.validate.filters")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.validate.print_layout")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.validate.no_broken_references")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.validate.no_formula_errors")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.validate.no_unintended_changes")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.formula.read_patterns")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.formula.copy_patterns")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.formula.fill_down")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.formula.fill_right")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.formula.validate")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.formula.validate_against_template")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.formula.repair_patterns")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.formula.find_errors")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.formula.get_dependency_graph")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.formula.trace_precedents")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.formula.trace_dependents")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.formula.convert_to_values")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.formula.recalculate")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.formula.explain")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.names.get")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.runtime.get_status")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.runtime.get_active_context")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.runtime.get_selection")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.runtime.get_capabilities")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workbook.save_as")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.workbook.export_copy")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.pivot.update_source")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.chart.create")?.planningStatus).toBe("covered");
    expect(coverage.find((entry) => entry.capability.name === "excel.formula.find_circular_references")).toBeUndefined();
    expect(coverage.find((entry) => entry.capability.name === "excel.style.copy_freeze_panes")).toBeUndefined();
  });
});
