import { describe, expect, it } from "vitest";
import { AGENT_ACTION_HANDLERS, findAgentActionHandler } from "./agent-action-handlers.js";
import { AGENT_INTENT_ACTIONS, type AgentRunInput } from "@components-kit/open-workbook-protocol";
import { getExcelCapability } from "./excel-capabilities.js";

describe("agent action handlers", () => {
  it("declares stable handler metadata", () => {
    expect(AGENT_ACTION_HANDLERS.length).toBeGreaterThanOrEqual(8);
    const protocolActions = new Set<string>(AGENT_INTENT_ACTIONS);
    const claimedScopedActions = new Set<string>();
    for (const handler of AGENT_ACTION_HANDLERS) {
      expect(handler.id).toMatch(/^[a-z0-9_]+$/);
      expect(handler.riskKind).toBeTruthy();
      expect(getExcelCapability(handler.capabilityName), handler.capabilityName).toBeTruthy();
      expect(typeof handler.matches).toBe("function");
      if (handler.intentAction) {
        expect(protocolActions.has(handler.intentAction), handler.id).toBe(true);
        const scopedAction = `${handler.requiresResolvedTarget ? "target" : "workbook"}:${handler.intentAction}`;
        expect(claimedScopedActions.has(scopedAction), scopedAction).toBe(false);
        claimedScopedActions.add(scopedAction);
      }
    }
  });

  it("matches caller intent and natural language to the same handler", () => {
    const hinted: AgentRunInput = {
      request: "Do it",
      intent: { action: "format_range" },
      target: { sheetName: "Data", range: "A1:D1" }
    };
    const natural: AgentRunInput = {
      request: "Format the header row on Data",
      target: { sheetName: "Data", range: "A1:D1" }
    };

    expect(findAgentActionHandler(hinted, "format_range", true)?.id).toBe("format_range");
    expect(findAgentActionHandler(natural, undefined, true)?.id).toBe("format_range");
  });

  it("matches promoted range-core actions by caller intent", () => {
    const target: AgentRunInput["target"] = { sheetName: "Data", range: "A1:B2" };

    expect(findAgentActionHandler({ request: "Do it", intent: { action: "write_formulas" }, target }, "write_formulas", true)?.id).toBe("write_formulas");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "write_number_formats" }, target }, "write_number_formats", true)?.id).toBe("write_number_formats");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "clear_range" }, target }, "clear_range", true)?.id).toBe("clear_range");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "clear_formats" }, target }, "clear_formats", true)?.id).toBe("clear_formats");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "autofit_rows" }, target }, "autofit_rows", true)?.id).toBe("autofit_rows");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "copy_range" } }, "copy_range", false)?.id).toBe("copy_range");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "move_range" } }, "move_range", false)?.id).toBe("move_range");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "reorder_range_columns" }, target }, "reorder_range_columns", true)?.id).toBe("reorder_range_columns");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "clear_values_raw" }, target }, "clear_values_raw", true)?.id).toBe("clear_values_raw");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "write_styles_many" } }, "write_styles_many", false)?.id).toBe("write_styles_many");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "write_data_validation" }, target }, "write_data_validation", true)?.id).toBe("write_data_validation");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "write_conditional_formatting" }, target }, "write_conditional_formatting", true)?.id).toBe("write_conditional_formatting");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "insert_rows" }, target }, "insert_rows", true)?.id).toBe("insert_rows");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "delete_rows" }, target }, "delete_rows", true)?.id).toBe("delete_rows");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "insert_columns" }, target }, "insert_columns", true)?.id).toBe("insert_columns");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "delete_columns" }, target }, "delete_columns", true)?.id).toBe("delete_columns");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "merge_range" }, target }, "merge_range", true)?.id).toBe("merge_range");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "unmerge_range" }, target }, "unmerge_range", true)?.id).toBe("unmerge_range");
  });

  it("matches promoted table-core actions by caller intent", () => {
    const target: AgentRunInput["target"] = { tableName: "Transactions" };

    expect(findAgentActionHandler({ request: "Do it", intent: { action: "append_table_rows" }, target }, "append_table_rows", true)?.id).toBe("append_table_rows");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "update_table_rows" }, target }, "update_table_rows", true)?.id).toBe("update_table_rows");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "create_table" } }, "create_table", false)?.id).toBe("create_table");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "resize_table" }, target }, "resize_table", true)?.id).toBe("resize_table");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "reorder_table_columns" }, target }, "reorder_table_columns", true)?.id).toBe("reorder_table_columns");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "clear_table_data" }, target }, "clear_table_data", true)?.id).toBe("clear_table_data");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "clear_table_filters" }, target }, "clear_table_filters", true)?.id).toBe("clear_table_filters");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "set_table_total_row" }, target }, "set_table_total_row", true)?.id).toBe("set_table_total_row");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "set_table_style" }, target }, "set_table_style", true)?.id).toBe("set_table_style");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "copy_table_structure" }, target }, "copy_table_structure", true)?.id).toBe("copy_table_structure");
  });

  it("matches promoted sheet-core actions by caller intent", () => {
    const target: AgentRunInput["target"] = { sheetName: "Report" };

    expect(findAgentActionHandler({ request: "Do it", intent: { action: "create_sheet" } }, "create_sheet", false)?.id).toBe("create_sheet");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "copy_sheet" }, target }, "copy_sheet", false)?.id).toBe("copy_sheet");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "rename_sheet" }, target }, "rename_sheet", false)?.id).toBe("rename_sheet");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "delete_sheet" }, target }, "delete_sheet", false)?.id).toBe("delete_sheet");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "hide_sheet" }, target }, "hide_sheet", false)?.id).toBe("hide_sheet");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "unhide_sheet" }, target }, "unhide_sheet", false)?.id).toBe("unhide_sheet");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "protect_sheet" }, target }, "protect_sheet", false)?.id).toBe("protect_sheet");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "unprotect_sheet" }, target }, "unprotect_sheet", false)?.id).toBe("unprotect_sheet");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "clear_sheet" }, target }, "clear_sheet", false)?.id).toBe("clear_sheet");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "set_sheet_tab_color" }, target }, "set_sheet_tab_color", false)?.id).toBe("set_sheet_tab_color");
  });

  it("matches promoted workbook mutation actions by caller intent", () => {
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "restore_workbook_backup" } }, "restore_workbook_backup", false)?.id).toBe("restore_workbook_backup");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "import_local_config" } }, "import_local_config", false)?.id).toBe("import_local_config");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "embed_local_config" } }, "embed_local_config", false)?.id).toBe("embed_local_config");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "import_embedded_local_config" } }, "import_embedded_local_config", false)?.id).toBe("import_embedded_local_config");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "close_workbook" } }, "close_workbook", false)?.id).toBe("close_workbook");
  });

  it("matches promoted formula mutation actions by caller intent", () => {
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "copy_formula_patterns" } }, "copy_formula_patterns", false)?.id).toBe("copy_formula_patterns");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "fill_formula_down" } }, "fill_formula_down", false)?.id).toBe("fill_formula_down");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "fill_formula_right" } }, "fill_formula_right", false)?.id).toBe("fill_formula_right");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "repair_formula_patterns" } }, "repair_formula_patterns", false)?.id).toBe("repair_formula_patterns");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "convert_formulas_to_values" }, target: { sheetName: "Report", range: "B2" } }, "convert_formulas_to_values", true)?.id).toBe("convert_formulas_to_values");
  });

  it("matches promoted name mutation actions by caller intent", () => {
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "create_name" } }, "create_name", false)?.id).toBe("create_name");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "update_name" } }, "update_name", false)?.id).toBe("update_name");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "delete_name" } }, "delete_name", false)?.id).toBe("delete_name");
  });

  it("matches promoted region mutation actions by caller intent", () => {
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "register_region" } }, "register_region", false)?.id).toBe("register_region");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "clear_region_values" } }, "clear_region_values", false)?.id).toBe("clear_region_values");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "write_region_values" } }, "write_region_values", false)?.id).toBe("write_region_values");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "fill_region" } }, "fill_region", false)?.id).toBe("fill_region");
  });

  it("matches promoted template mutation actions by caller intent", () => {
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "register_template" } }, "register_template", false)?.id).toBe("register_template");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "unregister_template" } }, "unregister_template", false)?.id).toBe("unregister_template");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "clear_template_data_regions" } }, "clear_template_data_regions", false)?.id).toBe("clear_template_data_regions");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "fill_template_regions" } }, "fill_template_regions", false)?.id).toBe("fill_template_regions");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "repair_sheet_from_template" } }, "repair_sheet_from_template", false)?.id).toBe("repair_sheet_from_template");
    expect(findAgentActionHandler({ request: "Register this sheet as a template" }, undefined, false)?.id).toBe("register_template");
    expect(findAgentActionHandler({ request: "Repair Q2 Report from the template" }, undefined, false)?.id).toBe("repair_sheet_from_template");
    expect(findAgentActionHandler({ request: "Create a new report from the template" }, undefined, false)?.id).toBe("copy_template_sheet");
  });

  it("matches promoted style mutation actions by caller intent", () => {
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "copy_style_from_template" } }, "copy_style_from_template", false)?.id).toBe("copy_style_from_template");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "repair_style_consistency" } }, "repair_style_consistency", false)?.id).toBe("repair_style_consistency");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "repair_style_from_template" } }, "repair_style_from_template", false)?.id).toBe("repair_style_from_template");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "repair_formulas_from_template" } }, "repair_formulas_from_template", false)?.id).toBe("repair_formulas_from_template");
    expect(findAgentActionHandler({ request: "Copy style from the template" }, undefined, false)?.id).toBe("copy_style_from_template");
    expect(findAgentActionHandler({ request: "Match the same style from source to target" }, undefined, false)?.id).toBe("copy_style_from_template");
    expect(findAgentActionHandler({ request: "Repair style consistency" }, undefined, false)?.id).toBe("repair_style_consistency");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "repair_table_structure" }, target: { tableName: "Transactions" } }, "repair_table_structure", true)?.id).toBe("repair_table_structure");
  });

  it("does not let range formatting and filtering requests fall through to sheet creation", () => {
    expect(findAgentActionHandler({ request: "Add autofilter to the header row of Booking sheet range A1:X7" }, undefined, false)).toBeUndefined();
    expect(findAgentActionHandler({ request: "Add autofilter to the header row of Booking sheet range A1:X7", target: { sheetName: "Booking", range: "A1:X7" } }, undefined, true)?.id).toBe("filter_range");
    expect(findAgentActionHandler({ request: "Add borders to Booking sheet range A1:X7" }, undefined, false)).toBeUndefined();
    expect(findAgentActionHandler({ request: "Add borders to Booking sheet range A1:X7", target: { sheetName: "Booking", range: "A1:X7" } }, undefined, true)?.id).toBe("format_range");
    expect(findAgentActionHandler({ request: "format_range action: Apply borders to range Booking!A1:X7", target: { sheetName: "Booking", range: "A1:X7" } }, undefined, true)?.id).toBe("format_range");
    expect(findAgentActionHandler({ request: "Remove all filters from Booking sheet range A1:X7" }, undefined, false)).toBeUndefined();
    expect(findAgentActionHandler({ request: "Remove all filters from Booking sheet range A1:X7", target: { sheetName: "Booking", range: "A1:X7" } }, undefined, true)?.id).toBe("clear_table_filters");
    expect(findAgentActionHandler({ request: "Add conditional formatting rule on Booking sheet range A2:X20. Formula =$D2=\"40HQ\" should fill the row yellow." }, undefined, false)).toBeUndefined();
    expect(findAgentActionHandler({ request: "Add conditional formatting rule on Booking sheet range A2:X20. Formula =$D2=\"40HQ\" should fill the row yellow.", target: { sheetName: "Booking", range: "A2:X20" } }, undefined, true)?.id).toBe("write_conditional_formatting");
    expect(findAgentActionHandler({ request: "Add data validation dropdown list to Booking D2:D7.", target: { sheetName: "Booking", range: "D2:D7" } }, undefined, true)?.id).toBe("write_data_validation");
    expect(findAgentActionHandler({ request: "Add new col next to Qty", target: { sheetName: "Booking", range: "D:D" } }, undefined, true)?.id).toBe("insert_columns");
  });

  it("matches promoted cleaning mutation actions by caller intent", () => {
    const target = { sheetName: "Data", range: "A1:D4" };
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "normalize_headers" }, target }, "normalize_headers", true)?.id).toBe("normalize_headers");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "trim_whitespace" }, target }, "trim_whitespace", true)?.id).toBe("trim_whitespace");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "remove_duplicates" }, target }, "remove_duplicates", true)?.id).toBe("remove_duplicates");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "parse_dates" }, target }, "parse_dates", true)?.id).toBe("parse_dates");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "parse_numbers" }, target }, "parse_numbers", true)?.id).toBe("parse_numbers");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "standardize_currency" }, target }, "standardize_currency", true)?.id).toBe("standardize_currency");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "fill_missing_values" }, target }, "fill_missing_values", true)?.id).toBe("fill_missing_values");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "split_column" }, target }, "split_column", true)?.id).toBe("split_column");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "merge_columns" }, target }, "merge_columns", true)?.id).toBe("merge_columns");
  });

  it("keeps workbook-level and target-level handlers separate", () => {
    const save: AgentRunInput = { request: "Save the workbook", intent: { action: "save" } };
    const filter: AgentRunInput = { request: "Add filters", intent: { action: "filter_range" }, target: { sheetName: "Data", range: "A1:D4" } };

    expect(findAgentActionHandler(save, "save", false)?.id).toBe("save_workbook");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "create_snapshot" } }, "create_snapshot", false)?.id).toBe("create_snapshot");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "create_backup" } }, "create_backup", false)?.id).toBe("create_backup");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "refresh_snapshot" } }, "refresh_snapshot", false)?.id).toBe("refresh_snapshot");
    expect(findAgentActionHandler(save, "save", true)).toBeUndefined();
    expect(findAgentActionHandler(filter, "filter_range", false)).toBeUndefined();
    expect(findAgentActionHandler(filter, "filter_range", true)?.id).toBe("filter_range");
  });

  it("matches backup lifecycle mutation actions by caller intent", () => {
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "create_file_backup" } }, "create_file_backup", false)?.id).toBe("create_file_backup");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "restore_file_backup" } }, "restore_file_backup", false)?.id).toBe("restore_file_backup");
    expect(findAgentActionHandler({ request: "Do it", intent: { action: "prune_backups" } }, "prune_backups", false)?.id).toBe("prune_backups");
  });
});
