#!/usr/bin/env node
import { getExposedToolCatalog } from "../packages/protocol/dist/tools.js";

const exposedTools = new Set(getExposedToolCatalog({ includePreview: true }).map((tool) => tool.name));

const required = [
  "excel.workflow.prepare_session",
  "excel.workflow.inspect_analyze",
  "excel.workflow.preview_risky_edit",
  "excel.range.read_compact",
  "excel.table.read_compact",
  "excel.validate.compact",
  "excel.snapshot.get_compact",
  "excel.snapshot.compare_compact",
  "excel.diff.get_compact",
  "excel.compact.get_resource",
  "excel.compact.context_stats",
  "excel.batch.apply",
  "excel.range.write_values",
  "excel.range.write_formulas",
  "excel.range.write_styles_many",
  "excel.table.apply_filters",
  "excel.table.sort"
];

const forbidden = [
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
];

const missing = required.filter((name) => !exposedTools.has(name));
const leaked = forbidden.filter((name) => exposedTools.has(name));

if (missing.length > 0 || leaked.length > 0) {
  console.error("Internal tool catalog validation failed.");
  if (missing.length > 0) {
    console.error(`Missing required optimized tools:\n${missing.map((name) => `- ${name}`).join("\n")}`);
  }
  if (leaked.length > 0) {
    console.error(`Raw/full or duplicate tools leaked into the optimized surface:\n${leaked.map((name) => `- ${name}`).join("\n")}`);
  }
  process.exit(1);
}

console.log(`Internal tool catalog check passed: ${exposedTools.size} stable capabilities, ${forbidden.length} raw/full or duplicate tools excluded.`);
