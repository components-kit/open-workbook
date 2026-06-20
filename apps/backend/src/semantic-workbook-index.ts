import type { AgentIntentAction, AgentSemanticIndexEntry, AgentSemanticRole, AgentSemanticWorkbookIndex } from "@components-kit/open-workbook-protocol";
import { normalizeHeaderName, type SheetKind, type WorkbookMetadata } from "./workbook-metadata-cache.js";

export function buildSemanticWorkbookIndex(metadata: WorkbookMetadata, options: { maxEntries?: number } = {}): AgentSemanticWorkbookIndex {
  const entries = semanticIndexEntries(metadata)
    .sort((left, right) => right.confidence - left.confidence || left.label.localeCompare(right.label))
    .slice(0, options.maxEntries ?? 50);
  return {
    kind: "semantic_workbook_index",
    source: "cached_metadata",
    workbook: {
      name: metadata.workbook.name,
      sheetCount: metadata.workbook.sheetCount,
      ...(metadata.workbook.workbookId !== undefined ? { workbookId: metadata.workbook.workbookId } : {}),
      ...(metadata.workbook.activeSheet !== undefined ? { activeSheet: metadata.workbook.activeSheet } : {})
    },
    detailLevel: metadata.detailLevel,
    entryCount: entries.length,
    entries
  };
}

export function semanticIndexEntries(metadata: WorkbookMetadata): AgentSemanticIndexEntry[] {
  const entries: AgentSemanticIndexEntry[] = [];
  entries.push({
    id: "semantic:workbook",
    label: metadata.workbook.name,
    role: "workbook",
    sourceKind: "workbook",
    aliases: compactAliases([metadata.workbook.name, "workbook", "excel workbook"]),
    confidence: 1,
    evidence: [`${metadata.sheets.length} sheet(s)`, `${metadata.tables.length} table(s)`, `${metadata.namedRanges.length} named range(s)`],
    supportedActions: ["get_workbook_info", "read_schema", "find_target", "validate_workbook"]
  });

  for (const sheet of metadata.sheets) {
    const columns = sheet.headers.flatMap((header) => header.columns.map((column) => column.name));
    const role = roleForSheetKind(sheet.kind);
    entries.push({
      id: sheet.id,
      label: sheet.name,
      role,
      sourceKind: "sheet",
      sheetName: sheet.name,
      ...(sheet.usedRange !== undefined ? { range: sheet.usedRange } : {}),
      aliases: compactAliases([sheet.name, sheet.kind, roleLabel(role), ...columns.slice(0, 8)]),
      confidence: sheetConfidence(sheet.kind, sheet.headers.length, sheet.tableIds.length),
      evidence: compactEvidence([
        sheet.usedRange ? `used range ${sheet.usedRange}` : undefined,
        sheet.headers.length > 0 ? `${sheet.headers.length} detected header row(s)` : undefined,
        sheet.tableIds.length > 0 ? `${sheet.tableIds.length} table(s)` : undefined
      ]),
      supportedActions: actionsForRole(role)
    });
  }

  for (const table of metadata.tables) {
    entries.push({
      id: table.id,
      label: table.name ?? table.id,
      role: "data_table",
      sourceKind: "table",
      sheetName: table.sheetName,
      ...(table.name !== undefined ? { tableName: table.name } : {}),
      range: table.range,
      aliases: compactAliases([table.name ?? table.id, table.sheetName, "table", "data table", ...table.columns.map((column) => column.name)]),
      confidence: table.columns.length > 0 ? 0.9 : 0.75,
      evidence: compactEvidence([table.range ? `table range ${table.range}` : undefined, `${table.columns.length} column(s)`]),
      supportedActions: ["read_schema", "read_values", "append_table_rows", "update_table_rows", "sort_table", "filter_range", "validate_tables"]
    });
  }

  for (const section of metadata.sections) {
    const role = roleForSection(section.label, section.kind, section.columns.map((column) => column.name));
    entries.push({
      id: section.id,
      label: section.label,
      role,
      sourceKind: "region",
      sheetName: section.sheetName,
      range: section.range,
      aliases: compactAliases([section.label, section.kind, section.sheetName, ...section.labels, ...section.columns.map((column) => column.name)]),
      confidence: Math.max(section.confidence, role === "unknown" ? 0.45 : 0.7),
      evidence: compactEvidence([`section ${section.range}`, section.headerRange ? `header ${section.headerRange}` : undefined, `${section.nonEmptyCellCount} non-empty cell(s)`]),
      supportedActions: actionsForRole(role)
    });
  }

  for (const name of metadata.namedRanges) {
    entries.push({
      id: `name:${name.name}`,
      label: name.name,
      role: "named_region",
      sourceKind: "region",
      ...(name.sheetName !== undefined ? { sheetName: name.sheetName } : {}),
      range: name.range,
      aliases: compactAliases([name.name, splitIdentifierWords(name.name), name.sheetName ?? "", "named range", "region"]),
      confidence: 0.78,
      evidence: [`named range ${name.range}`],
      supportedActions: ["read_named_item", "read_region", "write_region_values", "clear_region_values", "find_target"]
    });
  }

  for (const block of metadata.summaryBlocks) {
    entries.push({
      id: block.id,
      label: block.labels[0] ?? "summary block",
      role: "summary_sheet",
      sourceKind: "range",
      sheetName: block.sheetName,
      range: block.range,
      aliases: compactAliases([block.sheetName, "summary", "metrics", "report", ...block.labels]),
      confidence: block.confidence,
      evidence: [`summary block ${block.range}`],
      supportedActions: ["read_schema", "read_values", "validate_workbook", "find_target"]
    });
  }

  for (const region of metadata.formulaRegions) {
    entries.push({
      id: region.id,
      label: "formula region",
      role: "formula_region",
      sourceKind: "range",
      sheetName: region.sheetName,
      range: region.range,
      aliases: compactAliases([region.sheetName, "formula", "calculation", "formula region"]),
      confidence: Math.min(0.9, 0.65 + region.formulaCount * 0.01),
      evidence: [`${region.formulaCount} formula cell(s)`],
      supportedActions: ["read_formula_patterns", "validate_formula_range", "find_formula_errors", "repair_formula_patterns"]
    });
  }

  if (metadata.selection) {
    entries.push({
      id: "semantic:selection",
      label: "current selection",
      role: "selection",
      sourceKind: "selection",
      sheetName: metadata.selection.sheetName,
      range: metadata.selection.address,
      aliases: compactAliases(["current selection", "active selection", "selected range", metadata.selection.sheetName]),
      confidence: 0.95,
      evidence: [`selection ${metadata.selection.sheetName}!${metadata.selection.address}`],
      supportedActions: ["read_values", "read_style_summary", "format_diagnostics", "write_values", "format_range", "clear_style_dimensions"]
    });
  }

  return dedupeEntries(entries);
}

function roleForSheetKind(kind: SheetKind): AgentSemanticRole {
  if (kind === "transaction") return "transaction_sheet";
  if (kind === "summary" || kind === "dashboard") return "summary_sheet";
  if (kind === "template") return "template_sheet";
  if (kind === "lookup") return "lookup_sheet";
  return "unknown";
}

function roleForSection(label: string, kind: string, columns: string[]): AgentSemanticRole {
  const text = normalizeHeaderName([label, kind, ...columns].join(" "));
  if (/template/.test(text)) return "template_sheet";
  if (/invoice|receipt|booking|customer|client|vendor|form|field|contact|return|pickup/.test(text)) return "form_region";
  if (/formula|calculation|reconciliation|variance/.test(text)) return "formula_region";
  if (/summary|kpi|metric|total|revenue|expense|profit|dashboard/.test(text)) return "summary_sheet";
  if (/transaction|payment|status|amount|date|account|container/.test(text)) return "transaction_sheet";
  return "unknown";
}

function sheetConfidence(kind: SheetKind, headerCount: number, tableCount: number): number {
  const base = kind === "unknown" ? 0.5 : 0.72;
  return Number(Math.min(0.95, base + Math.min(headerCount, 3) * 0.05 + Math.min(tableCount, 2) * 0.06).toFixed(3));
}

function actionsForRole(role: AgentSemanticRole): AgentIntentAction[] {
  if (role === "template_sheet") {
    return ["read_template", "copy_template_sheet", "fill_template_regions", "validate_sheet_against_template", "repair_sheet_from_template", "copy_style_from_template"];
  }
  if (role === "summary_sheet") {
    return ["read_schema", "read_values", "validate_workbook", "create_pivot_chart_summary", "find_target"];
  }
  if (role === "formula_region") {
    return ["read_formula_patterns", "validate_formula_range", "find_formula_errors", "repair_formula_patterns"];
  }
  if (role === "form_region") {
    return ["read_values", "write_values", "replace_range_with_styled_table", "read_style_summary", "format_diagnostics"];
  }
  return ["read_schema", "read_values", "write_values", "read_style_summary", "format_diagnostics", "find_target"];
}

function roleLabel(role: AgentSemanticRole): string {
  return role.replace(/_/g, " ");
}

function compactAliases(values: Array<string | undefined>): string[] {
  const aliases = values
    .flatMap((value) => typeof value === "string" ? [value, splitIdentifierWords(value)] : [])
    .map((value) => value.trim())
    .filter(Boolean);
  return [...new Set(aliases)].slice(0, 16);
}

function compactEvidence(values: Array<string | undefined>): string[] {
  return values.filter((value): value is string => typeof value === "string" && value.trim().length > 0).slice(0, 6);
}

function splitIdentifierWords(value: string): string {
  return value.replace(/([a-z])([A-Z])/g, "$1 $2").replace(/[_-]+/g, " ");
}

function dedupeEntries(entries: AgentSemanticIndexEntry[]): AgentSemanticIndexEntry[] {
  const seen = new Set<string>();
  return entries.filter((entry) => {
    const key = `${entry.id}:${entry.sheetName ?? ""}:${entry.range ?? ""}`;
    if (seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}
