#!/usr/bin/env node
import { readFileSync } from "node:fs";

const source = readFileSync(new URL("../apps/mcp-server/src/index.ts", import.meta.url), "utf8");
const compactTools = parseStringSet(source, "COMPACT_PROFILE_TOOLS");

const required = [
  "excel.range.read_compact",
  "excel.table.read_compact",
  "excel.validate.compact",
  "excel.snapshot.get_compact",
  "excel.snapshot.compare_compact",
  "excel.diff.get_compact",
  "excel.compact.get_resource",
  "excel.compact.context_stats",
  "excel.workflow.preview_risky_edit",
  "excel.workflow.inspect_analyze",
  "excel.workflow.rollback_validate"
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
  "excel.diff.get_details",
  "excel.diff.export_json",
  "excel.diff.export_html"
];

const missing = required.filter((name) => !compactTools.has(name));
const leaked = forbidden.filter((name) => compactTools.has(name));

if (missing.length > 0 || leaked.length > 0) {
  console.error("Compact profile validation failed.");
  if (missing.length > 0) {
    console.error(`Missing compact-profile tools:\n${missing.map((name) => `- ${name}`).join("\n")}`);
  }
  if (leaked.length > 0) {
    console.error(`Raw/full payload tools leaked into compact profile:\n${leaked.map((name) => `- ${name}`).join("\n")}`);
  }
  process.exit(1);
}

console.log(`Compact profile check passed: ${compactTools.size} compact tools, ${forbidden.length} raw/full tools excluded.`);

function parseStringSet(text, constName) {
  const declaration = text.match(new RegExp(`const ${constName} = new Set\\(\\[`));
  if (declaration?.index === undefined) {
    throw new Error(`Unable to find ${constName}.`);
  }
  const arrayStart = text.indexOf("[", declaration.index);
  const arrayEnd = findMatching(text, arrayStart, "[", "]");
  if (arrayEnd < 0) {
    throw new Error(`Unable to parse ${constName}.`);
  }
  return new Set([...text.slice(arrayStart, arrayEnd + 1).matchAll(/"([^"]+)"/g)].map((match) => match[1]));
}

function findMatching(text, startIndex, open, close) {
  let depth = 0;
  let quote = undefined;
  for (let index = startIndex; index < text.length; index += 1) {
    const char = text[index];
    const previous = text[index - 1];
    if (quote !== undefined) {
      if (char === quote && previous !== "\\") {
        quote = undefined;
      }
      continue;
    }
    if (char === "\"" || char === "'" || char === "`") {
      quote = char;
      continue;
    }
    if (char === open) {
      depth += 1;
      continue;
    }
    if (char === close) {
      depth -= 1;
      if (depth === 0) {
        return index;
      }
    }
  }
  return -1;
}
