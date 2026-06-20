#!/usr/bin/env node
import { readFileSync } from "node:fs";
import { getInternalCapabilityCatalog } from "../../packages/protocol/dist/tools.js";

const toolSurface = readFileSync(new URL("../../docs/tool-surface.md", import.meta.url), "utf8");
const skill = readFileSync(new URL("../../skills/open-workbook-skills/SKILL.md", import.meta.url), "utf8");
const packagedSkill = readFileSync(new URL("../../packages/cli/assets/instructions/open-workbook-skills/SKILL.md", import.meta.url), "utf8");
const catalogTools = getInternalCapabilityCatalog({ includePreview: true }).map((tool) => tool.name).sort();
const missing = catalogTools.filter((name) => !toolSurface.includes(`\`${name}\``));

if (missing.length > 0) {
  console.error("Tool surface docs are missing stable or preview catalog tools.");
  console.error(missing.map((name) => `- ${name}`).join("\n"));
  process.exit(1);
}

const forbiddenDefaultSkillGuidance = [
  "Start with compact context tools",
  "Prefer `excel.range.read_compact` and `excel.table.read_compact` for exploratory data reads",
  "OPEN_WORKBOOK_MCP_SURFACE=advanced",
  "advanced compatibility/debug surface"
];
const skillSources = [
  ["docs/tool-surface.md", toolSurface],
  ["skills/open-workbook-skills/SKILL.md", skill],
  ["packages/cli/assets/instructions/open-workbook-skills/SKILL.md", packagedSkill]
];
const forbiddenMatches = skillSources.flatMap(([name, source]) =>
  forbiddenDefaultSkillGuidance
    .filter((phrase) => source.includes(phrase))
    .map((phrase) => `${name}: ${phrase}`)
);
if (forbiddenMatches.length > 0) {
  console.error("Default skill guidance should route normal agents through excel.agent.run, not compact primitive chains.");
  console.error(forbiddenMatches.map((match) => `- ${match}`).join("\n"));
  process.exit(1);
}

console.log(`Docs surface check passed: ${catalogTools.length} stable/preview internal capability(s) documented.`);
