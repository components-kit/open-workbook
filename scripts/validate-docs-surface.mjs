#!/usr/bin/env node
import { readFileSync } from "node:fs";
import { getExposedToolCatalog } from "../packages/protocol/dist/tools.js";

const toolSurface = readFileSync(new URL("../docs/tool-surface.md", import.meta.url), "utf8");
const exposedTools = getExposedToolCatalog({ includePreview: true }).map((tool) => tool.name).sort();
const missing = exposedTools.filter((name) => !toolSurface.includes(`\`${name}\``));

if (missing.length > 0) {
  console.error("Tool surface docs are missing exposed callable tools.");
  console.error(missing.map((name) => `- ${name}`).join("\n"));
  process.exit(1);
}

console.log(`Docs surface check passed: ${exposedTools.length} callable tool(s) documented.`);
