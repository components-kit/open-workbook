#!/usr/bin/env node
import { readFileSync } from "node:fs";
import { getExposedToolCatalog } from "../packages/protocol/dist/tools.js";

const source = readFileSync(new URL("../apps/mcp-server/src/index.ts", import.meta.url), "utf8");

const callableTools = getExposedToolCatalog({ includePreview: true }).map((tool) => tool.name).sort();
const registeredTools = [...collectRegisteredToolNames(source)].sort();

const missing = callableTools.filter((name) => !registeredTools.includes(name));
const extra = registeredTools.filter((name) => !callableTools.includes(name));

if (missing.length > 0 || extra.length > 0) {
  console.error("MCP tool catalog drift detected.");
  if (missing.length > 0) {
    console.error(`Missing MCP registrations:\n${missing.map((name) => `- ${name}`).join("\n")}`);
  }
  if (extra.length > 0) {
    console.error(`Unexpected MCP registrations:\n${extra.map((name) => `- ${name}`).join("\n")}`);
  }
  process.exit(1);
}

console.log(`MCP tool catalog check passed: ${callableTools.length} callable tools registered.`);

function collectRegisteredToolNames(text) {
  const names = new Set();
  for (const match of text.matchAll(/\bregister(?:McpTool|RangeOperation|SheetOperation)\(\s*mcp,\s*"([^"]+)"/g)) {
    names.add(match[1]);
  }
  for (const loop of findRegistrationLoops(text)) {
    for (const name of loop.names) {
      names.add(name);
    }
  }
  for (const name of parseObjectKeys(text, "STYLE_COPY_TOOL_DIMENSIONS")) {
    names.add(name);
  }
  return names;
}

function findRegistrationLoops(text) {
  const loops = [];
  const loopPattern = /for \(const (?:name|\[name,\s*[^\]]+\]) of \[/g;
  for (const match of text.matchAll(loopPattern)) {
    const arrayStart = match.index + match[0].length - 1;
    const arrayEnd = findMatching(text, arrayStart, "[", "]");
    if (arrayEnd < 0) {
      continue;
    }
    const blockStart = text.indexOf("{", arrayEnd);
    if (blockStart < 0) {
      continue;
    }
    const blockEnd = findMatching(text, blockStart, "{", "}");
    if (blockEnd < 0) {
      continue;
    }
    const body = text.slice(blockStart, blockEnd + 1);
    if (!/\bregister(?:McpTool|RangeOperation|SheetOperation)\(\s*mcp,\s*name\b/.test(body)) {
      continue;
    }
    loops.push({
      names: [...text.slice(arrayStart, arrayEnd + 1).matchAll(/"([^"]+)"/g)]
        .map((nameMatch) => nameMatch[1])
        .filter((name) => name.startsWith("excel."))
    });
  }
  return loops;
}

function parseObjectKeys(text, constName) {
  const declaration = text.match(new RegExp(`const ${constName}:[^{]+{`));
  if (!declaration?.index) {
    return [];
  }
  const objectStart = text.indexOf("{", declaration.index);
  const objectEnd = findMatching(text, objectStart, "{", "}");
  if (objectEnd < 0) {
    return [];
  }
  return [...text.slice(objectStart, objectEnd + 1).matchAll(/"([^"]+)":/g)]
    .map((match) => match[1])
    .filter((name) => name.startsWith("excel."));
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
