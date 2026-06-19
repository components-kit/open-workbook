#!/usr/bin/env node
import { mkdtempSync, rmSync } from "node:fs";
import { spawnSync } from "node:child_process";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";

const packages = [
  "./packages/protocol",
  "./packages/excel-core",
  "./packages/office-js-engine",
  "./apps/backend",
  "./apps/mcp-server",
  "./packages/cli"
];

const cacheDir = process.env.npm_config_cache ?? join(tmpdir(), "open-workbook-npm-cache");
const packDir = mkdtempSync(join(tmpdir(), "open-workbook-pack-"));
let failed = false;

for (const packagePath of packages) {
  const pnpmEntrypoint = process.env.npm_execpath;
  const command = pnpmEntrypoint ? process.execPath : "pnpm";
  const packArgs = ["pack", "--json", "--pack-destination", packDir];
  const args = pnpmEntrypoint ? [pnpmEntrypoint, ...packArgs] : packArgs;
  const result = spawnSync(command, args, {
    encoding: "utf8",
    cwd: resolve(packagePath),
    env: {
      ...process.env,
      npm_config_cache: cacheDir
    }
  });
  const output = `${result.stdout ?? ""}${result.stderr ?? ""}`;
  process.stdout.write(output);
  if (result.status !== 0) {
    failed = true;
    continue;
  }
  const packedEntries = parsePackedEntries(result.stdout);
  const packedFiles = packedEntries.flatMap((entry) => entry.files);
  if (packedFiles.some((file) => /[./](?:[^/\n]+)\.(?:test|spec)\.(?:js|d\.ts|map)\b/.test(file))) {
    console.error(`Package dry run for ${packagePath} includes compiled test artifacts.`);
    failed = true;
  }
  if (packagePath === "./packages/cli" && !packedFiles.includes("assets/instructions/open-workbook-excel/SKILL.md")) {
    console.error("CLI package dry run is missing generic instruction assets.");
    failed = true;
  }
  for (const entry of packedEntries) {
    if (packedPackageJson(entry.filename).includes("workspace:*")) {
      console.error(`Package dry run for ${packagePath} contains unresolved workspace dependencies.`);
      failed = true;
    }
  }
}

rmSync(packDir, { recursive: true, force: true });

if (failed) {
  process.exit(1);
}

function parsePackedEntries(stdout) {
  try {
    const jsonStart = stdout.indexOf("{");
    const jsonEnd = stdout.lastIndexOf("}");
    const parsed = JSON.parse(jsonStart === -1 || jsonEnd === -1 ? stdout : stdout.slice(jsonStart, jsonEnd + 1));
    const entries = Array.isArray(parsed) ? parsed : [parsed];
    return entries.map((entry) => ({
      filename: String(entry.filename ?? ""),
      files: Array.isArray(entry.files) ? entry.files.map((file) => file.path).filter(Boolean) : []
    }));
  } catch {
    return [{
      filename: "",
      files: Array.from(stdout.matchAll(/"path":\s*"([^"]+)"/g), (match) => match[1]).filter(Boolean)
    }];
  }
}

function packedPackageJson(filename) {
  if (!filename) {
    return "";
  }
  const result = spawnSync("tar", ["-xOf", filename, "package/package.json"], { encoding: "utf8" });
  return result.status === 0 ? result.stdout : "";
}
