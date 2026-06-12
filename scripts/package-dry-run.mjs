#!/usr/bin/env node
import { spawnSync } from "node:child_process";
import { tmpdir } from "node:os";
import { join } from "node:path";

const packages = [
  "./packages/protocol",
  "./packages/excel-core",
  "./packages/office-js-engine",
  "./apps/backend",
  "./apps/mcp-server",
  "./packages/cli"
];

const cacheDir = process.env.npm_config_cache ?? join(tmpdir(), "open-workbook-npm-cache");
let failed = false;

for (const packagePath of packages) {
  const result = spawnSync("npm", ["pack", "--dry-run", packagePath], {
    encoding: "utf8",
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
  if (/[./](?:[^/\n]+)\.(?:test|spec)\.(?:js|d\.ts|map)\b/.test(output)) {
    console.error(`Package dry run for ${packagePath} includes compiled test artifacts.`);
    failed = true;
  }
}

if (failed) {
  process.exit(1);
}
