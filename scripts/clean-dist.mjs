#!/usr/bin/env node
import { rmSync } from "node:fs";
import { join, resolve } from "node:path";
import { dirname } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const distDirs = [
  "apps/backend/dist",
  "apps/excel-addin/dist",
  "apps/mcp-server/dist",
  "packages/cli/dist",
  "packages/excel-core/dist",
  "packages/office-js-engine/dist",
  "packages/protocol/dist",
  "packages/cli/assets"
];

const buildInfoFiles = [
  "apps/backend/tsconfig.tsbuildinfo",
  "apps/excel-addin/tsconfig.tsbuildinfo",
  "apps/mcp-server/tsconfig.tsbuildinfo",
  "packages/cli/tsconfig.tsbuildinfo",
  "packages/excel-core/tsconfig.tsbuildinfo",
  "packages/office-js-engine/tsconfig.tsbuildinfo",
  "packages/protocol/tsconfig.tsbuildinfo"
];

for (const dir of distDirs) {
  rmSync(join(root, dir), { recursive: true, force: true });
}

for (const file of buildInfoFiles) {
  rmSync(join(root, file), { force: true });
}
