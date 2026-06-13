#!/usr/bin/env node
import { spawnSync } from "node:child_process";

const packages = [
  "@component-kit/open-workbook-protocol",
  "@component-kit/open-workbook-excel-core",
  "@component-kit/open-workbook-office-js-engine",
  "@component-kit/open-workbook-backend",
  "@component-kit/open-workbook-mcp-server",
  "@component-kit/open-workbook"
];

const extraArgs = process.argv.slice(2);
if (extraArgs.includes("--help") || extraArgs.includes("-h")) {
  console.log(`Usage: corepack pnpm publish:npm -- [pnpm publish options]

Publishes Open Workbook public npm packages in dependency order:
${packages.map((packageName) => `- ${packageName}`).join("\n")}

Examples:
  corepack pnpm publish:npm
  corepack pnpm publish:npm -- --otp 123456
  corepack pnpm publish:npm -- --dry-run
`);
  process.exit(0);
}

const pnpmEntrypoint = process.env.npm_execpath;
const command = pnpmEntrypoint ? process.execPath : "pnpm";
let failed = false;

for (const packageName of packages) {
  const publishArgs = [
    "--filter",
    packageName,
    "publish",
    "--access",
    "public",
    ...extraArgs
  ];
  const args = pnpmEntrypoint ? [pnpmEntrypoint, ...publishArgs] : publishArgs;
  console.log(`\nPublishing ${packageName}...`);
  const result = spawnSync(command, args, {
    stdio: "inherit",
    env: process.env
  });
  if (result.status !== 0) {
    failed = true;
    console.error(`Publishing ${packageName} failed.`);
    break;
  }
}

if (failed) {
  process.exit(1);
}

console.log("\nPublished Open Workbook npm packages.");
