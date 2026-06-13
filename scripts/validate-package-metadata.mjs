#!/usr/bin/env node
import { existsSync, readFileSync } from "node:fs";
import { join } from "node:path";

const rootPackage = readJson("package.json");
const rootVersion = requiredString(rootPackage.version, "package.json version");
const repositoryUrl = requiredString(rootPackage.repository?.url, "package.json repository.url");
const bugsUrl = requiredString(rootPackage.bugs?.url, "package.json bugs.url");
const homepage = requiredString(rootPackage.homepage, "package.json homepage");

const publishablePackages = [
  "packages/protocol",
  "packages/excel-core",
  "packages/office-js-engine",
  "apps/backend",
  "apps/mcp-server",
  "packages/cli"
];
const expectedPackageNames = new Map([
  ["packages/protocol", "@components-kit/open-workbook-protocol"],
  ["packages/excel-core", "@components-kit/open-workbook-excel-core"],
  ["packages/office-js-engine", "@components-kit/open-workbook-office-js-engine"],
  ["apps/backend", "@components-kit/open-workbook-backend"],
  ["apps/mcp-server", "@components-kit/open-workbook-mcp-server"],
  ["packages/cli", "@components-kit/open-workbook"]
]);
const privatePackages = ["apps/excel-addin"];
const expectedPrivatePackageNames = new Map([
  ["apps/excel-addin", "@components-kit/open-workbook-excel-addin"]
]);

const errors = [];

for (const packageDir of publishablePackages) {
  const packageJsonPath = join(packageDir, "package.json");
  const manifest = readJson(packageJsonPath);
  requireField(manifest.name, `${packageJsonPath} name`);
  expect(manifest.name === expectedPackageNames.get(packageDir), `${packageJsonPath} name must be ${expectedPackageNames.get(packageDir)}`);
  requireField(manifest.description, `${packageJsonPath} description`);
  requireField(manifest.license, `${packageJsonPath} license`);
  requireField(manifest.type, `${packageJsonPath} type`);
  requireField(manifest.repository?.directory, `${packageJsonPath} repository.directory`);
  requireField(manifest.files, `${packageJsonPath} files`);
  requireField(manifest.publishConfig?.access, `${packageJsonPath} publishConfig.access`);
  expect(manifest.version === rootVersion, `${packageJsonPath} version must match root ${rootVersion}`);
  expect(manifest.license === rootPackage.license, `${packageJsonPath} license must match root ${rootPackage.license}`);
  expect(manifest.type === "module", `${packageJsonPath} type must be module`);
  expect(manifest.homepage === homepage, `${packageJsonPath} homepage must match root`);
  expect(manifest.repository?.url === repositoryUrl, `${packageJsonPath} repository.url must match root`);
  expect(manifest.bugs?.url === bugsUrl, `${packageJsonPath} bugs.url must match root`);
  expect(manifest.repository?.directory === packageDir, `${packageJsonPath} repository.directory must be ${packageDir}`);
  expect(manifest.publishConfig?.access === "public", `${packageJsonPath} publishConfig.access must be public`);
  expect(manifest.private !== true, `${packageJsonPath} must not be private`);
  expect(Array.isArray(manifest.files) && manifest.files.includes("dist"), `${packageJsonPath} files must include dist`);
  expect(Array.isArray(manifest.files) && manifest.files.includes("README.md"), `${packageJsonPath} files must include README.md`);
  expect(existsSync(join(packageDir, "README.md")), `${packageDir}/README.md must exist`);

  if (manifest.bin !== undefined) {
    for (const [binName, binPath] of Object.entries(manifest.bin)) {
      expect(typeof binPath === "string" && binPath.startsWith("dist/"), `${packageJsonPath} bin ${binName} must point into dist`);
    }
  }
  if (manifest.exports !== undefined) {
    for (const [exportName, exportValue] of Object.entries(manifest.exports)) {
      const value = exportValue;
      expect(Boolean(value?.types), `${packageJsonPath} export ${exportName} must declare types`);
      expect(Boolean(value?.default), `${packageJsonPath} export ${exportName} must declare default`);
      expect(String(value?.types).startsWith("./dist/"), `${packageJsonPath} export ${exportName} types must point into dist`);
      expect(String(value?.default).startsWith("./dist/"), `${packageJsonPath} export ${exportName} default must point into dist`);
    }
  } else {
    expect(typeof manifest.main === "string" && manifest.main.includes("dist/"), `${packageJsonPath} must declare main in dist when exports is absent`);
    expect(typeof manifest.types === "string" && manifest.types.includes("dist/"), `${packageJsonPath} must declare types in dist when exports is absent`);
  }
}

for (const packageDir of privatePackages) {
  const packageJsonPath = join(packageDir, "package.json");
  const manifest = readJson(packageJsonPath);
  expect(manifest.name === expectedPrivatePackageNames.get(packageDir), `${packageJsonPath} name must be ${expectedPrivatePackageNames.get(packageDir)}`);
  expect(manifest.version === rootVersion, `${packageJsonPath} version must match root ${rootVersion}`);
  expect(manifest.private === true, `${packageJsonPath} must remain private`);
  expect(manifest.publishConfig === undefined, `${packageJsonPath} must not declare publishConfig`);
  expect(existsSync(join(packageDir, "README.md")), `${packageDir}/README.md must exist`);
}

if (errors.length > 0) {
  console.error("Package metadata validation failed.");
  console.error(errors.map((error) => `- ${error}`).join("\n"));
  process.exit(1);
}

console.log(`Package metadata check passed: ${publishablePackages.length} publishable package(s), ${privatePackages.length} private package(s).`);

function readJson(filePath) {
  return JSON.parse(readFileSync(filePath, "utf8"));
}

function requireField(value, label) {
  expect(value !== undefined && value !== null && value !== "", `${label} is required`);
}

function requiredString(value, label) {
  requireField(value, label);
  return String(value);
}

function expect(condition, message) {
  if (!condition) {
    errors.push(message);
  }
}
