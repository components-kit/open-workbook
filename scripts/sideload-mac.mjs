import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { join, resolve } from "node:path";

if (process.platform !== "darwin") {
  console.error("Mac sideload copies to the Excel for macOS WEF folder and must be run on macOS.");
  process.exit(1);
}

const source = resolve("apps/excel-addin/manifest.xml");
const targetDir = join(homedir(), "Library/Containers/com.microsoft.Excel/Data/Documents/wef");
const target = join(targetDir, "open-workbook-local.xml");
const developmentManifestId = "6f2d2ac1-69b0-4eb6-a256-0a1fcb00d3e1";

mkdirSync(targetDir, { recursive: true });
writeFileSync(target, generateManifest(), "utf8");

console.log(`Copied Excel add-in manifest to: ${target}`);
console.log("Manifest variant: development");
console.log(`Taskpane URL: ${defaultAddinUrl()}`);
console.log(`Backend URL: ${defaultBackendUrl()}`);
console.log("Restart Excel, then insert or open the OpenWorkbook Local add-in.");

function generateManifest() {
  const addinUrl = trimTrailingSlash(defaultAddinUrl());
  const backendUrl = defaultBackendUrl();
  const taskpaneUrl = `${addinUrl}/taskpane.html?backendUrl=${encodeURIComponent(backendUrl)}`;
  return applyDevelopmentManifestIdentity(readFileSync(source, "utf8"))
    .replaceAll("http://localhost:37846/taskpane.html", taskpaneUrl)
    .replaceAll("http://localhost:37846", addinUrl);
}

function applyDevelopmentManifestIdentity(manifest) {
  return manifest
    .replace(/<Id>[^<]+<\/Id>/, `<Id>${developmentManifestId}</Id>`)
    .replaceAll("OpenWorkbook.Group.ComponentsKit", "OpenWorkbookLocal.Group.ComponentsKit")
    .replaceAll("OpenWorkbook.TaskpaneButton.ComponentsKit", "OpenWorkbookLocal.TaskpaneButton.ComponentsKit")
    .replaceAll("OpenWorkbook.Taskpane.ComponentsKit", "OpenWorkbookLocal.Taskpane.ComponentsKit")
    .replaceAll("OpenWorkbook.", "OpenWorkbookLocal.")
    .replaceAll('DisplayName DefaultValue="OpenWorkbook"', 'DisplayName DefaultValue="OpenWorkbook Local"')
    .replaceAll('DefaultValue="OpenWorkbook loaded"', 'DefaultValue="OpenWorkbook Local loaded"')
    .replaceAll('DefaultValue="OpenWorkbook"', 'DefaultValue="OpenWorkbook Local"')
    .replaceAll("OpenWorkbook connects Excel", "OpenWorkbook Local connects Excel")
    .replaceAll("Open the OpenWorkbook taskpane.", "Open the OpenWorkbook Local taskpane.");
}

function defaultAddinUrl() {
  const host = process.env.OPEN_WORKBOOK_ADDIN_HOST ?? "127.0.0.1";
  const port = process.env.OPEN_WORKBOOK_ADDIN_PORT ?? "37846";
  const protocol = process.env.OPEN_WORKBOOK_ADDIN_HTTPS === "1" || process.env.OPEN_WORKBOOK_ADDIN_PROTOCOL === "https" ? "https" : "http";
  return `${protocol}://${host}:${port}`;
}

function defaultBackendUrl() {
  const host = process.env.OPEN_WORKBOOK_HOST ?? "127.0.0.1";
  const port = process.env.OPEN_WORKBOOK_PORT ?? "37845";
  const path = process.env.OPEN_WORKBOOK_ADDIN_PATH ?? "/addin";
  return `ws://${host}:${port}${path}`;
}

function trimTrailingSlash(value) {
  return value.endsWith("/") ? value.slice(0, -1) : value;
}
