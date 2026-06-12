import { mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { join, resolve } from "node:path";

if (process.platform !== "darwin") {
  console.error("Mac sideload copies to the Excel for macOS WEF folder and must be run on macOS.");
  process.exit(1);
}

const source = resolve("apps/excel-addin/manifest.xml");
const targetDir = join(homedir(), "Library/Containers/com.microsoft.Excel/Data/Documents/wef");
const target = join(targetDir, "open-workbook.xml");

mkdirSync(targetDir, { recursive: true });
writeFileSync(target, generateManifest(), "utf8");

console.log(`Copied Excel add-in manifest to: ${target}`);
console.log(`Taskpane URL: ${defaultAddinUrl()}`);
console.log(`Backend URL: ${defaultBackendUrl()}`);
console.log("Restart Excel, then insert or open the Open Workbook add-in.");

function generateManifest() {
  const addinUrl = trimTrailingSlash(defaultAddinUrl());
  const backendUrl = defaultBackendUrl();
  const taskpaneUrl = `${addinUrl}/taskpane.html?backendUrl=${encodeURIComponent(backendUrl)}`;
  return readFileSync(source, "utf8")
    .replaceAll("http://localhost:37846/taskpane.html", taskpaneUrl)
    .replaceAll("http://localhost:37846", addinUrl);
}

function defaultAddinUrl() {
  const host = process.env.OPEN_WORKBOOK_ADDIN_HOST ?? "127.0.0.1";
  const port = process.env.OPEN_WORKBOOK_ADDIN_PORT ?? "37846";
  return `http://${host}:${port}`;
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
