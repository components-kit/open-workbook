import { readFileSync, writeFileSync } from "node:fs";
import { resolve } from "node:path";

const source = resolve("apps/excel-addin/manifest.xml");
const outIndex = process.argv.indexOf("--out");
const outputPath = resolve(outIndex >= 0 && process.argv[outIndex + 1] ? process.argv[outIndex + 1] : "open-workbook.xml");

writeFileSync(outputPath, generateManifest(), "utf8");

console.log(`Wrote manifest to ${outputPath}`);
console.log("Windows Excel sideloading uses a trusted shared-folder add-in catalog.");
console.log("");
console.log("Recommended steps:");
console.log("1. Create a folder such as C:\\open-workbook-addins.");
console.log("2. Share that folder in Windows and note its UNC path, for example \\\\YOUR-PC\\open-workbook-addins.");
console.log(`3. Copy ${outputPath} into the shared folder.`);
console.log("4. In Excel: File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs.");
console.log("5. Add the UNC shared-folder path as a trusted catalog and select Show in Menu.");
console.log("6. Restart Excel and insert the Open Workbook add-in from Shared Folder.");

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
