import { existsSync, readFileSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

export function findRepoRoot(importMetaUrl) {
  let current = dirname(fileURLToPath(importMetaUrl));
  while (current !== dirname(current)) {
    const packageJsonPath = join(current, "package.json");
    if (existsSync(packageJsonPath)) {
      try {
        const manifest = JSON.parse(readFileSync(packageJsonPath, "utf8"));
        if (manifest.name === "open-workbook" && manifest.private === true) {
          return current;
        }
      } catch {
        // Keep walking; malformed package files should not hide the repo root.
      }
    }
    current = dirname(current);
  }
  return resolve(dirname(fileURLToPath(importMetaUrl)), "../..");
}
