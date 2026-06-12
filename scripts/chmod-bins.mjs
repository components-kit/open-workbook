import { chmodSync, existsSync } from "node:fs";
import { resolve } from "node:path";

for (const file of ["apps/mcp-server/dist/index.js", "packages/cli/dist/index.js"]) {
  const path = resolve(file);
  if (existsSync(path)) {
    chmodSync(path, 0o755);
  }
}
