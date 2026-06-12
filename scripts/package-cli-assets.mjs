import { cpSync, existsSync, mkdirSync, rmSync } from "node:fs";
import { dirname, join, resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(dirname(fileURLToPath(import.meta.url)), "..");
const out = join(root, "packages/cli/assets");

const copies = [
  ["apps/backend/dist", "backend/dist"],
  ["apps/mcp-server/dist", "mcp-server/dist"],
  ["apps/excel-addin/dist", "excel-addin/dist"],
  ["apps/excel-addin/public", "excel-addin/public"],
  ["apps/excel-addin/scripts", "excel-addin/scripts"],
  ["apps/excel-addin/manifest.xml", "excel-addin/manifest.xml"]
];

rmSync(out, { recursive: true, force: true });
mkdirSync(out, { recursive: true });

for (const [from, to] of copies) {
  const source = join(root, from);
  if (!existsSync(source)) {
    throw new Error(`Missing package asset: ${source}. Run corepack pnpm build first.`);
  }
  cpSync(source, join(out, to), { recursive: true });
}

console.log(`Packaged CLI assets in ${out}`);
