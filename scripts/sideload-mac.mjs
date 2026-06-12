import { copyFileSync, mkdirSync } from "node:fs";
import { homedir } from "node:os";
import { join, resolve } from "node:path";

const source = resolve("apps/excel-addin/manifest.xml");
const targetDir = join(homedir(), "Library/Containers/com.microsoft.Excel/Data/Documents/wef");
const target = join(targetDir, "open-workbook.xml");

mkdirSync(targetDir, { recursive: true });
copyFileSync(source, target);

console.log(`Copied Excel add-in manifest to: ${target}`);
console.log("Restart Excel, then insert or open the Open Workbook add-in.");
