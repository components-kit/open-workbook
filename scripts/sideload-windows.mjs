import { resolve } from "node:path";

const manifest = resolve("apps/excel-addin/manifest.xml");

console.log("Windows Excel sideloading uses a trusted add-in catalog folder.");
console.log(`Manifest to copy: ${manifest}`);
console.log("");
console.log("Recommended steps:");
console.log("1. Create a folder such as C:\\open-workbook-addins.");
console.log("2. Copy manifest.xml into that folder.");
console.log("3. In Excel: File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs.");
console.log("4. Add the folder path as a trusted catalog and enable it in the menu.");
console.log("5. Restart Excel and insert the Open Workbook add-in from Shared Folder.");
