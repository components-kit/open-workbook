#!/usr/bin/env node
import { Command } from "commander";
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { spawn } from "node:child_process";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(__dirname, "../../..");
const packageRoot = resolve(__dirname, "..");
const sourcePaths = {
  mcpServer: resolve(repoRoot, "apps/mcp-server/dist/index.js"),
  addinServer: resolve(repoRoot, "apps/excel-addin/scripts/dev-server.mjs"),
  manifest: resolve(repoRoot, "apps/excel-addin/manifest.xml")
};
const bundledPaths = {
  mcpServer: resolve(packageRoot, "assets/mcp-server/dist/index.js"),
  addinServer: resolve(packageRoot, "assets/excel-addin/scripts/dev-server.mjs"),
  manifest: resolve(packageRoot, "assets/excel-addin/manifest.xml")
};
const dependencyPaths = {
  mcpServer: resolve(packageRoot, "node_modules/@open-workbook/mcp-server/dist/index.js"),
  addinServer: bundledPaths.addinServer,
  manifest: bundledPaths.manifest
};

const program = new Command();

program.name("owb").description("Open Workbook local CLI").version("0.1.0");

program
  .command("mcp")
  .description("Start the Open Workbook MCP server and embedded add-in backend")
  .action(() => {
    runNode(resolveAsset("mcpServer"));
  });

program
  .command("addin")
  .description("Manage the local Excel add-in")
  .argument("<command>", "Command: serve")
  .action((command: string) => {
    if (command !== "serve") {
      fail(`Unknown addin command: ${command}`);
    }
    runNode(resolveAsset("addinServer"));
  });

const sideload = program.command("sideload").description("Sideload the Excel add-in manifest");

sideload
  .command("mac")
  .description("Copy manifest to the Excel for macOS sideload folder")
  .option("--addin-url <url>", "Taskpane base URL", defaultAddinUrl())
  .option("--backend-url <url>", "Backend WebSocket URL", defaultBackendUrl())
  .action((options: { addinUrl: string; backendUrl: string }) => {
    if (process.platform !== "darwin") {
      fail("Mac sideload copies to the Excel for macOS WEF folder and must be run on macOS. Use `owb sideload manifest --out open-workbook.xml` on other platforms.");
    }
    const targetDir = join(homedir(), "Library/Containers/com.microsoft.Excel/Data/Documents/wef");
    const target = join(targetDir, "open-workbook.xml");
    mkdirSync(targetDir, { recursive: true });
    writeFileSync(target, generateManifest(options), "utf8");
    console.log(`Copied manifest to ${target}`);
    console.log(`Taskpane URL: ${options.addinUrl}`);
    console.log(`Backend URL: ${options.backendUrl}`);
    console.log("Restart Excel, then open the Open Workbook add-in.");
  });

sideload
  .command("windows")
  .description("Print Windows trusted catalog sideload instructions")
  .option("--out <path>", "Write generated manifest to this path", "open-workbook.xml")
  .option("--addin-url <url>", "Taskpane base URL", defaultAddinUrl())
  .option("--backend-url <url>", "Backend WebSocket URL", defaultBackendUrl())
  .action((options: { out: string; addinUrl: string; backendUrl: string }) => {
    const manifest = generateManifest(options);
    const outputPath = resolve(options.out);
    writeFileSync(outputPath, manifest, "utf8");
    console.log(`Wrote manifest to ${outputPath}`);
    console.log("Windows Excel sideloading uses a trusted shared-folder add-in catalog.");
    console.log("");
    console.log("1. Create a folder such as C:\\open-workbook-addins.");
    console.log("2. Share that folder in Windows and note its UNC path, for example \\\\YOUR-PC\\open-workbook-addins.");
    console.log(`3. Copy ${outputPath} into the shared folder.`);
    console.log("4. In Excel: File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs.");
    console.log("5. Add the UNC shared-folder path as a trusted catalog and select Show in Menu.");
    console.log("6. Restart Excel and insert Open Workbook from Shared Folder.");
  });

sideload
  .command("manifest")
  .description("Print or write a generated Excel add-in manifest")
  .option("--out <path>", "Write manifest to a file instead of stdout")
  .option("--addin-url <url>", "Taskpane base URL", defaultAddinUrl())
  .option("--backend-url <url>", "Backend WebSocket URL", defaultBackendUrl())
  .action((options: { out?: string; addinUrl: string; backendUrl: string }) => {
    const manifest = generateManifest(options);
    if (options.out) {
      writeFileSync(resolve(options.out), manifest, "utf8");
      console.log(`Wrote manifest to ${resolve(options.out)}`);
      return;
    }
    console.log(manifest);
  });

const opencode = program.command("opencode").description("Generate OpenCode integration snippets");

opencode
  .command("config")
  .description("Print an OpenCode MCP server config snippet")
  .option("--id <id>", "MCP server id in OpenCode config", "open-workbook")
  .option("--command <command>", "Command to run Open Workbook MCP", "owb")
  .action((options: { id: string; command: string }) => {
    const config = {
      mcp: {
        [options.id]: {
          type: "local",
          command: [options.command, "mcp"],
          enabled: true
        }
      }
    };
    console.log(JSON.stringify(config, null, 2));
  });

program
  .command("paths")
  .description("Print resolved package paths")
  .action(() => {
    console.log(
      JSON.stringify(
        {
          mcpServer: resolveAsset("mcpServer"),
          addinServer: resolveAsset("addinServer"),
          manifest: resolveAsset("manifest"),
          mode: existsSync(sourcePaths.mcpServer) ? "source" : "bundled"
        },
        null,
        2
      )
    );
  });

program
  .command("doctor")
  .description("Check local Open Workbook install assets")
  .action(() => {
    const checks = [
      checkPath("MCP server", resolveAsset("mcpServer")),
      checkPath("Add-in server", resolveAsset("addinServer")),
      checkPath("Manifest", resolveAsset("manifest"))
    ];
    for (const check of checks) {
      console.log(`${check.ok ? "ok" : "missing"} ${check.label}: ${check.path}`);
    }
    if (checks.some((check) => !check.ok)) {
      process.exitCode = 1;
      return;
    }
    console.log(`Taskpane URL: ${defaultAddinUrl()}`);
    console.log(`Backend URL: ${defaultBackendUrl()}`);
  });

program.parse();

type AssetName = keyof typeof sourcePaths;

function resolveAsset(name: AssetName): string {
  if (existsSync(sourcePaths[name])) {
    return sourcePaths[name];
  }
  if (existsSync(bundledPaths[name])) {
    return bundledPaths[name];
  }
  if (existsSync(dependencyPaths[name])) {
    return dependencyPaths[name];
  }
  return sourcePaths[name];
}

function runNode(entrypoint: string): void {
  if (!existsSync(entrypoint)) {
    fail(`Missing built entrypoint: ${entrypoint}\nRun: corepack pnpm build`);
  }

  const child = spawn(process.execPath, [entrypoint], {
    stdio: "inherit",
    env: process.env
  });

  child.on("exit", (code, signal) => {
    if (signal) {
      process.kill(process.pid, signal);
    }
    process.exit(code ?? 0);
  });
}

function generateManifest(options: { addinUrl: string; backendUrl: string }): string {
  const manifestPath = resolveAsset("manifest");
  if (!existsSync(manifestPath)) {
    fail(`Missing manifest: ${manifestPath}`);
  }
  const addinUrl = trimTrailingSlash(options.addinUrl);
  const taskpaneUrl = `${addinUrl}/taskpane.html?backendUrl=${encodeURIComponent(options.backendUrl)}`;
  return readFileSync(manifestPath, "utf8")
    .replaceAll("http://localhost:37846/taskpane.html", taskpaneUrl)
    .replaceAll("http://localhost:37846", addinUrl);
}

function defaultAddinUrl(): string {
  const host = process.env.OPEN_WORKBOOK_ADDIN_HOST ?? "127.0.0.1";
  const port = process.env.OPEN_WORKBOOK_ADDIN_PORT ?? "37846";
  return `http://${host}:${port}`;
}

function defaultBackendUrl(): string {
  const host = process.env.OPEN_WORKBOOK_HOST ?? "127.0.0.1";
  const port = process.env.OPEN_WORKBOOK_PORT ?? "37845";
  const path = process.env.OPEN_WORKBOOK_ADDIN_PATH ?? "/addin";
  return `ws://${host}:${port}${path}`;
}

function trimTrailingSlash(value: string): string {
  return value.endsWith("/") ? value.slice(0, -1) : value;
}

function checkPath(label: string, path: string): { label: string; path: string; ok: boolean } {
  return { label, path, ok: existsSync(path) };
}

function fail(message: string): never {
  console.error(message);
  process.exit(1);
}
