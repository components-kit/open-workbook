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
  backend: resolve(repoRoot, "apps/backend/dist/index.js"),
  fileBridge: resolve(repoRoot, "apps/backend/dist/file-bridge.js"),
  addinServer: resolve(repoRoot, "apps/excel-addin/scripts/dev-server.mjs"),
  manifest: resolve(repoRoot, "apps/excel-addin/manifest.xml")
};
const bundledPaths = {
  mcpServer: resolve(packageRoot, "assets/mcp-server/dist/index.js"),
  backend: resolve(packageRoot, "assets/backend/dist/index.js"),
  fileBridge: resolve(packageRoot, "assets/backend/dist/file-bridge.js"),
  addinServer: resolve(packageRoot, "assets/excel-addin/scripts/dev-server.mjs"),
  manifest: resolve(packageRoot, "assets/excel-addin/manifest.xml")
};
const dependencyPaths = {
  mcpServer: resolve(packageRoot, "node_modules/@open-workbook/mcp-server/dist/index.js"),
  backend: resolve(packageRoot, "node_modules/@open-workbook/backend/dist/index.js"),
  fileBridge: resolve(packageRoot, "node_modules/@open-workbook/backend/dist/file-bridge.js"),
  addinServer: bundledPaths.addinServer,
  manifest: bundledPaths.manifest
};

const program = new Command();

program.name("owb").description("Open Workbook local CLI").version("0.1.0");

program
  .command("mcp")
  .description("Start the Open Workbook MCP adapter; attaches to daemon when available")
  .option("--agent-name <name>", "Agent name shown in collaboration status")
  .option("--daemon-url <url>", "Daemon base URL", defaultDaemonUrl())
  .option("--standalone", "Start a single-process MCP server with embedded backend")
  .action((options: { agentName?: string; daemonUrl: string; standalone?: boolean }) => {
    const args: string[] = [];
    if (options.agentName !== undefined) {
      args.push("--agent-name", options.agentName);
    }
    if (options.daemonUrl !== undefined) {
      args.push("--daemon-url", options.daemonUrl);
    }
    if (options.standalone) {
      args.push("--standalone");
    }
    runNode(resolveAsset("mcpServer"), args);
  });

const daemon = program.command("daemon").description("Manage the shared Open Workbook daemon");

daemon
  .command("start")
  .description("Start the shared local Open Workbook daemon")
  .action(() => {
    runNode(resolveAsset("backend"));
  });

daemon
  .command("status")
  .description("Print daemon status from the local health endpoint")
  .option("--daemon-url <url>", "Daemon base URL", defaultDaemonUrl())
  .action(async (options: { daemonUrl: string }) => {
    const response = await fetchLocal(`${trimTrailingSlash(options.daemonUrl)}/status`, "Daemon status", "owb daemon start");
    if (!response.ok) {
      fail(`Daemon status failed: ${response.status} ${response.statusText}`);
    }
    console.log(JSON.stringify(await response.json(), null, 2));
  });

daemon
  .command("stop")
  .description("Stop the shared local Open Workbook daemon")
  .option("--daemon-url <url>", "Daemon base URL", defaultDaemonUrl())
  .action(async (options: { daemonUrl: string }) => {
    const response = await fetchLocal(`${trimTrailingSlash(options.daemonUrl)}/shutdown`, "Daemon stop", "owb daemon start", { method: "POST" });
    if (!response.ok) {
      fail(`Daemon stop failed: ${response.status} ${response.statusText}`);
    }
    console.log("Stopped Open Workbook daemon.");
  });

const fileBridge = program.command("file-bridge").description("Manage the native workbook file bridge");

fileBridge
  .command("start")
  .description("Start the local native file bridge for Save As and host file operations")
  .action(() => {
    runNode(resolveAsset("fileBridge"));
  });

fileBridge
  .command("status")
  .description("Print native file bridge status")
  .option("--bridge-url <url>", "File bridge base URL", defaultFileBridgeUrl())
  .action(async (options: { bridgeUrl: string }) => {
    const response = await fetchLocal(`${trimTrailingSlash(options.bridgeUrl)}/status`, "File bridge status", "owb file-bridge start");
    if (!response.ok) {
      fail(`File bridge status failed: ${response.status} ${response.statusText}`);
    }
    console.log(JSON.stringify(await response.json(), null, 2));
  });

fileBridge
  .command("stop")
  .description("Stop the local native file bridge")
  .option("--bridge-url <url>", "File bridge base URL", defaultFileBridgeUrl())
  .action(async (options: { bridgeUrl: string }) => {
    const response = await fetchLocal(`${trimTrailingSlash(options.bridgeUrl)}/shutdown`, "File bridge stop", "owb file-bridge start", { method: "POST" });
    if (!response.ok) {
      fail(`File bridge stop failed: ${response.status} ${response.statusText}`);
    }
    console.log("Stopped Open Workbook file bridge.");
  });

const service = program.command("service").description("Generate local auto-start service wrappers");

service
  .command("manifest")
  .description("Print or write a launchd, systemd user, or Windows scheduled-task wrapper")
  .option("--target <target>", "Target wrapper: macos, systemd, windows", defaultServiceTarget())
  .option("--service <service>", "Service to run: addin, daemon, or file-bridge", "addin")
  .option("--out <path>", "Write wrapper to a file instead of stdout")
  .option("--command <command>", "Base CLI command to run", defaultServiceCommand())
  .action((options: { target: string; service: string; out?: string; command: string }) => {
    const target = normalizeServiceTarget(options.target);
    const serviceName = normalizeServiceName(options.service);
    const manifest = generateServiceManifest({
      target,
      serviceName,
      command: options.command
    });
    if (options.out) {
      writeFileSync(resolve(options.out), manifest, "utf8");
      console.log(`Wrote ${target} ${serviceName} service wrapper to ${resolve(options.out)}`);
      return;
    }
    console.log(manifest);
  });

program
  .command("addin")
  .description("Manage the local Excel add-in")
  .argument("<command>", "Command: serve")
  .option("--https", "Serve the add-in over HTTPS with a local certificate")
  .option("--tls-cert <path>", "TLS certificate PEM path for HTTPS add-in serving")
  .option("--tls-key <path>", "TLS private key PEM path for HTTPS add-in serving")
  .action((command: string, options: { https?: boolean; tlsCert?: string; tlsKey?: string }) => {
    if (command !== "serve") {
      fail(`Unknown addin command: ${command}`);
    }
    if (options.https) {
      process.env.OPEN_WORKBOOK_ADDIN_HTTPS = "1";
      process.env.OPEN_WORKBOOK_ADDIN_PROTOCOL = "https";
    }
    if (options.tlsCert !== undefined) {
      process.env.OPEN_WORKBOOK_ADDIN_TLS_CERT = resolve(options.tlsCert);
    }
    if (options.tlsKey !== undefined) {
      process.env.OPEN_WORKBOOK_ADDIN_TLS_KEY = resolve(options.tlsKey);
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
  .option("--agent-name <name>", "Agent name passed to the MCP adapter")
  .action((options: { id: string; command: string; agentName?: string }) => {
    const command = [options.command, "mcp"];
    if (options.agentName !== undefined) {
      command.push("--agent-name", options.agentName);
    }
    const config = {
      mcp: {
        [options.id]: {
          type: "local",
          command,
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
          backend: resolveAsset("backend"),
          fileBridge: resolveAsset("fileBridge"),
          addinServer: resolveAsset("addinServer"),
          manifest: resolveAsset("manifest"),
          stateDir: defaultStateDir(),
          exportDir: defaultExportDir(),
          fileBridgeUrl: defaultFileBridgeUrl(),
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
      checkPath("Backend daemon", resolveAsset("backend")),
      checkPath("File bridge", resolveAsset("fileBridge")),
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
    console.log(`File bridge URL: ${defaultFileBridgeUrl()}`);
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

function runNode(entrypoint: string, args: string[] = []): void {
  if (!existsSync(entrypoint)) {
    fail(`Missing built entrypoint: ${entrypoint}\nRun: corepack pnpm build`);
  }

  const child = spawn(process.execPath, [entrypoint, ...args], {
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

type ServiceTarget = "macos" | "systemd" | "windows";
type ServiceName = "addin" | "daemon" | "file-bridge";

function generateServiceManifest(options: { target: ServiceTarget; serviceName: ServiceName; command: string }): string {
  const args = options.serviceName === "addin" ? ["addin", "serve"] : options.serviceName === "daemon" ? ["daemon", "start"] : ["file-bridge", "start"];
  const label = `com.open-workbook.${options.serviceName}`;
  const description =
    options.serviceName === "addin"
      ? "Open Workbook Excel add-in asset server"
      : options.serviceName === "daemon"
        ? "Open Workbook shared daemon"
        : "Open Workbook native file bridge";
  const commandParts = [options.command, ...args];
  switch (options.target) {
    case "macos":
      return generateLaunchdPlist(label, commandParts);
    case "systemd":
      return generateSystemdUnit(label, description, commandParts);
    case "windows":
      return generateWindowsScheduledTask(label, description, commandParts);
  }
}

function generateLaunchdPlist(label: string, commandParts: string[]): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "https://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>${escapeXml(label)}</string>
  <key>ProgramArguments</key>
  <array>
${commandParts.map((part) => `    <string>${escapeXml(part)}</string>`).join("\n")}
  </array>
  <key>RunAtLoad</key>
  <true/>
  <key>KeepAlive</key>
  <true/>
  <key>StandardOutPath</key>
  <string>${escapeXml(join(homedir(), "Library/Logs", `${label}.out.log`))}</string>
  <key>StandardErrorPath</key>
  <string>${escapeXml(join(homedir(), "Library/Logs", `${label}.err.log`))}</string>
</dict>
</plist>
`;
}

function generateSystemdUnit(label: string, description: string, commandParts: string[]): string {
  return `[Unit]
Description=${description}
After=network.target

[Service]
Type=simple
ExecStart=${commandParts.map(shellQuote).join(" ")}
Restart=on-failure
RestartSec=5

[Install]
WantedBy=default.target

# Save as ~/.config/systemd/user/${label}.service
# Enable with: systemctl --user enable --now ${label}.service
`;
}

function generateWindowsScheduledTask(label: string, description: string, commandParts: string[]): string {
  const executable = commandParts[0]!;
  const argumentsText = commandParts.slice(1).join(" ");
  return `$Action = New-ScheduledTaskAction -Execute ${powerShellQuote(executable)} -Argument ${powerShellQuote(argumentsText)}
$Trigger = New-ScheduledTaskTrigger -AtLogOn
$Principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel LeastPrivilege
$Settings = New-ScheduledTaskSettingsSet -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
Register-ScheduledTask -TaskName ${powerShellQuote(label)} -Description ${powerShellQuote(description)} -Action $Action -Trigger $Trigger -Principal $Principal -Settings $Settings -Force
`;
}

function defaultAddinUrl(): string {
  const host = process.env.OPEN_WORKBOOK_ADDIN_HOST ?? "127.0.0.1";
  const port = process.env.OPEN_WORKBOOK_ADDIN_PORT ?? "37846";
  const protocol = process.env.OPEN_WORKBOOK_ADDIN_HTTPS === "1" || process.env.OPEN_WORKBOOK_ADDIN_PROTOCOL === "https" ? "https" : "http";
  return `${protocol}://${host}:${port}`;
}

function defaultBackendUrl(): string {
  const host = process.env.OPEN_WORKBOOK_HOST ?? "127.0.0.1";
  const port = process.env.OPEN_WORKBOOK_PORT ?? "37845";
  const path = process.env.OPEN_WORKBOOK_ADDIN_PATH ?? "/addin";
  return `ws://${host}:${port}${path}`;
}

function defaultDaemonUrl(): string {
  const host = process.env.OPEN_WORKBOOK_HOST ?? "127.0.0.1";
  const port = process.env.OPEN_WORKBOOK_PORT ?? "37845";
  return `http://${host}:${port}`;
}

function defaultFileBridgeUrl(): string {
  if (process.env.OPEN_WORKBOOK_FILE_BRIDGE_URL !== undefined) {
    return process.env.OPEN_WORKBOOK_FILE_BRIDGE_URL;
  }
  const host = process.env.OPEN_WORKBOOK_FILE_BRIDGE_HOST ?? "127.0.0.1";
  const port = process.env.OPEN_WORKBOOK_FILE_BRIDGE_PORT ?? "37847";
  return `http://${host}:${port}`;
}

function defaultServiceTarget(): ServiceTarget {
  if (process.platform === "darwin") {
    return "macos";
  }
  if (process.platform === "win32") {
    return "windows";
  }
  return "systemd";
}

function defaultServiceCommand(): string {
  return process.env.OPEN_WORKBOOK_SERVICE_COMMAND ?? "owb";
}

function defaultStateDir(): string {
  return process.env.OPEN_WORKBOOK_STATE_DIR ?? resolve(process.cwd(), ".open-workbook/state");
}

function defaultExportDir(): string {
  return process.env.OPEN_WORKBOOK_EXPORT_DIR ?? resolve(process.cwd(), ".open-workbook/exports");
}

function trimTrailingSlash(value: string): string {
  return value.endsWith("/") ? value.slice(0, -1) : value;
}

function normalizeServiceTarget(value: string): ServiceTarget {
  if (value === "mac" || value === "macos" || value === "launchd") {
    return "macos";
  }
  if (value === "win" || value === "windows" || value === "task-scheduler") {
    return "windows";
  }
  if (value === "linux" || value === "systemd") {
    return "systemd";
  }
  fail(`Unknown service target: ${value}`);
}

function normalizeServiceName(value: string): ServiceName {
  if (value === "addin" || value === "daemon" || value === "file-bridge") {
    return value;
  }
  fail(`Unknown service name: ${value}`);
}

function escapeXml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;")
    .replaceAll("'", "&apos;");
}

function shellQuote(value: string): string {
  return `'${value.replaceAll("'", "'\\''")}'`;
}

function powerShellQuote(value: string): string {
  return `'${value.replaceAll("'", "''")}'`;
}

function checkPath(label: string, path: string): { label: string; path: string; ok: boolean } {
  return { label, path, ok: existsSync(path) };
}

async function fetchLocal(url: string, label: string, startCommand: string, init?: RequestInit): Promise<Response> {
  try {
    return await fetch(url, init);
  } catch (error) {
    const detail = error instanceof Error ? error.message : String(error);
    fail(`${label} failed: could not connect to ${url}. Start it with \`${startCommand}\` or check the configured port. ${detail}`);
  }
}

function fail(message: string): never {
  console.error(message);
  process.exit(1);
}
