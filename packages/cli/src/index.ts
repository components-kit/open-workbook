#!/usr/bin/env node
import { Command } from "commander";
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { homedir } from "node:os";
import { dirname, join, resolve } from "node:path";
import { spawn, type ChildProcess } from "node:child_process";
import { fileURLToPath } from "node:url";

const __dirname = dirname(fileURLToPath(import.meta.url));
const repoRoot = resolve(__dirname, "../../..");
const packageRoot = resolve(__dirname, "..");
const publicPackageName = "@components-kit/open-workbook";
const currentVersion = readPackageVersion([join(packageRoot, "package.json"), join(repoRoot, "package.json")]) ?? "0.1.3";
const instructionFileName = "open-workbook-excel.md";
const sourcePaths = {
  mcpServer: resolve(repoRoot, "apps/mcp-server/dist/index.js"),
  backend: resolve(repoRoot, "apps/backend/dist/index.js"),
  fileBridge: resolve(repoRoot, "apps/backend/dist/file-bridge.js"),
  addinServer: resolve(repoRoot, "apps/excel-addin/scripts/dev-server.mjs"),
  manifest: resolve(repoRoot, "apps/excel-addin/manifest.xml"),
  instructions: resolve(repoRoot, "skills/open-workbook-excel")
};
const bundledPaths = {
  mcpServer: resolve(packageRoot, "assets/mcp-server/dist/index.js"),
  backend: resolve(packageRoot, "assets/backend/dist/index.js"),
  fileBridge: resolve(packageRoot, "assets/backend/dist/file-bridge.js"),
  addinServer: resolve(packageRoot, "assets/excel-addin/scripts/dev-server.mjs"),
  manifest: resolve(packageRoot, "assets/excel-addin/manifest.xml"),
  instructions: resolve(packageRoot, "assets/instructions/open-workbook-excel")
};
const dependencyPaths = {
  mcpServer: resolve(packageRoot, "node_modules/@components-kit/open-workbook-mcp-server/dist/index.js"),
  backend: resolve(packageRoot, "node_modules/@components-kit/open-workbook-backend/dist/index.js"),
  fileBridge: resolve(packageRoot, "node_modules/@components-kit/open-workbook-backend/dist/file-bridge.js"),
  addinServer: bundledPaths.addinServer,
  manifest: bundledPaths.manifest,
  instructions: bundledPaths.instructions
};

const program = new Command();

program.name("owb").description("Open Workbook local CLI").version(currentVersion);

program
  .command("mcp")
  .description("Start the Open Workbook MCP adapter and local Excel add-in asset server")
  .option("--agent-name <name>", "Agent name shown in collaboration status")
  .option("--daemon-url <url>", "Daemon base URL", defaultDaemonUrl())
  .option("--standalone", "Start a single-process MCP server with embedded backend")
  .option("--no-addin-server", "Do not start the companion local Excel add-in asset server")
  .action(async (options: { agentName?: string; daemonUrl: string; standalone?: boolean; addinServer?: boolean }) => {
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
    const companionProcesses: ChildProcess[] = [];
    if (options.addinServer !== false) {
      const addinServer = await startAddinServerIfNeeded();
      if (addinServer) {
        companionProcesses.push(addinServer);
      }
    }
    runNode(resolveAsset("mcpServer"), args, companionProcesses);
  });

program
  .command("setup")
  .description("Initialize Open Workbook and prepare the Excel add-in manifest")
  .option("--dry-run", "Print setup actions without writing files")
  .option("--instructions-out <path>", "Instruction file path", defaultInstructionsPath())
  .option("--manifest-out <path>", "Manifest path for non-macOS setup", defaultSetupManifestPath())
  .option("--addin-url <url>", "Taskpane base URL", defaultAddinUrl())
  .option("--backend-url <url>", "Backend WebSocket URL", defaultBackendUrl())
  .action((options: SetupOptions) => runSetup(options, "setup"));

program
  .command("upgrade")
  .description("Upgrade local Open Workbook setup assets and print the current MCP config")
  .option("--dry-run", "Print upgrade actions without writing files")
  .option("--instructions-out <path>", "Instruction file path", defaultInstructionsPath())
  .option("--manifest-out <path>", "Manifest path for non-macOS setup", defaultSetupManifestPath())
  .option("--addin-url <url>", "Taskpane base URL", defaultAddinUrl())
  .option("--backend-url <url>", "Backend WebSocket URL", defaultBackendUrl())
  .action((options: SetupOptions) => runSetup(options, "upgrade"));

program
  .command("instructions")
  .description("Print or write the generic Open Workbook Excel agent instructions")
  .option("--out <path>", "Write instructions to a file instead of stdout")
  .action((options: { out?: string }) => {
    const instructions = generateGenericInstructions();
    if (options.out) {
      writeFileEnsuringDir(resolve(options.out), instructions);
      console.log(`Wrote instructions to ${resolve(options.out)}`);
      return;
    }
    console.log(instructions);
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
  .command("smoke")
  .description("Run a real Excel host smoke through the native file bridge")
  .requiredOption("--workbook <nameOrPath>", "Open Excel workbook name or full path to match")
  .option("--target <path>", "Target .xlsx path for the smoke output")
  .option("--operation <operation>", "Operation: export-copy or save-as", "export-copy")
  .option("--bridge-url <url>", "File bridge base URL", defaultFileBridgeUrl())
  .option("--bridge-path <path>", "File bridge operation route path")
  .option("--confirm-save-as", "Allow the smoke to run workbook.save_as, which changes the open workbook identity")
  .action(async (options: { workbook: string; target?: string; operation: string; bridgeUrl: string; bridgePath?: string; confirmSaveAs?: boolean }) => {
    const operation = normalizeFileBridgeSmokeOperation(options.operation);
    if (operation === "workbook.save_as" && !options.confirmSaveAs) {
      fail("Refusing to run save-as smoke without --confirm-save-as because it changes the open workbook file identity. Use the default export-copy smoke for non-destructive verification.");
    }

    const bridgeUrl = trimTrailingSlash(options.bridgeUrl);
    const statusResponse = await fetchLocal(`${bridgeUrl}/status`, "File bridge smoke status", "owb file-bridge start");
    if (!statusResponse.ok) {
      fail(`File bridge smoke status failed: ${statusResponse.status} ${statusResponse.statusText}`);
    }
    const status = await statusResponse.json().catch(() => undefined) as { route?: string } | undefined;
    if (!status || typeof status !== "object") {
      fail("File bridge smoke status returned non-JSON output.");
    }

    const route = normalizeBridgePath(options.bridgePath ?? status.route ?? defaultFileBridgePath());
    const targetPath = resolve(options.target ?? defaultHostSmokeTarget(operation));
    const request = {
      operation,
      workbookId: options.workbook,
      targetPath,
      reason: "Open Workbook native file bridge smoke"
    };
    const smokeResponse = await fetchLocal(`${bridgeUrl}${route}`, "File bridge smoke operation", "owb file-bridge start", {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify(request)
    });
    if (!smokeResponse.ok) {
      fail(`File bridge smoke operation failed: ${smokeResponse.status} ${smokeResponse.statusText}`);
    }
    const result = await smokeResponse.json().catch(() => undefined) as { ok?: boolean; error?: string } | undefined;
    if (!result || typeof result !== "object") {
      fail("File bridge smoke operation returned non-JSON output.");
    }
    if (result.ok !== true) {
      fail(`File bridge smoke failed: ${result.error ?? "unknown bridge error"}`);
    }

    console.log(JSON.stringify({ ok: true, bridgeUrl, route, request, result }, null, 2));
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
          instructions: resolveAsset("instructions"),
          stateDir: defaultStateDir(),
          exportDir: defaultExportDir(),
          userConfigDir: defaultUserConfigDir(),
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
  .action(async () => {
    const checks = [
      checkPath("MCP server", resolveAsset("mcpServer")),
      checkPath("Backend daemon", resolveAsset("backend")),
      checkPath("File bridge", resolveAsset("fileBridge")),
      checkPath("Add-in server", resolveAsset("addinServer")),
      checkPath("Manifest", resolveAsset("manifest")),
      checkPath("Instructions", resolveAsset("instructions"))
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
    await printAddinServerStatus(defaultAddinUrl());
    await printLatestVersionNotice();
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

function runNode(entrypoint: string, args: string[] = [], companionProcesses: ChildProcess[] = []): void {
  if (!existsSync(entrypoint)) {
    fail(`Missing built entrypoint: ${entrypoint}\nRun: corepack pnpm build`);
  }

  const cleanup = () => {
    for (const companion of companionProcesses) {
      if (!companion.killed) {
        companion.kill();
      }
    }
  };

  const child = spawn(process.execPath, [entrypoint, ...args], {
    stdio: "inherit",
    env: childEnv()
  });

  process.once("SIGINT", () => {
    cleanup();
    child.kill("SIGINT");
  });
  process.once("SIGTERM", () => {
    cleanup();
    child.kill("SIGTERM");
  });

  child.on("exit", (code, signal) => {
    cleanup();
    if (signal) {
      process.kill(process.pid, signal);
    }
    process.exit(code ?? 0);
  });
}

async function startAddinServerIfNeeded(): Promise<ChildProcess | undefined> {
  const addinUrl = defaultAddinUrl();
  const status = await fetchAddinServerStatus(addinUrl);
  if (status.kind === "open-workbook") {
    if (status.version === currentVersion) {
      return undefined;
    }
    fail(
      [
        `Open Workbook add-in server ${status.version} is already running at ${addinUrl}, but this CLI is ${currentVersion}.`,
        "Restart your MCP host or agent UI so it starts the new runtime.",
        `Then run: npx -y ${publicPackageName}@latest mcp`
      ].join("\n")
    );
  }
  if (status.kind === "unknown") {
    fail(
      [
        `A server is already running at ${addinUrl}, but it did not identify as the current Open Workbook add-in server.`,
        "This is usually a stale Open Workbook runtime from an older npm version.",
        "Restart your MCP host or agent UI, or stop the process using that port, then try again."
      ].join("\n")
    );
  }
  if (await urlAvailable(addinUrl)) {
    fail(`A non-Open Workbook server is already responding at ${addinUrl}. Stop that process or set OPEN_WORKBOOK_ADDIN_PORT to another port.`);
  }
  const entrypoint = resolveAsset("addinServer");
  if (!existsSync(entrypoint)) {
    fail(`Missing add-in server entrypoint: ${entrypoint}\nRun: corepack pnpm build`);
  }
  return spawn(process.execPath, [entrypoint], {
    stdio: ["ignore", "ignore", "inherit"],
    env: childEnv()
  });
}

function childEnv(): NodeJS.ProcessEnv {
  return {
    ...process.env,
    OPEN_WORKBOOK_VERSION: process.env.OPEN_WORKBOOK_VERSION ?? currentVersion
  };
}

type AddinServerStatus =
  | { kind: "absent" }
  | { kind: "unknown"; statusCode?: number }
  | { kind: "open-workbook"; version: string; payload: Record<string, unknown> };

async function fetchAddinServerStatus(addinUrl: string): Promise<AddinServerStatus> {
  try {
    const response = await fetch(`${trimTrailingSlash(addinUrl)}/status`);
    if (!response.ok) {
      return { kind: "unknown", statusCode: response.status };
    }
    const payload = await response.json().catch(() => undefined) as Record<string, unknown> | undefined;
    if (payload?.service === "open-workbook-addin-server" && typeof payload.version === "string") {
      return { kind: "open-workbook", version: payload.version, payload };
    }
    return { kind: "unknown", statusCode: response.status };
  } catch {
    return { kind: "absent" };
  }
}

async function urlAvailable(url: string): Promise<boolean> {
  try {
    const response = await fetch(url);
    return response.ok;
  } catch {
    return false;
  }
}

async function printAddinServerStatus(addinUrl: string): Promise<void> {
  const status = await fetchAddinServerStatus(addinUrl);
  if (status.kind === "absent") {
    console.log("Add-in server: not running");
    return;
  }
  if (status.kind === "unknown") {
    console.log(`Add-in server: unknown service at ${addinUrl}`);
    return;
  }
  const marker = status.version === currentVersion ? "ok" : "stale";
  console.log(`Add-in server: ${marker} ${status.version} at ${addinUrl}`);
  if (status.version !== currentVersion) {
    console.log("Restart the agent UI or MCP host before loading the Excel add-in.");
  }
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

function generateGenericInstructions(): string {
  const instructionsDir = resolveAsset("instructions");
  const skillPath = join(instructionsDir, "SKILL.md");
  if (!existsSync(skillPath)) {
    fail(`Missing instruction source: ${skillPath}`);
  }
  const sections = [
    "# Open Workbook Excel Instructions",
    stripFrontmatter(readFileSync(skillPath, "utf8")).trim()
  ];
  const references = [
    ["Tool Selection", "tool-selection.md"],
    ["Workflows", "workflows.md"],
    ["Reliability", "reliability.md"],
    ["Performance", "performance.md"],
    ["Multi-Agent", "multi-agent.md"]
  ] as const;
  for (const [title, fileName] of references) {
    const referencePath = join(instructionsDir, "references", fileName);
    if (existsSync(referencePath)) {
      sections.push(`## ${title}\n\n${readFileSync(referencePath, "utf8").trim()}`);
    }
  }
  return `${sections.join("\n\n")}\n`;
}

function stripFrontmatter(markdown: string): string {
  if (!markdown.startsWith("---\n")) {
    return markdown;
  }
  const end = markdown.indexOf("\n---\n", 4);
  return end === -1 ? markdown : markdown.slice(end + 5);
}

function writeFileEnsuringDir(path: string, content: string): void {
  mkdirSync(dirname(path), { recursive: true });
  writeFileSync(path, content, "utf8");
}

type SetupOptions = {
  dryRun?: boolean;
  instructionsOut: string;
  manifestOut: string;
  addinUrl: string;
  backendUrl: string;
};

async function runSetup(options: SetupOptions, mode: "setup" | "upgrade"): Promise<void> {
  const instructionsPath = resolve(options.instructionsOut);
  const manifestPath = setupManifestPath(options.manifestOut);
  const instructions = generateGenericInstructions();
  const manifest = generateManifest({ addinUrl: options.addinUrl, backendUrl: options.backendUrl });

  if (!options.dryRun) {
    writeFileEnsuringDir(instructionsPath, instructions);
    writeFileEnsuringDir(manifestPath, manifest);
  }

  const action = mode === "upgrade" ? "upgrade" : "setup";
  console.log(options.dryRun ? `Open Workbook ${action} dry run` : `Open Workbook ${action} complete`);
  console.log("");
  console.log(`${options.dryRun ? "Would write" : "Wrote"} fallback instructions: ${instructionsPath}`);
  console.log(`${options.dryRun ? "Would write" : "Wrote"} Excel add-in manifest: ${manifestPath}`);
  console.log(`Taskpane URL: ${options.addinUrl}`);
  console.log(`Backend URL: ${options.backendUrl}`);
  console.log("");
  printManifestNextSteps(manifestPath);
  console.log("");
  console.log("Paste this generic MCP config into any MCP-capable agent UI:");
  console.log(JSON.stringify(genericMcpConfig(), null, 2));
  console.log("");
  console.log("Install or update the Open Workbook Excel skill with skills.sh:");
  console.log("npx skills add components-kit/open-workbook --skill open-workbook-excel");
  console.log("");
  console.log("For a global OpenCode install:");
  console.log("npx skills add components-kit/open-workbook --skill open-workbook-excel -a opencode -g -y");
  console.log("");
  console.log("The fallback instruction file above is for clients that do not support skills.sh.");
  console.log("Start the agent UI before opening the Excel add-in so `npx ... mcp` can serve the taskpane and backend.");
  console.log("After upgrading Open Workbook, restart the agent UI or MCP host so `@latest` starts the new runtime.");
  await printLatestVersionNotice();
}

function genericMcpConfig(): unknown {
  return {
    mcpServers: {
      "open-workbook": {
        command: "npx",
        args: ["-y", `${publicPackageName}@latest`, "mcp"]
      }
    }
  };
}

async function printLatestVersionNotice(): Promise<void> {
  const latest = await fetchLatestPackageVersion();
  if (latest.status === "disabled") {
    return;
  }
  if (latest.status === "unavailable") {
    console.log(`Update check: unavailable (${latest.reason})`);
    return;
  }
  if (compareSemver(latest.version, currentVersion) <= 0) {
    console.log(`Update check: current ${currentVersion}`);
    return;
  }
  console.log("");
  console.log(`Open Workbook ${latest.version} is available. Current CLI: ${currentVersion}.`);
  console.log(`Upgrade: npx -y ${publicPackageName}@latest upgrade`);
  console.log("After upgrading, restart OpenCode or your MCP host so it starts the new runtime.");
}

type LatestVersionResult =
  | { status: "disabled" }
  | { status: "unavailable"; reason: string }
  | { status: "available"; version: string };

async function fetchLatestPackageVersion(): Promise<LatestVersionResult> {
  if (process.env.OPEN_WORKBOOK_DISABLE_UPDATE_CHECK === "1") {
    return { status: "disabled" };
  }
  try {
    const registryUrl = process.env.OPEN_WORKBOOK_UPDATE_CHECK_URL ?? `https://registry.npmjs.org/${encodeURIComponent(publicPackageName)}/latest`;
    const response = await fetch(registryUrl, {
      signal: AbortSignal.timeout(1_500)
    });
    if (!response.ok) {
      return { status: "unavailable", reason: `${response.status} ${response.statusText}` };
    }
    const payload = await response.json().catch(() => undefined) as { version?: unknown } | undefined;
    if (typeof payload?.version !== "string") {
      return { status: "unavailable", reason: "registry response did not include a version" };
    }
    return { status: "available", version: payload.version };
  } catch (error) {
    return { status: "unavailable", reason: error instanceof Error ? error.message : String(error) };
  }
}

function compareSemver(left: string, right: string): number {
  const leftParts = parseSemverCore(left);
  const rightParts = parseSemverCore(right);
  for (let index = 0; index < 3; index += 1) {
    const delta = leftParts[index]! - rightParts[index]!;
    if (delta !== 0) {
      return delta;
    }
  }
  return 0;
}

function parseSemverCore(value: string): [number, number, number] {
  const core = value.split("-", 1)[0] ?? "";
  const parts = core.split(".").map((part) => Number.parseInt(part, 10));
  return [
    Number.isFinite(parts[0]) ? parts[0]! : 0,
    Number.isFinite(parts[1]) ? parts[1]! : 0,
    Number.isFinite(parts[2]) ? parts[2]! : 0
  ];
}

function readPackageVersion(paths: string[]): string | undefined {
  for (const packageJsonPath of paths) {
    if (!existsSync(packageJsonPath)) {
      continue;
    }
    const parsed = JSON.parse(readFileSync(packageJsonPath, "utf8")) as { version?: unknown };
    if (typeof parsed.version === "string") {
      return parsed.version;
    }
  }
  return undefined;
}

function printManifestNextSteps(manifestPath: string): void {
  if (process.platform === "darwin") {
    console.log("Restart Excel, then open Insert > Add-ins > Open Workbook.");
    return;
  }
  if (process.platform === "win32") {
    console.log("Windows Excel sideloading requires a trusted shared-folder add-in catalog:");
    console.log("1. Copy the manifest into a shared folder, for example C:\\open-workbook-addins.");
    console.log("2. In Excel, open File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs.");
    console.log("3. Add the shared-folder UNC path, select Show in Menu, restart Excel, and insert Open Workbook.");
    return;
  }
  console.log(`Use the generated manifest at ${manifestPath} with an Office add-in catalog supported by your Excel host.`);
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

function defaultFileBridgePath(): string {
  return process.env.OPEN_WORKBOOK_FILE_BRIDGE_PATH ?? "/v1/workbook-file";
}

function defaultHostSmokeTarget(operation: "workbook.export_copy" | "workbook.save_as"): string {
  const suffix = operation === "workbook.export_copy" ? "export-copy" : "save-as";
  return resolve(process.cwd(), ".open-workbook/host-smoke", `open-workbook-${suffix}-${Date.now()}.xlsx`);
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

function defaultUserConfigDir(): string {
  return process.env.OPEN_WORKBOOK_CONFIG_DIR ?? join(homedir(), ".open-workbook");
}

function defaultInstructionsPath(): string {
  return join(defaultUserConfigDir(), "instructions", instructionFileName);
}

function defaultSetupManifestPath(): string {
  return join(defaultUserConfigDir(), "open-workbook.xml");
}

function setupManifestPath(manifestOut: string): string {
  if (process.platform === "darwin" && manifestOut === defaultSetupManifestPath()) {
    return join(homedir(), "Library/Containers/com.microsoft.Excel/Data/Documents/wef/open-workbook.xml");
  }
  return resolve(manifestOut);
}

function trimTrailingSlash(value: string): string {
  return value.endsWith("/") ? value.slice(0, -1) : value;
}

function normalizeBridgePath(value: string): string {
  return value.startsWith("/") ? value : `/${value}`;
}

function normalizeFileBridgeSmokeOperation(value: string): "workbook.export_copy" | "workbook.save_as" {
  if (value === "export-copy" || value === "export_copy" || value === "workbook.export_copy") {
    return "workbook.export_copy";
  }
  if (value === "save-as" || value === "save_as" || value === "workbook.save_as") {
    return "workbook.save_as";
  }
  fail(`Unknown file bridge smoke operation: ${value}. Use export-copy or save-as.`);
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
