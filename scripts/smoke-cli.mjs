#!/usr/bin/env node
import { mkdtempSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { spawn, spawnSync } from "node:child_process";
import { createServer } from "node:http";

const cli = "packages/cli/dist/index.js";
const tempDir = mkdtempSync(join(tmpdir(), "open-workbook-cli-smoke-"));
const manifestPath = join(tempDir, "open-workbook.xml");

const checks = [
  {
    name: "doctor",
    args: ["doctor"],
    assert: (result) => result.status === 0 && result.stdout.includes("ok MCP server") && result.stdout.includes("Taskpane URL:")
  },
  {
    name: "paths",
    args: ["paths"],
    assert: (result) => {
      if (result.status !== 0) {
        return false;
      }
      const parsed = JSON.parse(result.stdout);
      return Boolean(parsed.mcpServer) && Boolean(parsed.backend) && Boolean(parsed.fileBridge) && Boolean(parsed.addinServer) && Boolean(parsed.manifest) && Boolean(parsed.instructions) && Boolean(parsed.fileBridgeUrl);
    }
  },
  {
    name: "setup dry run",
    args: ["setup", "--dry-run", "--manifest-out", join(tempDir, "setup-manifest.xml"), "--instructions-out", join(tempDir, "instructions.md")],
    assert: (result) =>
      result.status === 0 &&
      result.stdout.includes("@components-kit/open-workbook@latest") &&
      result.stdout.includes("MCP launch command for your agent UI:") &&
      result.stdout.includes("npx -y @components-kit/open-workbook@latest mcp") &&
      result.stdout.includes("local stdio MCP server command") &&
      result.stdout.includes("npx skills add components-kit/open-workbook --skill open-workbook-excel") &&
      result.stdout.includes("instructions.md") &&
      result.stdout.includes("setup-manifest.xml") &&
      result.stdout.includes("restart the agent UI or MCP host")
  },
  {
    name: "upgrade dry run",
    args: ["upgrade", "--dry-run", "--manifest-out", join(tempDir, "upgrade-manifest.xml"), "--instructions-out", join(tempDir, "upgrade-instructions.md")],
    assert: (result) =>
      result.status === 0 &&
      result.stdout.includes("Open Workbook upgrade dry run") &&
      result.stdout.includes("upgrade-instructions.md") &&
      result.stdout.includes("upgrade-manifest.xml") &&
      result.stdout.includes("@components-kit/open-workbook@latest") &&
      result.stdout.includes("MCP launch command for your agent UI:") &&
      result.stdout.includes("npx -y @components-kit/open-workbook@latest mcp") &&
      result.stdout.includes("local stdio MCP server command") &&
      result.stdout.includes("restart the agent UI or MCP host")
  },
  {
    name: "instructions",
    args: ["instructions"],
    assert: (result) =>
      result.status === 0 &&
      result.stdout.includes("# Open Workbook Excel Instructions") &&
      result.stdout.includes("excel.runtime.get_status") &&
      result.stdout.includes("## Tool Selection")
  },
  {
    name: "opencode config",
    args: ["opencode", "config", "--id", "open-workbook-smoke", "--command", "node packages/cli/dist/index.js", "--agent-name", "smoke-agent"],
    assert: (result) => {
      if (result.status !== 0) {
        return false;
      }
      const parsed = JSON.parse(result.stdout);
      return parsed.mcp?.["open-workbook-smoke"]?.enabled === true && parsed.mcp["open-workbook-smoke"].command.includes("--agent-name");
    }
  },
  {
    name: "sideload manifest",
    args: ["sideload", "manifest", "--out", manifestPath, "--addin-url", "http://127.0.0.1:37846", "--backend-url", "ws://127.0.0.1:37845/addin"],
    assert: (result) => {
      if (result.status !== 0) {
        return false;
      }
      const manifest = readFileSync(manifestPath, "utf8");
      return manifest.includes("taskpane.html?backendUrl=") && manifest.includes("ws%3A%2F%2F127.0.0.1%3A37845%2Faddin");
    }
  },
  {
    name: "service manifest daemon",
    args: ["service", "manifest", "--target", "systemd", "--service", "daemon", "--command", "owb"],
    assert: (result) => result.status === 0 && result.stdout.includes("ExecStart='owb' 'daemon' 'start'")
  },
  {
    name: "service manifest file bridge",
    args: ["service", "manifest", "--target", "systemd", "--service", "file-bridge", "--command", "owb"],
    assert: (result) => result.status === 0 && result.stdout.includes("ExecStart='owb' 'file-bridge' 'start'")
  },
  {
    name: "file bridge smoke help",
    args: ["file-bridge", "smoke", "--help"],
    assert: (result) => result.status === 0 && result.stdout.includes("Run a real Excel host smoke") && result.stdout.includes("--workbook")
  },
  {
    name: "disconnected daemon status",
    args: ["daemon", "status", "--daemon-url", "http://127.0.0.1:37999"],
    assert: (result) =>
      result.status === 1 &&
      result.stderr.includes("Daemon status failed: could not connect to http://127.0.0.1:37999/status") &&
      result.stderr.includes("owb daemon start")
  },
  {
    name: "disconnected file bridge status",
    args: ["file-bridge", "status", "--bridge-url", "http://127.0.0.1:37998"],
    assert: (result) =>
      result.status === 1 &&
      result.stderr.includes("File bridge status failed: could not connect to http://127.0.0.1:37998/status") &&
      result.stderr.includes("owb file-bridge start")
  }
];

const failures = [];

for (const check of checks) {
  const result = spawnSync(process.execPath, [cli, ...check.args], {
    encoding: "utf8",
    env: {
      ...process.env,
      OPEN_WORKBOOK_DISABLE_UPDATE_CHECK: "1",
      OPEN_WORKBOOK_STATE_DIR: join(tempDir, "state")
    }
  });
  if (!check.assert(result)) {
    failures.push({
      name: check.name,
      status: result.status,
      stdout: result.stdout,
      stderr: result.stderr
    });
  }
}

await smokePackagedAddinServer();
await smokeStaleAddinServerGuard();
await smokeLatestVersionNotice();

if (failures.length > 0) {
  console.error("CLI smoke failed.");
  for (const failure of failures) {
    console.error(`\n[${failure.name}] status=${failure.status}`);
    if (failure.stdout) {
      console.error(`stdout:\n${failure.stdout}`);
    }
    if (failure.stderr) {
      console.error(`stderr:\n${failure.stderr}`);
    }
  }
  process.exit(1);
}

console.log(`CLI smoke passed: ${checks.length + 3} command(s).`);

async function smokePackagedAddinServer() {
  const port = 37977;
  const server = spawn(process.execPath, ["packages/cli/assets/excel-addin/scripts/dev-server.mjs"], {
    stdio: ["ignore", "pipe", "pipe"],
    env: {
      ...process.env,
      OPEN_WORKBOOK_ADDIN_PORT: String(port)
    }
  });
  let stdout = "";
  let stderr = "";
  server.stdout?.on("data", (chunk) => {
    stdout += String(chunk);
  });
  server.stderr?.on("data", (chunk) => {
    stderr += String(chunk);
  });

  try {
    await waitForHttp(`http://127.0.0.1:${port}/taskpane.html`);
    const status = JSON.parse(await fetchText(`http://127.0.0.1:${port}/status`));
    const moduleChecks = await Promise.all([
      fetchText(`http://127.0.0.1:${port}/workspace/excel-core/index.js`),
      fetchText(`http://127.0.0.1:${port}/workspace/protocol/index.js`)
    ]);
    if (
      status.service !== "open-workbook-addin-server" ||
      typeof status.version !== "string" ||
      status.workspaceModules?.["excel-core"]?.available !== true ||
      status.workspaceModules?.protocol?.available !== true ||
      !moduleChecks.every((body) => body.includes("export"))
    ) {
      failures.push({
        name: "packaged addin server status and workspace modules",
        status: 1,
        stdout,
        stderr: `Expected status and workspace module endpoints to report healthy JavaScript assets.\nstatus=${JSON.stringify(status)}\n${stderr}`
      });
    }
  } catch (error) {
    failures.push({
      name: "packaged addin server status and workspace modules",
      status: server.exitCode,
      stdout,
      stderr: `${error instanceof Error ? error.message : String(error)}\n${stderr}`.trim()
    });
  } finally {
    server.kill();
  }
}

async function smokeStaleAddinServerGuard() {
  const port = 37978;
  const server = createServer((request, response) => {
    if (request.url === "/status") {
      response.writeHead(200, { "content-type": "application/json" });
      response.end(JSON.stringify({ ok: true, service: "open-workbook-addin-server", version: "0.0.1" }));
      return;
    }
    response.writeHead(404, { "content-type": "application/json" });
    response.end(JSON.stringify({ error: "not_found" }));
  });

  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(port, "127.0.0.1", resolve);
  });

  try {
    const result = await spawnCli(["mcp"], {
      OPEN_WORKBOOK_ADDIN_PORT: String(port),
      OPEN_WORKBOOK_DISABLE_UPDATE_CHECK: "1",
      OPEN_WORKBOOK_STATE_DIR: join(tempDir, "stale-state")
    });
    if (result.status !== 1 || !result.stderr.includes("add-in server 0.0.1") || !result.stderr.includes("Restart your MCP host")) {
      failures.push({
        name: "stale addin server guard",
        status: result.status,
        stdout: result.stdout,
        stderr: result.stderr
      });
    }
  } finally {
    await new Promise((resolve) => server.close(resolve));
  }
}

async function smokeLatestVersionNotice() {
  const port = 37979;
  const server = createServer((request, response) => {
    if (request.url === "/latest") {
      response.writeHead(200, { "content-type": "application/json" });
      response.end(JSON.stringify({ version: "99.0.0" }));
      return;
    }
    response.writeHead(404, { "content-type": "application/json" });
    response.end(JSON.stringify({ error: "not_found" }));
  });

  await new Promise((resolve, reject) => {
    server.once("error", reject);
    server.listen(port, "127.0.0.1", resolve);
  });

  try {
    const result = await spawnCli(["doctor"], {
      OPEN_WORKBOOK_ADDIN_PORT: "37980",
      OPEN_WORKBOOK_UPDATE_CHECK_URL: `http://127.0.0.1:${port}/latest`,
      OPEN_WORKBOOK_STATE_DIR: join(tempDir, "update-check-state")
    });
    if (result.status !== 0 || !result.stdout.includes("Open Workbook 99.0.0 is available") || !result.stdout.includes("Upgrade: npx -y @components-kit/open-workbook@latest upgrade")) {
      failures.push({
        name: "latest version notice",
        status: result.status,
        stdout: result.stdout,
        stderr: result.stderr
      });
    }
  } finally {
    await new Promise((resolve) => server.close(resolve));
  }
}

function spawnCli(args, env) {
  return new Promise((resolve) => {
    const child = spawn(process.execPath, [cli, ...args], {
      stdio: ["ignore", "pipe", "pipe"],
      env: {
        ...process.env,
        ...env
      }
    });
    let stdout = "";
    let stderr = "";
    const timeout = setTimeout(() => {
      child.kill();
      resolve({ status: -1, stdout, stderr: `${stderr}\nTimed out waiting for CLI process.`.trim() });
    }, 5_000);
    child.stdout.on("data", (chunk) => {
      stdout += String(chunk);
    });
    child.stderr.on("data", (chunk) => {
      stderr += String(chunk);
    });
    child.on("exit", (code) => {
      clearTimeout(timeout);
      resolve({ status: code, stdout, stderr });
    });
  });
}

async function waitForHttp(url) {
  const deadline = Date.now() + 5_000;
  let lastError;
  while (Date.now() < deadline) {
    try {
      await fetchText(url);
      return;
    } catch (error) {
      lastError = error;
      await new Promise((resolve) => setTimeout(resolve, 100));
    }
  }
  throw lastError ?? new Error(`Timed out waiting for ${url}`);
}

async function fetchText(url) {
  const response = await fetch(url);
  if (!response.ok) {
    throw new Error(`${url} returned ${response.status}`);
  }
  return response.text();
}
