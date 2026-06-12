#!/usr/bin/env node
import { mkdtempSync, readFileSync } from "node:fs";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { spawnSync } from "node:child_process";

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
      return Boolean(parsed.mcpServer) && Boolean(parsed.backend) && Boolean(parsed.fileBridge) && Boolean(parsed.addinServer) && Boolean(parsed.manifest) && Boolean(parsed.fileBridgeUrl);
    }
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

console.log(`CLI smoke passed: ${checks.length} command(s).`);
