#!/usr/bin/env node
import { spawn } from "node:child_process";
import { mkdtempSync, mkdirSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const tempRoot = mkdtempSync(path.join(tmpdir(), "open-workbook-e2e-agent-surface-"));
const artifactsDir = path.join(tempRoot, "artifacts");
mkdirSync(artifactsDir, { recursive: true });

const backendPort = 39280 + Math.floor(Math.random() * 500);
const transcript = [];

async function main() {
  const server = spawn(process.execPath, ["apps/mcp-server/dist/index.js", "--standalone", "--agent-name", "e2e-agent-surface"], {
    cwd: repoRoot,
    stdio: ["pipe", "pipe", "pipe"],
    env: {
      ...process.env,
      OPEN_WORKBOOK_HOST: "127.0.0.1",
      OPEN_WORKBOOK_PORT: String(backendPort),
      OPEN_WORKBOOK_ADDIN_PATH: "/addin",
      OPEN_WORKBOOK_STATE_DIR: path.join(tempRoot, "state"),
      OPEN_WORKBOOK_BACKUP_DIR: path.join(tempRoot, "backups"),
      OPEN_WORKBOOK_DISABLE_UPDATE_CHECK: "1"
    }
  });
  const mcp = new McpClient(server);
  let serverStderr = "";
  server.stderr.on("data", (chunk) => {
    serverStderr += String(chunk);
  });

  try {
    await mcp.request("initialize", {
      protocolVersion: "2024-11-05",
      capabilities: {},
      clientInfo: { name: "open-workbook-e2e-agent-surface", version: "0.0.0" }
    });
    mcp.notify("notifications/initialized", {});

    const listed = await mcp.request("tools/list", {});
    const toolNames = listed.tools.map((tool) => tool.name);
    assert(toolNames.length === 1 && toolNames[0] === "excel.agent.run", `default surface should expose only excel.agent.run, got ${toolNames.join(", ")}`);
    assert(listed.tools[0].inputSchema?.properties?.mode, "excel.agent.run should publish its request schema");

    const status = await callTool(mcp, "excel.agent.run", { request: "Check Open Workbook status", mode: "status" });
    assert(status.status === "SUCCESS", `agent status should succeed: ${JSON.stringify(status)}`);
    assert(status.nextAction === "manual_review", "status without add-in should tell the agent to request manual review");
    assert(status.telemetry?.payloadBytes > 0 && status.telemetry?.estimatedTokens > 0, "agent run should include payload and token telemetry");

    const resources = await mcp.request("resources/list", {});
    assert(Array.isArray(resources.resources), "resources/list should be available on the default surface");

    writeArtifact("e2e-agent-surface-transcript.json", { transcript, toolNames });
    console.log(`E2E agent surface passed. Artifacts: ${artifactsDir}`);
  } catch (error) {
    writeArtifact("e2e-agent-surface-failure.json", {
      error: error instanceof Error ? { message: error.message, stack: error.stack } : String(error),
      transcript,
      serverStderr
    });
    console.error(`E2E agent surface failed. Artifacts: ${artifactsDir}`);
    throw error;
  } finally {
    mcp.close();
    server.kill();
  }
}

async function callTool(client, name, args) {
  const result = await client.request("tools/call", { name, arguments: args });
  if (result.isError) {
    throw new Error(`${name} returned MCP error: ${JSON.stringify(result)}`);
  }
  const text = result.content?.find((item) => item.type === "text")?.text;
  assert(text, `${name} returned no text content`);
  const parsed = JSON.parse(text);
  transcript.push({ tool: name, args, result: parsed });
  return parsed;
}

class McpClient {
  constructor(child) {
    this.child = child;
    this.nextId = 1;
    this.pending = new Map();
    this.buffer = Buffer.alloc(0);
    child.stdout.on("data", (chunk) => this.read(chunk));
    child.on("exit", (code, signal) => {
      for (const pending of this.pending.values()) {
        pending.reject(new Error(`MCP server exited code=${code} signal=${signal}`));
      }
      this.pending.clear();
    });
  }

  request(method, params) {
    const id = this.nextId++;
    this.write({ jsonrpc: "2.0", id, method, params });
    return new Promise((resolve, reject) => {
      const timeout = setTimeout(() => {
        this.pending.delete(id);
        reject(new Error(`Timed out waiting for MCP method ${method}`));
      }, 30_000);
      this.pending.set(id, { resolve, reject, timeout });
    });
  }

  notify(method, params) {
    this.write({ jsonrpc: "2.0", method, params });
  }

  write(message) {
    this.child.stdin.write(`${JSON.stringify(message)}\n`);
  }

  read(chunk) {
    this.buffer = Buffer.concat([this.buffer, Buffer.from(chunk)]);
    while (this.buffer.length > 0) {
      const parsed = this.readFramed() ?? this.readLineDelimited();
      if (!parsed) {
        return;
      }
      this.handle(parsed);
    }
  }

  readFramed() {
    const marker = this.buffer.indexOf("\r\n\r\n");
    if (marker < 0) {
      return undefined;
    }
    const header = this.buffer.slice(0, marker).toString("utf8");
    const match = /^Content-Length:\s*(\d+)/im.exec(header);
    if (!match) {
      return undefined;
    }
    const length = Number(match[1]);
    const bodyStart = marker + 4;
    const bodyEnd = bodyStart + length;
    if (this.buffer.length < bodyEnd) {
      return undefined;
    }
    const body = this.buffer.slice(bodyStart, bodyEnd).toString("utf8");
    this.buffer = this.buffer.slice(bodyEnd);
    return JSON.parse(body);
  }

  readLineDelimited() {
    const newline = this.buffer.indexOf("\n");
    if (newline < 0) {
      return undefined;
    }
    const line = this.buffer.slice(0, newline).toString("utf8").trim();
    this.buffer = this.buffer.slice(newline + 1);
    if (!line) {
      return undefined;
    }
    return JSON.parse(line);
  }

  handle(message) {
    if (!("id" in message) || "method" in message) {
      return;
    }
    const pending = this.pending.get(message.id);
    if (!pending) {
      return;
    }
    clearTimeout(pending.timeout);
    this.pending.delete(message.id);
    if (message.error) {
      pending.reject(new Error(message.error.message));
    } else {
      pending.resolve(message.result);
    }
  }

  close() {
    this.child.stdin.destroy();
  }
}

function writeArtifact(name, value) {
  writeFileSync(path.join(artifactsDir, name), JSON.stringify(value, null, 2));
}

function assert(condition, message) {
  if (!condition) {
    throw new Error(message);
  }
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
