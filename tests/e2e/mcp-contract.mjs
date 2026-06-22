#!/usr/bin/env node
import { spawn } from "node:child_process";
import { mkdtempSync, mkdirSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const tempRoot = mkdtempSync(path.join(tmpdir(), "open-workbook-e2e-mcp-contract-"));
const artifactsDir = path.join(tempRoot, "artifacts");
mkdirSync(artifactsDir, { recursive: true });

const backendPort = 40180 + Math.floor(Math.random() * 500);
const transcript = [];

async function main() {
  const server = spawn(process.execPath, ["apps/mcp-server/dist/index.js", "--standalone", "--agent-name", "e2e-mcp-contract"], {
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
    const initialized = await mcp.request("initialize", {
      protocolVersion: "2024-11-05",
      capabilities: {},
      clientInfo: { name: "open-workbook-e2e-mcp-contract", version: "0.0.0" }
    });
    assert(initialized.protocolVersion, "initialize should return negotiated protocolVersion");
    mcp.notify("notifications/initialized", {});

    const listed = await mcp.request("tools/list", {});
    const tool = listed.tools?.find((item) => item.name === "excel.agent.run");
    assert(listed.tools?.length === 1 && tool, "tools/list should expose only excel.agent.run");
    assert(tool.inputSchema?.properties?.request, "excel.agent.run should publish request input schema");
    assert(tool.outputSchema?.properties?.status, "excel.agent.run should publish status output schema");

    await expectToolResult(mcp, "valid status call", {
      name: "excel.agent.run",
      arguments: { request: "Check Open Workbook status", mode: "status" }
    }, (result) => {
      assert(result.structuredContent?.status === "NEEDS_INPUT", "status without Excel should be a structured NEEDS_INPUT result");
      assert(result.structuredContent?.nextAction === "ask_user", "status without Excel should ask user for setup");
      assert(result.content?.some((item) => item.type === "text" && item.text.includes("NEEDS_INPUT")), "tool text should include structured status");
      assert(result.isError !== true, "business setup state should not be an MCP protocol failure");
    });

    await expectCallFailure(mcp, "unknown tool", {
      name: "excel.missing.tool",
      arguments: { request: "noop" }
    }, { messageIncludes: "tool", codeOneOf: [-32602, -32603] });

    await expectCallFailure(mcp, "missing required request", {
      name: "excel.agent.run",
      arguments: { mode: "status" }
    }, { messageIncludes: "request", codeOneOf: [-32602] });

    await expectCallFailure(mcp, "invalid mode enum", {
      name: "excel.agent.run",
      arguments: { request: "Check status", mode: "not_a_mode" }
    }, { messageIncludes: "mode", codeOneOf: [-32602] });

    await expectCallFailure(mcp, "bad target string", {
      name: "excel.agent.run",
      arguments: { request: "Read Booking A1", target: "Booking!A1" }
    }, { messageIncludes: "target", codeOneOf: [-32602] });

    await expectCallFailure(mcp, "bad values patch target string", {
      name: "excel.agent.run",
      arguments: {
        request: "Update Booking A1",
        values: { patches: [{ target: "Booking!A1", values: [["x"]] }] }
      }
    }, { messageIncludes: "target", codeOneOf: [-32602] });

    await expectCallFailure(mcp, "bad style fill color type", {
      name: "excel.agent.run",
      arguments: {
        request: "Set Sales header fill to black.",
        mode: "preview_update",
        intent: { action: "format_range" },
        target: { sheetName: "Sales", range: "A1:E1" },
        values: { style: { fillColor: 42 } }
      }
    }, { messageIncludes: "fillColor", codeOneOf: [-32602] });

    await expectCallFailure(mcp, "bad dropdown option type", {
      name: "excel.agent.run",
      arguments: {
        request: "Add a dropdown list to Sales E2:E6.",
        mode: "preview_update",
        intent: { action: "write_data_validation" },
        target: { sheetName: "Sales", range: "E2:E6" },
        values: { options: ["Open", 7, "Closed"] }
      }
    }, { messageIncludes: "options", codeOneOf: [-32602] });

    await expectCallFailure(mcp, "bad validation source type", {
      name: "excel.agent.run",
      arguments: {
        request: "Add a dropdown list to Sales E2:E6.",
        mode: "preview_update",
        intent: { action: "write_data_validation" },
        target: { sheetName: "Sales", range: "E2:E6" },
        values: { validation: { source: ["Open", 7] } }
      }
    }, { messageIncludes: "source", codeOneOf: [-32602] });

    await expectCallFailure(mcp, "bad conditional formula type", {
      name: "excel.agent.run",
      arguments: {
        request: "Add conditional formatting to Sales A2:E6.",
        mode: "preview_update",
        intent: { action: "write_conditional_formatting" },
        target: { sheetName: "Sales", range: "A2:E6" },
        values: { rule: { formula: 42, style: { fillColor: "#FFFF00" } } }
      }
    }, { messageIncludes: "formula", codeOneOf: [-32602] });

    await expectCallFailure(mcp, "bad column order entry", {
      name: "excel.agent.run",
      arguments: {
        request: "Swap Sales columns Date and Customer.",
        mode: "preview_update",
        intent: { action: "reorder_range_columns" },
        target: { sheetName: "Sales", range: "A1:B6" },
        values: { columnOrder: [2, { bad: true }] }
      }
    }, { messageIncludes: "columnOrder", codeOneOf: [-32602] });

    await expectCallFailure(mcp, "arguments must be object", {
      name: "excel.agent.run",
      arguments: "{\"request\":\"Check status\",\"mode\":\"status\"}"
    }, { messageIncludes: "arguments", codeOneOf: [-32602, -32603] });

    const resources = await mcp.request("resources/list", {});
    assert(Array.isArray(resources.resources), "resources/list should return a resources array");

    await expectResourceFailure(mcp, "unknown resource", {
      uri: "excel://agent/results/missing"
    }, { messageIncludes: "resource", codeOneOf: [-32002, -32602, -32603] });

    writeArtifact("mcp-contract-transcript.json", { transcript });
    console.log(`MCP contract E2E passed. Artifacts: ${artifactsDir}`);
  } catch (error) {
    writeArtifact("mcp-contract-failure.json", {
      error: error instanceof Error ? { message: error.message, stack: error.stack } : String(error),
      transcript,
      serverStderr
    });
    console.error(`MCP contract E2E failed. Artifacts: ${artifactsDir}`);
    throw error;
  } finally {
    mcp.close();
    server.kill();
  }
}

async function expectToolResult(client, label, params, assertResult) {
  const result = await client.request("tools/call", params);
  transcript.push({ label, ok: true, result });
  assertResult(result);
}

async function expectCallFailure(client, label, params, expectation) {
  try {
    const result = await client.request("tools/call", params);
    const text = result.content?.map((item) => item.text ?? "").join("\n") ?? "";
    assert(result.isError === true, `${label} should fail as a tool error or JSON-RPC error`);
    assert(text.toLowerCase().includes(expectation.messageIncludes.toLowerCase()), `${label} error should mention ${expectation.messageIncludes}, got ${text}`);
    assert(!text.includes("inputSchema"), `${label} error should not dump inputSchema`);
    assert(!text.includes("outputSchema"), `${label} error should not dump outputSchema`);
    assert(!/send.*schema|resend.*schema/i.test(text), `${label} error should not tell agents to resend schemas`);
    transcript.push({ label, ok: false, errorKind: "tool", result });
  } catch (error) {
    const rpcError = error?.rpcError;
    assert(rpcError, `${label} should fail with a JSON-RPC error or tool error`);
    if (expectation.codeOneOf) {
      assert(expectation.codeOneOf.includes(rpcError.code), `${label} expected code ${expectation.codeOneOf.join(" or ")}, got ${rpcError.code}`);
    }
    assert(String(rpcError.message).toLowerCase().includes(expectation.messageIncludes.toLowerCase()), `${label} error should mention ${expectation.messageIncludes}, got ${rpcError.message}`);
    assert(!String(rpcError.message).includes("inputSchema"), `${label} error should not dump inputSchema`);
    assert(!String(rpcError.message).includes("outputSchema"), `${label} error should not dump outputSchema`);
    assert(!/send.*schema|resend.*schema/i.test(String(rpcError.message)), `${label} error should not tell agents to resend schemas`);
    transcript.push({ label, ok: false, errorKind: "json-rpc", error: rpcError });
  }
}

async function expectResourceFailure(client, label, params, expectation) {
  try {
    const result = await client.request("resources/read", params);
    const text = result.contents?.map((item) => item.text ?? "").join("\n") ?? JSON.stringify(result);
    assert(text.toLowerCase().includes(expectation.messageIncludes.toLowerCase()) || text.toLowerCase().includes("not found"), `${label} error should mention ${expectation.messageIncludes}, got ${text}`);
    transcript.push({ label, ok: false, errorKind: "resource-result", result });
  } catch (error) {
    const rpcError = error?.rpcError;
    assert(rpcError, `${label} should fail with a JSON-RPC error or resource error result`);
    if (expectation.codeOneOf) {
      assert(expectation.codeOneOf.includes(rpcError.code), `${label} expected code ${expectation.codeOneOf.join(" or ")}, got ${rpcError.code}`);
    }
    assert(String(rpcError.message).toLowerCase().includes(expectation.messageIncludes.toLowerCase()) || String(rpcError.message).toLowerCase().includes("not found"), `${label} error should mention ${expectation.messageIncludes}, got ${rpcError.message}`);
    transcript.push({ label, ok: false, errorKind: "json-rpc", error: rpcError });
  }
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
      const error = new Error(message.error.message);
      error.rpcError = message.error;
      pending.reject(error);
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
