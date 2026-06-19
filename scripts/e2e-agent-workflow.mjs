#!/usr/bin/env node
import { spawn } from "node:child_process";
import { createHash } from "node:crypto";
import { mkdtempSync, mkdirSync, writeFileSync } from "node:fs";
import { tmpdir } from "node:os";
import path from "node:path";
import { fileURLToPath } from "node:url";

const repoRoot = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "..");
const tempRoot = mkdtempSync(path.join(tmpdir(), "open-workbook-e2e-agent-workflow-"));
const artifactsDir = path.join(tempRoot, "artifacts");
mkdirSync(artifactsDir, { recursive: true });

const workbookId = "workbook_e2e_agent";
const backendPort = 39780 + Math.floor(Math.random() * 500);
const backendUrl = `http://127.0.0.1:${backendPort}`;
const backendWsUrl = `ws://127.0.0.1:${backendPort}/addin`;
const transcript = [];

async function main() {
  const server = spawn(process.execPath, ["apps/mcp-server/dist/index.js", "--standalone", "--agent-name", "e2e-agent-workflow"], {
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
  let addin;
  let serverStderr = "";
  server.stderr.on("data", (chunk) => {
    serverStderr += String(chunk);
  });

  try {
    await waitForHttp(`${backendUrl}/status`, 15_000);
    addin = await FakeAddin.connect(backendWsUrl, createWorkbookFixture(workbookId));

    await mcp.request("initialize", {
      protocolVersion: "2024-11-05",
      capabilities: {},
      clientInfo: { name: "open-workbook-e2e-agent-workflow", version: "0.0.0" }
    });
    mcp.notify("notifications/initialized", {});

    const listed = await mcp.request("tools/list", {});
    const toolNames = listed.tools.map((tool) => tool.name);
    assert(toolNames.length === 1 && toolNames[0] === "excel.agent.run", `default surface should expose only excel.agent.run, got ${toolNames.join(", ")}`);
    const agentOutputSchema = listed.tools[0].outputSchema;

    const prepared = await agentRun(mcp, { request: "Prepare workbook", mode: "prepare" }, agentOutputSchema);
    assert(prepared.status === "SUCCESS", "prepare should succeed");
    assert(prepared.workbookContextId, "prepare should return workbookContextId");
    assert(prepared.telemetry.metadataCacheStatus === "miss", "first prepare should miss metadata cache");

    const preparedAgain = await agentRun(mcp, { request: "Prepare workbook again", mode: "prepare", workbookContextId: prepared.workbookContextId }, agentOutputSchema);
    assert(preparedAgain.telemetry.cacheHit === true, "second prepare should hit metadata cache");
    assert(preparedAgain.telemetry.metadataCacheStatus === "hit", "second prepare should report cache hit status");

    const found = await agentRun(mcp, { request: "Find Transactions amount status", mode: "find", workbookContextId: prepared.workbookContextId, budget: { maxExamples: 3, maxPayloadBytes: 1200 } }, agentOutputSchema);
    assert(found.status === "SUCCESS", "find should locate workbook candidates");
    assert((found.candidates?.length ?? 0) <= 3, "find should obey maxExamples");
    assert(found.telemetry.payloadBytes <= 1400, `find payload should stay compact, got ${found.telemetry.payloadBytes}`);

    const answered = await agentRun(mcp, { request: "Answer amount from Transactions table", mode: "answer", workbookContextId: prepared.workbookContextId }, agentOutputSchema);
    assert(answered.status === "SUCCESS", "answer should succeed");
    assert(answered.telemetry.internalReadCount === 1, "answer should use one targeted internal read");
    assert(answered.telemetry.fullReadCellCount <= 16, `answer should avoid broad reads, read ${answered.telemetry.fullReadCellCount} cells`);

    const naturalSheet = await agentRun(mcp, { request: "Analyze the June financial sheet", mode: "answer", workbookContextId: prepared.workbookContextId }, agentOutputSchema);
    assert(naturalSheet.status === "SUCCESS", "natural language sheet request should resolve");
    assert(naturalSheet.proof?.[0]?.sheetName === "Financials - June 2026", `natural language sheet request should use June financial sheet, got ${naturalSheet.proof?.[0]?.sheetName}`);

    const ambiguous = await agentRun(mcp, { request: "Analyze financial 2026", mode: "answer", workbookContextId: prepared.workbookContextId }, agentOutputSchema);
    assert(ambiguous.status === "AMBIGUOUS_TARGET", "ambiguous financial request should return candidates instead of guessing");

    const rawAprRange = await agentRun(mcp, {
      request: "Read 'Apr 2026'!O1:AE3 actual values",
      mode: "answer",
      workbookContextId: prepared.workbookContextId
    }, agentOutputSchema);
    assert(rawAprRange.status === "SUCCESS", "quoted raw sheet range should resolve");
    assert(rawAprRange.proof?.[0]?.sheetName === "Apr 2026", `raw range should use Apr 2026, got ${rawAprRange.proof?.[0]?.sheetName}`);
    assert(rawAprRange.proof?.[0]?.range === "O1:AE3", `raw range should preserve O1:AE3, got ${rawAprRange.proof?.[0]?.range}`);
    assert(rawAprRange.answer?.kind === "range_profile" && rawAprRange.answer?.source === "live_read", "raw range should be a live read");

    const rawAprInvoice = await agentRun(mcp, {
      request: "Read Apr 2026 invoice rows",
      mode: "answer",
      workbookContextId: prepared.workbookContextId
    }, agentOutputSchema);
    assert(rawAprInvoice.status === "SUCCESS", "raw invoice block should resolve");
    assert(rawAprInvoice.proof?.[0]?.sheetName === "Apr 2026", "raw invoice block should use Apr 2026");
    assert(rawAprInvoice.proof?.[0]?.range === "O1:AE3", `raw invoice block should use invoice range, got ${rawAprInvoice.proof?.[0]?.range}`);
    assert(rawAprInvoice.telemetry?.internalReadCount === 1, "raw invoice block should perform one live read");

    addin.workbook.sheet("Data").writeValues("F1:I4", [
      ["Date", "Account", "Amount", "Status"],
      ["2026-02-01", "B-100", 300, "Open"],
      ["2026-02-02", "B-200", 400, "Closed"],
      ["2026-02-03", "B-300", 500, "Open"]
    ]);
    addin.workbook.createTable({ sheetName: "Data", address: "F1:I4", tableName: "TransactionsArchive" });
    const refreshed = await agentRun(mcp, { request: "Refresh workbook after adding another transaction table", mode: "prepare", workbookContextId: prepared.workbookContextId }, agentOutputSchema);
    const tableCandidates = await agentRun(mcp, { request: "Find TransactionsArchive table", mode: "find", workbookContextId: refreshed.workbookContextId }, agentOutputSchema);
    const archiveCandidateId = tableCandidates.candidates?.find((candidate) => candidate.tableName === "TransactionsArchive")?.id;
    assert(archiveCandidateId, "find response should include a usable TransactionsArchive candidateId");
    const selectedSchema = await agentRun(mcp, {
      request: "Read the selected table schema",
      mode: "answer",
      workbookContextId: refreshed.workbookContextId,
      target: { candidateId: archiveCandidateId }
    }, agentOutputSchema);
    assert(selectedSchema.status === "SUCCESS", "candidateId retry should resolve the selected schema");
    assert(selectedSchema.answer?.kind === "table_schema", "candidateId schema retry should return table schema");
    assert(selectedSchema.answer?.tableName === "TransactionsArchive", `candidateId schema retry should select TransactionsArchive, got ${selectedSchema.answer?.tableName}`);
    assert(selectedSchema.telemetry?.internalReadCount === 0, "schema answers should use cached metadata without full reads");

    const selectedRows = await agentRun(mcp, {
      request: "Read headers and first 2 rows from the selected table",
      mode: "answer",
      workbookContextId: refreshed.workbookContextId,
      target: { candidateId: archiveCandidateId }
    }, agentOutputSchema);
    assert(selectedRows.status === "SUCCESS", "candidateId row request should read live values");
    assert(selectedRows.answer?.kind === "table_compact_read", "row request should return compact table data");
    assert(selectedRows.answer?.source === "runtime_table_read", "row request should report runtime table read source");
    assert(selectedRows.answer?.profile?.kind === "range_profile", "row request should include a live range profile");
    assert(selectedRows.answer?.profile?.source === "live_read", "row request profile should report live_read source");
    assert(selectedRows.telemetry?.internalReadCount === 1, "row request should perform one internal read");

    const appendPreview = await agentRun(mcp, {
      request: "Append transaction rows to the selected table",
      mode: "preview_update",
      workbookContextId: refreshed.workbookContextId,
      target: { candidateId: archiveCandidateId },
      values: { rows: [["2026-02-04", "B-400", 600, "Open"]] }
    }, agentOutputSchema);
    assert(appendPreview.status === "PREVIEW_READY", "table append should return a pending operation");
    assert(appendPreview.answer?.kind === "table_append_preview", "table append should return append preview metadata");
    const appendApplied = await agentRun(mcp, {
      request: "Apply table append",
      mode: "apply_update",
      operationId: appendPreview.operationId,
      confirmationToken: appendPreview.confirmationToken
    }, agentOutputSchema);
    assert(appendApplied.status === "SUCCESS", "table append apply should succeed");
    assert(addin.workbook.table("TransactionsArchive").info().rowCount === 4, "fake add-in table should have one appended row");

    const preview = await agentRun(mcp, {
      request: "Update Data B2",
      mode: "preview_update",
      workbookContextId: prepared.workbookContextId,
      target: { sheetName: "Data", range: "B2" },
      values: { values: [[999]] }
    }, agentOutputSchema);
    assert(preview.status === "PREVIEW_READY", "preview_update should return a pending operation");
    assert(preview.operationId && preview.confirmationToken, "preview_update should return operationId and confirmationToken");

    const missingToken = await agentRun(mcp, { request: "Apply update without token", mode: "apply_update", operationId: preview.operationId }, agentOutputSchema);
    assert(missingToken.status === "NEEDS_INPUT", "apply_update without confirmationToken should be blocked");

    const applied = await agentRun(mcp, {
      request: "Apply update",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    }, agentOutputSchema);
    assert(applied.status === "SUCCESS", "apply_update with confirmationToken should succeed");

    const autoApplied = await agentRun(mcp, {
      request: "Change Data C2 to 321",
      workbookContextId: prepared.workbookContextId,
      target: { sheetName: "Data", range: "C2" },
      values: { values: [[321]] }
    }, agentOutputSchema);
    assert(autoApplied.status === "PREVIEW_READY", "auto mode should preview scoped value edits when auto-apply is disabled");
    assert(autoApplied.mode === "auto", "auto preview result should preserve auto mode");
    assert(autoApplied.confirmationToken, "auto preview result should expose a confirmation token");
    assert(autoApplied.telemetry?.autoApplied !== true, "auto preview result should not report telemetry.autoApplied");
    assert(autoApplied.telemetry?.safetyDecision === "manual_review:auto_apply_disabled", "auto preview result should report disabled auto-apply decision");

    const formulaBlocked = await agentRun(mcp, {
      request: "Fix formula in Report A10",
      workbookContextId: prepared.workbookContextId,
      target: { sheetName: "Report", range: "A10" },
      values: { values: [[100]] }
    }, agentOutputSchema);
    assert(formulaBlocked.status === "NEEDS_INPUT", "auto mode should not treat formula repair as a value write");
    assert(formulaBlocked.telemetry?.safetyDecision === "manual_review:advanced_workflow", "formula-sensitive auto request should report manual review decision");

    const validated = await agentRun(mcp, { request: "Validate workbook", mode: "validate" }, agentOutputSchema);
    assert(validated.status === "SUCCESS", "validate should succeed after apply");

    writeArtifact("e2e-agent-workflow-transcript.json", { transcript, fakeAddin: addin.summary() });
    console.log(`E2E agent workflow passed. Artifacts: ${artifactsDir}`);
  } catch (error) {
    writeArtifact("e2e-agent-workflow-failure.json", {
      error: error instanceof Error ? { message: error.message, stack: error.stack } : String(error),
      transcript,
      serverStderr,
      fakeAddin: addin?.summary()
    });
    console.error(`E2E agent workflow failed. Artifacts: ${artifactsDir}`);
    throw error;
  } finally {
    addin?.close();
    mcp.close();
    server.kill();
  }
}

async function agentRun(client, args, outputSchema) {
  return callTool(client, "excel.agent.run", args, outputSchema);
}

async function callTool(client, name, args, outputSchema) {
  const result = await client.request("tools/call", { name, arguments: args });
  if (result.isError) {
    throw new Error(`${name} returned MCP error: ${JSON.stringify(result)}`);
  }
  assert(result.structuredContent, `${name} should return structuredContent for strict MCP clients`);
  assertTelemetryKeysDeclared(result.structuredContent, outputSchema);
  const text = result.content?.find((item) => item.type === "text")?.text;
  assert(text, `${name} returned no text content`);
  const parsed = JSON.parse(text);
  assert(JSON.stringify(parsed) === JSON.stringify(result.structuredContent), `${name} text content should mirror structuredContent`);
  transcript.push({ tool: name, args, telemetry: parsed.telemetry, result: parsed });
  return parsed;
}

function assertTelemetryKeysDeclared(structuredContent, outputSchema) {
  const telemetryProperties = outputSchema?.properties?.telemetry?.properties;
  assert(telemetryProperties, "excel.agent.run should publish telemetry output schema");
  for (const key of Object.keys(structuredContent.telemetry ?? {})) {
    assert(key in telemetryProperties, `structured telemetry key ${key} must be declared in outputSchema`);
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
      if (!parsed) return;
      this.handle(parsed);
    }
  }

  readFramed() {
    const marker = this.buffer.indexOf("\r\n\r\n");
    if (marker < 0) return undefined;
    const header = this.buffer.slice(0, marker).toString("utf8");
    const match = /^Content-Length:\s*(\d+)/im.exec(header);
    if (!match) return undefined;
    const length = Number(match[1]);
    const bodyStart = marker + 4;
    const bodyEnd = bodyStart + length;
    if (this.buffer.length < bodyEnd) return undefined;
    const body = this.buffer.slice(bodyStart, bodyEnd).toString("utf8");
    this.buffer = this.buffer.slice(bodyEnd);
    return JSON.parse(body);
  }

  readLineDelimited() {
    const newline = this.buffer.indexOf("\n");
    if (newline < 0) return undefined;
    const line = this.buffer.slice(0, newline).toString("utf8").trim();
    this.buffer = this.buffer.slice(newline + 1);
    return line ? JSON.parse(line) : undefined;
  }

  handle(message) {
    if (!("id" in message) || "method" in message) return;
    const pending = this.pending.get(message.id);
    if (!pending) return;
    clearTimeout(pending.timeout);
    this.pending.delete(message.id);
    if (message.error) pending.reject(new Error(message.error.message));
    else pending.resolve(message.result);
  }

  close() {
    this.child.stdin.destroy();
  }
}

class FakeAddin {
  static async connect(url, workbook) {
    const { WebSocket } = await import("../apps/backend/node_modules/ws/wrapper.mjs");
    const socket = new WebSocket(url);
    const addin = new FakeAddin(socket, workbook);
    await new Promise((resolve, reject) => {
      socket.once("open", resolve);
      socket.once("error", reject);
    });
    socket.on("message", (raw) => addin.onMessage(JSON.parse(String(raw))));
    await addin.waitForConnectionId();
    addin.sendNotification("addin.hello", {
      capabilities: { platform: "mac", officeVersion: "fake-agent-workflow", apiSets: { ExcelApi: "1.16" }, features: { ranges: "supported", tables: "supported", formulas: "supported" } },
      activeWorkbook: workbook.ref
    });
    return addin;
  }

  constructor(socket, workbook) {
    this.socket = socket;
    this.workbook = workbook;
    this.connectionId = undefined;
    this.calls = [];
    this.connectedResolvers = [];
  }

  waitForConnectionId() {
    return this.connectionId ? Promise.resolve() : new Promise((resolve) => this.connectedResolvers.push(resolve));
  }

  onMessage(message) {
    if (message.method === "backend.connected") {
      this.connectionId = message.params.connectionId;
      for (const resolve of this.connectedResolvers.splice(0)) resolve();
      return;
    }
    if (!("id" in message) || !message.method) return;
    Promise.resolve()
      .then(() => this.handleRequest(message.method, message.params ?? {}))
      .then((result) => this.socket.send(JSON.stringify({ jsonrpc: "2.0", id: message.id, result })))
      .catch((error) => this.socket.send(JSON.stringify({ jsonrpc: "2.0", id: message.id, error: { code: -32000, message: error instanceof Error ? error.message : String(error) } })));
  }

  handleRequest(method, params) {
    this.calls.push({ method, params });
    switch (method) {
      case "runtime.ping":
        return { ok: true, at: params.at };
      case "runtime.get_active_context":
        return this.workbook.ref;
      case "workbook.get_map":
        return this.workbook.getMap();
      case "workbook.snapshot_ranges":
        return this.workbook.snapshotRanges(params.ranges);
      case "table.list":
        return { ok: true, tables: [...this.workbook.tables.values()].map((table) => table.info()) };
      case "table.get_info":
        return { ok: true, info: this.workbook.table(params.tableName).info() };
      case "table.append_rows":
        return this.workbook.table(params.tableName).appendRows(params.values);
      case "names.list":
        return { ok: true, names: [...this.workbook.names.values()] };
      case "operation.execute_batch":
        return this.workbook.executeBatch(params.request);
      case "range.find_errors":
        return { ok: true, data: { isNullObject: true, cellCount: 0, areas: [] }, warnings: [] };
      default:
        return { ok: true, warnings: [] };
    }
  }

  sendNotification(method, params) {
    this.socket.send(JSON.stringify({ jsonrpc: "2.0", method, params }));
  }

  close() {
    this.socket.close();
  }

  summary() {
    return { connectionId: this.connectionId, calls: this.calls.map((call) => call.method), workbook: this.workbook.getMap() };
  }
}

class FakeWorkbook {
  constructor(id) {
    this.ref = { workbookId: id, name: "Agent Workflow.xlsx", path: path.join(tempRoot, "Agent Workflow.xlsx"), platform: "mac" };
    this.sheets = new Map();
    this.tables = new Map();
    this.names = new Map();
  }

  addSheet(name) {
    const sheet = new FakeSheet(this.ref.workbookId, name);
    this.sheets.set(name, sheet);
    return sheet;
  }

  sheet(name) {
    const sheet = this.sheets.get(name);
    if (!sheet) throw new Error(`Unknown fake sheet: ${name}`);
    return sheet;
  }

  table(name) {
    const table = this.tables.get(name);
    if (!table) throw new Error(`Unknown fake table: ${name}`);
    return table;
  }

  createTable({ sheetName, address, tableName }) {
    const table = new FakeTable(this, tableName, sheetName, address);
    this.tables.set(tableName, table);
    return table;
  }

  getMap() {
    return {
      workbook: this.ref,
      activeSheet: "Data",
      sheets: [...this.sheets.values()].map((sheet) => ({
        workbookId: this.ref.workbookId,
        worksheetId: `sheet_${sheet.name}`,
        name: sheet.name,
        usedRange: sheet.usedRange(),
        tables: [...this.tables.values()].filter((table) => table.sheetName === sheet.name).map((table) => table.info())
      }))
    };
  }

  executeBatch(request) {
    const readData = [];
    let cellsRead = 0;
    let cellsWritten = 0;
    for (const operation of request.operations) {
      if (operation.kind === "range.read_full") {
        const snapshot = this.sheet(operation.target.sheetName).snapshot(operation.target);
        cellsRead += snapshot.fingerprint.cellCount;
        readData.push({ operationId: operation.operationId, snapshot });
      }
      if (operation.kind === "range.write_values") {
        cellsWritten += this.sheet(operation.target.sheetName).writeValues(operation.target.address, operation.values);
      }
    }
    return {
      ok: true,
      rollbackAvailable: request.mode === "apply",
      backups: [],
      warnings: [],
      readData,
      diffSummary: { title: "Agent workflow fake batch", changedRanges: [], cellsChanged: cellsWritten, formulasChanged: 0, stylesChanged: 0, tablesChanged: 0, sheetsChanged: 0, destructiveLevel: cellsWritten > 0 ? "values" : "none" },
      telemetry: { cellsRead, cellsWritten, syncCount: 1, rangeCount: request.operations.length, chunkCount: 1, warningCount: 0 }
    };
  }

  snapshotRanges(ranges) {
    return {
      workbookId: this.ref.workbookId,
      capturedAt: fixedNow(),
      workbookFingerprint: {
        workbookId: this.ref.workbookId,
        workbookHash: stableHash(this.getMap()),
        structureHash: stableHash([...this.sheets.keys(), ...this.tables.keys()]),
        capturedAt: fixedNow()
      },
      rangeSnapshots: ranges.map((range) => this.sheet(range.sheetName).snapshot(range))
    };
  }
}

class FakeSheet {
  constructor(workbookId, name) {
    this.workbookId = workbookId;
    this.name = name;
    this.cells = new Map();
  }

  writeValues(address, values) {
    const range = parseRange(address);
    forEachMatrix(values, (value, row, col) => {
      this.cell(range.startRow + row, range.startCol + col).value = value;
    });
    return matrixCellCount(values);
  }

  writeFormulas(address, formulas) {
    const range = parseRange(address);
    forEachMatrix(formulas, (formula, row, col) => {
      this.cell(range.startRow + row, range.startCol + col).formula = formula;
    });
    return matrixCellCount(formulas);
  }

  snapshot(rangeRef) {
    const range = parseRange(rangeRef.address);
    const values = [];
    const formulas = [];
    const text = [];
    for (let row = range.startRow; row <= range.endRow; row += 1) {
      const valueRow = [];
      const formulaRow = [];
      const textRow = [];
      for (let col = range.startCol; col <= range.endCol; col += 1) {
        const cell = this.cell(row, col);
        valueRow.push(cell.value);
        formulaRow.push(cell.formula);
        textRow.push(cell.formula ?? (cell.value === null || cell.value === undefined ? "" : String(cell.value)));
      }
      values.push(valueRow);
      formulas.push(formulaRow);
      text.push(textRow);
    }
    return {
      fingerprint: { range: rangeRef, hash: stableHash({ values, formulas }), cellCount: range.rowCount * range.columnCount, capturedAt: fixedNow() },
      values,
      formulas,
      text
    };
  }

  usedRange() {
    let maxRow = 1;
    let maxCol = 1;
    for (const key of this.cells.keys()) {
      const [row, col] = key.split(":").map(Number);
      maxRow = Math.max(maxRow, row);
      maxCol = Math.max(maxCol, col);
    }
    return { workbookId: this.workbookId, sheetName: this.name, address: `A1:${columnName(maxCol)}${maxRow}`, rowCount: maxRow, columnCount: maxCol };
  }

  cell(row, col) {
    const key = `${row}:${col}`;
    if (!this.cells.has(key)) this.cells.set(key, { value: null, formula: null });
    return this.cells.get(key);
  }
}

class FakeTable {
  constructor(workbook, tableName, sheetName, address) {
    this.workbook = workbook;
    this.tableName = tableName;
    this.sheetName = sheetName;
    this.address = address;
  }

  info() {
    const range = parseRange(this.address);
    const headers = this.workbook.sheet(this.sheetName).snapshot({ workbookId: this.workbook.ref.workbookId, sheetName: this.sheetName, address: `${columnName(range.startCol)}${range.startRow}:${columnName(range.endCol)}${range.startRow}` }).values[0].map(String);
    return {
      workbookId: this.workbook.ref.workbookId,
      tableName: this.tableName,
      name: this.tableName,
      id: `table_${this.tableName}`,
      sheetName: this.sheetName,
      address: this.address,
      headerAddress: `${columnName(range.startCol)}${range.startRow}:${columnName(range.endCol)}${range.startRow}`,
      rowCount: Math.max(0, range.rowCount - 1),
      columnCount: range.columnCount,
      columns: headers.map((name, index) => ({ id: index + 1, index, name }))
    };
  }

  appendRows(values) {
    const range = parseRange(this.address);
    const startAddress = `${columnName(range.startCol)}${range.endRow + 1}`;
    this.workbook.sheet(this.sheetName).writeValues(startAddress, values);
    this.address = `${columnName(range.startCol)}${range.startRow}:${columnName(range.endCol)}${range.endRow + values.length}`;
    return { ok: true, rowCount: values.length, tableName: this.tableName, address: this.address };
  }
}

function createWorkbookFixture(id) {
  const workbook = new FakeWorkbook(id);
  workbook.addSheet("Data").writeValues("A1:D4", [
    ["Date", "Account", "Amount", "Status"],
    ["2026-01-01", "A-100", 100, "Open"],
    ["2026-01-02", "A-200", 200, "Closed"],
    ["2026-01-03", "A-300", 300, "Open"]
  ]);
  workbook.addSheet("Report").writeValues("A1:B3", [["Metric", "Value"], ["Revenue", 1000], ["Total", 1000]]);
  workbook.sheet("Report").writeFormulas("A10:A12", [["=SUM(B1:B3)"], ["=A10"], ["=A11"]]);
  workbook.addSheet("Apr 2026").writeValues("A1:AE3", [
    ["Transaction Date", "Job ID", "Truck ID", "Description", "Transaction Type", "Direction", "Cash Amount", "Actual Amount", "Payment Variance", "Reconciliation Note", "Transfer From/To", "Proof File", "Detail Notes", "", "Invoice No", "Job ID", "Invoice Date", "Billed To", "Booking No", "Customer", "Job", "Container No", "Container Size", "Job Price", "Lifting On", "Lifting Off", "Total Lifting", "Other Fees", "Gross Billed", "W/H Tax", "Net Collect"],
    ["2026-04-01", "204", "71-4653", "Company gas top-up", "company_gas_topup", "Outflow", "2211.21", "2211.21", "0", "", "Bank", "proof.pdf", "text note", "", "INV-001", "204", "2026-04-01", "ACME", "BK-001", "Customer A", "งาน 204", "CONT-1", "20GP", "10000", "1000", "1000", "2000", "0", "12000", "360", "11640"],
    ["", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", ""]
  ]);
  workbook.addSheet("Financials - June 2026").writeValues("A1:C4", [["Metric", "Jun 2026", "Variance"], ["Revenue", 1200, 50], ["Expense", 700, -20], ["Profit", 500, 70]]);
  workbook.addSheet("Financials - May 2026").writeValues("A1:C4", [["Metric", "May 2026", "Variance"], ["Revenue", 1100, 30], ["Expense", 680, 10], ["Profit", 420, 20]]);
  workbook.createTable({ sheetName: "Data", address: "A1:D4", tableName: "Transactions" });
  workbook.names.set(":RevenueTotal", { workbookId: id, name: "RevenueTotal", scope: "workbook", address: "Report!B2" });
  return workbook;
}

function parseRange(address) {
  const [start, end = start] = address.replace(/\$/g, "").split(":");
  const startCell = parseCell(start);
  const endCell = parseCell(end);
  return { startRow: startCell.row, startCol: startCell.col, endRow: endCell.row, endCol: endCell.col, rowCount: endCell.row - startCell.row + 1, columnCount: endCell.col - startCell.col + 1 };
}

function parseCell(cell) {
  const match = /^([A-Z]+)(\d+)$/i.exec(cell);
  if (!match) throw new Error(`Unsupported A1 cell: ${cell}`);
  return { col: columnIndex(match[1].toUpperCase()), row: Number(match[2]) };
}

function columnIndex(name) {
  let value = 0;
  for (const char of name) value = value * 26 + char.charCodeAt(0) - 64;
  return value;
}

function columnName(index) {
  let value = "";
  let remaining = index;
  while (remaining > 0) {
    const mod = (remaining - 1) % 26;
    value = String.fromCharCode(65 + mod) + value;
    remaining = Math.floor((remaining - mod) / 26);
  }
  return value || "A";
}

function forEachMatrix(matrix, callback) {
  for (let row = 0; row < matrix.length; row += 1) {
    for (let col = 0; col < matrix[row].length; col += 1) callback(matrix[row][col], row, col);
  }
}

function matrixCellCount(matrix) {
  return matrix.reduce((total, row) => total + row.length, 0);
}

function stableHash(value) {
  return createHash("sha256").update(JSON.stringify(value)).digest("hex");
}

function fixedNow() {
  return "2026-01-01T00:00:00.000Z";
}

async function waitForHttp(url, timeoutMs) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      const response = await fetch(url);
      if (response.ok) return;
    } catch {
      // keep polling
    }
    await new Promise((resolve) => setTimeout(resolve, 100));
  }
  throw new Error(`Timed out waiting for ${url}`);
}

function writeArtifact(name, value) {
  writeFileSync(path.join(artifactsDir, name), JSON.stringify(value, null, 2));
}

function assert(condition, message) {
  if (!condition) throw new Error(message);
}

main().catch((error) => {
  console.error(error);
  process.exit(1);
});
