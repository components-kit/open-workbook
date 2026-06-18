import type { BatchRequest, WorkbookId } from "@components-kit/open-workbook-protocol";
import { createMetadataFingerprint, type WorkbookMetadata } from "./workbook-metadata-cache.js";

export const workbookId = "workbook_agent_unit" as WorkbookId;
export const activeWorkbook = { workbookId, name: "Agent Unit.xlsx", path: "/tmp/Agent Unit.xlsx" };
export const sheets = [
  { workbookId, worksheetId: "sheet_Data", name: "Data", usedRange: { address: "A1:D4", rowCount: 4, columnCount: 4 }, tables: [{ name: "Transactions" }] },
  { workbookId, worksheetId: "sheet_Report", name: "Report", usedRange: { address: "A1:B20", rowCount: 20, columnCount: 2 }, tables: [] },
  { workbookId, worksheetId: "sheet_Operations", name: "Operations", usedRange: { address: "A1:J24", rowCount: 24, columnCount: 10 }, tables: [] },
  { workbookId, worksheetId: "sheet_Mar", name: "Mar 2026", usedRange: { address: "A1:AJ206", rowCount: 206, columnCount: 36 }, tables: [] },
  { workbookId, worksheetId: "sheet_Apr", name: "Apr 2026", usedRange: { address: "A1:AJ244", rowCount: 244, columnCount: 36 }, tables: [] },
  { workbookId, worksheetId: "sheet_FinJun", name: "Financials - June 2026", usedRange: { address: "A1:C4", rowCount: 4, columnCount: 3 }, tables: [] },
  { workbookId, worksheetId: "sheet_FinMay", name: "Financials - May 2026", usedRange: { address: "A1:C4", rowCount: 4, columnCount: 3 }, tables: [] }
];

export class FakeAgentRuntime {
  readBatchCount = 0;
  writeBatchCount = 0;
  appendTableRowCount = 0;
  returnDataOnly = false;
  omitOkOnWrite = false;
  lastWriteOperations: Array<Extract<BatchRequest["operations"][number], { target: any }>> = [];
  selection: any;
  readiness: any;

  sessions = { getActive: () => ({ activeWorkbook }) };

  getStatus() {
    return { activeAddinConnected: true, connectionState: "ready", activeWorkbookAvailable: true, activeWorkbook };
  }

  async getConnectionReadiness() {
    if (this.readiness) return this.readiness;
    return { ok: true, connectionState: "ready", activeWorkbook, status: this.getStatus() };
  }

  async getActiveContext() {
    return { ok: true, activeWorkbook };
  }

  async getSelection() {
    return this.selection ? { workbook: activeWorkbook, selection: this.selection } : { ok: false };
  }

  async getWorkbookMap() {
    return { ok: true, map: { workbook: activeWorkbook, activeSheet: "Data", sheets } };
  }

  async listTables() {
    return { ok: true, tables: [{ name: "Transactions" }] };
  }

  async getTableInfo() {
    return {
      ok: true,
      info: {
        tableName: "Transactions",
        sheetName: "Data",
        address: "A1:D4",
        headerAddress: "A1:D1",
        dataRange: "A2:D4",
        columns: [{ name: "Date" }, { name: "Account" }, { name: "Amount" }, { name: "Status" }]
      }
    };
  }

  async listNames() {
    return { ok: true, names: [{ workbookId, name: "RevenueTotal", scope: "workbook", address: "Report!B2" }] };
  }

  listRegions() {
    return { ok: true, regions: [{ name: "InputRegion", sheetName: "Report", address: "B1:B3" }] };
  }

  async applyBatch(request: BatchRequest) {
    if (request.operations.every((operation) => operation.kind === "range.read_full")) {
      this.readBatchCount += 1;
      const readData = request.operations.map((operation) => {
        const sheetName = operation.target.sheetName;
        const address = operation.target.address;
        return {
          operationId: operation.operationId,
          snapshot: {
            values: valuesFor(sheetName, address),
            formulas: formulasFor(sheetName, address),
            text: valuesFor(sheetName, address).map((row) => row.map((value) => value === null || value === undefined ? "" : String(value)))
          }
        };
      });
      return {
        ok: true,
        ...(this.returnDataOnly ? { data: readData } : { readData }),
        telemetry: { cellsRead: request.operations.reduce((total, operation) => total + valuesFor(operation.target.sheetName, operation.target.address).flat().length, 0) }
      };
    }
    this.writeBatchCount += 1;
    this.lastWriteOperations = request.operations.filter((operation): operation is Extract<BatchRequest["operations"][number], { target: any }> => "target" in operation);
    return {
      ...(this.omitOkOnWrite ? {} : { ok: true }),
      backups: [],
      warnings: [],
      telemetry: { cellsWritten: 1 }
    };
  }

  async appendTableRows() {
    this.appendTableRowCount += 1;
    return { ok: true, backups: [], warnings: [], telemetry: { rowsWritten: 1 } };
  }

  async validateWorkbook() {
    return { ok: true, issues: [] };
  }
}

export function selectionInfo(sheetName: string, address: string, position = { row: 2, column: 2 }) {
  return {
    workbookId,
    sheetName,
    address,
    startCell: {
      workbookId,
      sheetName,
      address,
      row: position.row,
      column: position.column,
      rowIndex: position.row - 1,
      columnIndex: position.column - 1
    },
    endCell: {
      workbookId,
      sheetName,
      address,
      row: position.row,
      column: position.column,
      rowIndex: position.row - 1,
      columnIndex: position.column - 1
    },
    rowCount: 1,
    columnCount: 1,
    cellCount: 1,
    isSingleCell: true
  };
}

function valuesFor(sheetName: string, address: string) {
  if (sheetName === "Data") {
    if (address === "A1:J10") {
      return [
        ["Owner", "", "", "", "", "", "", "", "", "Status"],
        ["", "", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", "", ""],
        ["", "", "", "", "", "", "", "", "", "Ready"]
      ];
    }
    if (address === "B2") return [["A-100"]];
    if (address === "B3") return [[""]];
    if (address === "C2") return [[100]];
    if (address === "C1:C4") return [["Amount"], [100], [200], [300]];
    return [
      ["Date", "Account", "Amount", "Status"],
      ["2026-01-01", "A-100", 100, "Open"],
      ["2026-01-02", "A-200", 200, "Closed"],
      ["2026-01-03", "A-300", 300, "Open"]
    ];
  }
  if (sheetName === "Financials - June 2026") {
    if (address === "B2") return [[1200]];
    return [["Metric", "Jun 2026", "Variance"], ["Revenue", 1200, 50], ["Expense", 700, -20], ["Profit", 500, 70]];
  }
  if (sheetName === "Financials - May 2026") {
    return [["Metric", "May 2026", "Variance"], ["Revenue", 1100, 30], ["Expense", 680, 10], ["Profit", 420, 20]];
  }
  if (sheetName === "Mar 2026" || sheetName === "Apr 2026") {
    return monthlySheetValues(sheetName, address);
  }
  if (sheetName === "Operations") {
    const rows = [
      ["Operations Review", "", "", "", "", "", "", "", "", ""],
      ["Period", "Jun 2026", "Prepared by", "Ops", "", "", "", "", "", ""],
      ["", "", "", "", "", "", "", "", "", ""],
      ["Metric", "Value", "Status", "", "", "", "", "Owner", "Action", "Due"],
      ["Revenue", 7200, "On Track", "", "", "", "", "A. Chen", "Review invoices", "2026-06-20"],
      ["Expense", 2800, "Watch", "", "", "", "", "M. Lee", "Check variance", "2026-06-21"],
      ["Profit", 4400, "On Track", "", "", "", "", "", "", ""],
      ["", "", "", "", "", "", "", "", "", ""],
      ["", "", "", "", "", "", "", "", "", ""],
      ["Invoice No", "Customer", "Job", "Amount", "Status", "Owner", "", "", "", ""],
      ["INV-100", "Acme", "Lift", 1200, "Open", "A. Chen", "", "", "", ""],
      ["INV-101", "Northwind", "Haul", 900, "Paid", "M. Lee", "", "", "", ""],
      ["INV-102", "Contoso", "Storage", 450, "Open", "A. Chen", "", "", "", ""],
      ["", "", "", "", "", "", "", "", "", ""],
      ["Reconciliation", "", "", "", "", "", "", "", "", ""],
      ["Expected", "Actual", "Variance", "Check", "", "", "", "", "", ""],
      [2500, 2550, "=B17-A17", "Review", "", "", "", "", "", ""]
    ];
    if (address === "A10:F13") {
      return rows.slice(9, 13).map((row) => row.slice(0, 6));
    }
    return rows;
  }
  if (address === "B1") return [["input"]];
  if (address === "A12") return [["=SUM(B1:B10)"]];
  return [
    ["Metric", "Value"],
    ["Revenue", 1000],
    ["Total", 1000]
  ];
}

function monthlySheetValues(sheetName: string, address: string) {
  const month = sheetName.startsWith("Mar") ? "March" : "April";
  const cashReceived = sheetName.startsWith("Mar") ? 280000 : 333881.72;
  const cashSpent = sheetName.startsWith("Mar") ? 240000 : 363263.96;
  const profit = cashReceived - cashSpent;
  const summaryRows: unknown[][] = [
    [`${month} 2026 Summary`, "", "", ""],
    ["", "", "", ""],
    ["Cash Summary", "", "", ""],
    ["Metric", "Amount (THB)", "View", "Notes"],
    [`Cash received in ${month}`, cashReceived, "Cash in", `Actual cash/bank inflows received in ${month}.`],
    [`Cash spent in ${month}`, cashSpent, "Cash out", `Actual cash/bank outflows paid in ${month}.`],
    ["Net cash movement", profit, "Cash net", "Cash received less cash spent."],
    ["", "", "", ""],
    ["P&L / Credit Float", "", "", ""],
    ["Billed / earned transport revenue", cashReceived + 45000, "Revenue", "Invoice revenue for jobs in the month."],
    ["Operating spend", cashSpent, "Expense", "Cash operating spend."],
    ["Profit / loss", profit, "Profit", "Revenue less spend."],
    ["", "", "", ""],
    ["Spend Breakdown", "", "", ""],
    ["Company gas top-up", sheetName.startsWith("Mar") ? 58000 : 71000, "Expense", "Fuel-related spend."],
    ["Truck repair", sheetName.startsWith("Mar") ? 18000 : 42000, "Expense", "Maintenance spend."],
    ["Driver salary", sheetName.startsWith("Mar") ? 64000 : 65000, "Expense", "Payroll spend."],
    ["", "", "", ""],
    ["Checks / Interpretation", "", "", ""],
    ["Management takeaway", profit >= 0 ? "Profitable" : "Loss", "Status", profit >= 0 ? "Month remained profitable." : "Cash out exceeded cash in."]
  ];
  if (address.startsWith("AG")) {
    return summaryRows;
  }
  const row1 = padToSummary([
    "Transaction Date", "Job ID", "Truck ID", "Description", "Transaction Type", "Direction", "Cash Amount", "Actual Amount", "Payment Variance", "Reconciliation Note", "Transfer From/To", "Proof File", "Detail Notes", "",
    "Invoice No", "Job ID", "Invoice Date", "Billed To", "Booking No", "Customer", "Job", "Container No", "Container Size", "Job Price", "Lifting On", "Lifting Off", "Total Lifting", "Other Fees", "Gross Billed", "W/H Tax", "Net Collect"
  ], summaryRows[0]!);
  const row2 = padToSummary([
    sheetName.startsWith("Mar") ? "2026-03-01" : "2026-04-01", "204", "71-4653", "Company gas top-up", "company_gas_topup", "Outflow", "2211.21", "2211.21", "0", "", "Bank", "proof.pdf", "text note", "",
    "INV-001", "204", sheetName.startsWith("Mar") ? "2026-03-01" : "2026-04-01", "ACME", "BK-001", "Customer A", "Job 204", "CONT-1", "20GP", "10000", "1000", "1000", "2000", "0", "12000", "360", "11640"
  ], summaryRows[1]!);
  const rows = [row1, row2, ...summaryRows.slice(2).map((summary) => padToSummary([], summary))];
  if (address.startsWith("O")) {
    return rows.map((row) => row.slice(14, 31));
  }
  return rows;
}

function padToSummary(left: unknown[], summary: unknown[]) {
  const row: unknown[] = Array.from({ length: 36 }, (_cell, index) => left[index] ?? "");
  summary.forEach((value, index) => {
    row[32 + index] = value;
  });
  return row;
}

function formulasFor(sheetName: string, address: string) {
  const values = valuesFor(sheetName, address);
  return values.map((row) => row.map((value) => typeof value === "string" && value.startsWith("=") ? value : null));
}

export function createCachedMetadata(workbookContextId: string): WorkbookMetadata {
  const fingerprint = createMetadataFingerprint({ workbookId, workbook: activeWorkbook, sheets });
  return {
    workbookContextId,
    workbookKey: `id:${workbookId}`,
    detailLevel: "sampled",
    workbook: { workbookId, name: activeWorkbook.name, path: activeWorkbook.path, sheetCount: sheets.length, activeSheet: "Data" },
    sheets: [
      { id: "sheet:0", name: "Data", index: 0, usedRange: "A1:D4", rowCount: 4, columnCount: 4, kind: "transaction", headers: [], tableIds: ["table:Transactions"], sectionIds: [], summaryBlockIds: [], formulaRegionIds: [] },
      { id: "sheet:1", name: "Report", index: 1, usedRange: "A1:B20", rowCount: 20, columnCount: 2, kind: "summary", headers: [], tableIds: [], sectionIds: [], summaryBlockIds: [], formulaRegionIds: ["formula:manual"] }
    ],
    tables: [{ id: "table:Transactions", sheetName: "Data", name: "Transactions", range: "A1:D4", columns: [] }],
    namedRanges: [],
    sections: [],
    summaryBlocks: [],
    formulaRegions: [],
    fingerprint,
    createdAt: Date.now(),
    updatedAt: Date.now(),
    expiresAt: Date.now() + 60_000
  };
}
