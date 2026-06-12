import type { A1Range, WorkbookId } from "@open-workbook/protocol";

export interface ParsedA1Address {
  sheetName?: string;
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}

const A1_RE = /^(?:(?<sheet>(?:'[^']+'|[^!]+))!)?(?<startCol>[A-Z]+)(?<startRow>\d+)(?::(?<endCol>[A-Z]+)(?<endRow>\d+))?$/i;

export function parseA1Address(address: string): ParsedA1Address {
  const match = A1_RE.exec(address.trim());
  if (!match?.groups) {
    throw new Error(`Invalid A1 address: ${address}`);
  }

  const { startCol, startRow: startRowText, endCol, endRow: endRowText, sheet } = match.groups;
  if (!startCol || !startRowText) {
    throw new Error(`Invalid A1 address: ${address}`);
  }

  const startColumn = columnNameToNumber(startCol);
  const startRow = Number(startRowText);
  const endColumn = columnNameToNumber(endCol ?? startCol);
  const endRow = Number(endRowText ?? startRowText);

  if (startRow < 1 || endRow < startRow || endColumn < startColumn) {
    throw new Error(`Invalid A1 range bounds: ${address}`);
  }

  const parsed: ParsedA1Address = {
    startRow,
    startColumn,
    endRow,
    endColumn
  };
  const sheetName = unquoteSheetName(sheet);
  if (sheetName !== undefined) {
    parsed.sheetName = sheetName;
  }
  return parsed;
}

export function normalizeA1Range(workbookId: WorkbookId, sheetName: string, address: string): A1Range {
  const parsed = parseA1Address(address);
  return {
    workbookId,
    sheetName: parsed.sheetName ?? sheetName,
    address: formatA1Address(parsed)
  };
}

export function cellCount(address: string): number {
  const parsed = parseA1Address(address);
  return (parsed.endRow - parsed.startRow + 1) * (parsed.endColumn - parsed.startColumn + 1);
}

export function columnNameToNumber(columnName: string): number {
  let value = 0;
  for (const char of columnName.toUpperCase()) {
    const code = char.charCodeAt(0);
    if (code < 65 || code > 90) {
      throw new Error(`Invalid column name: ${columnName}`);
    }
    value = value * 26 + (code - 64);
  }
  return value;
}

export function numberToColumnName(columnNumber: number): string {
  if (!Number.isInteger(columnNumber) || columnNumber < 1) {
    throw new Error(`Invalid column number: ${columnNumber}`);
  }

  let n = columnNumber;
  let value = "";
  while (n > 0) {
    n -= 1;
    value = String.fromCharCode(65 + (n % 26)) + value;
    n = Math.floor(n / 26);
  }
  return value;
}

export function formatA1Address(parsed: ParsedA1Address): string {
  const start = `${numberToColumnName(parsed.startColumn)}${parsed.startRow}`;
  const end = `${numberToColumnName(parsed.endColumn)}${parsed.endRow}`;
  const range = start === end ? start : `${start}:${end}`;
  return parsed.sheetName ? `${quoteSheetName(parsed.sheetName)}!${range}` : range;
}

function unquoteSheetName(sheetName: string | undefined): string | undefined {
  if (!sheetName) {
    return undefined;
  }
  if (sheetName.startsWith("'") && sheetName.endsWith("'")) {
    return sheetName.slice(1, -1).replace(/''/g, "'");
  }
  return sheetName;
}

function quoteSheetName(sheetName: string): string {
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(sheetName) ? sheetName : `'${sheetName.replace(/'/g, "''")}'`;
}
