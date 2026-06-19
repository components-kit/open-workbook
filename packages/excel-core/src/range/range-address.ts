import type { A1Range, WorkbookId } from "@components-kit/open-workbook-protocol";

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

export function tryParseA1Address(address: string): ParsedA1Address | undefined {
  try {
    return parseA1Address(address);
  } catch {
    return undefined;
  }
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

export function cellCountFromAddress(address: string): number | undefined {
  const parsed = tryParseA1Address(stripSheetName(address));
  return parsed ? (parsed.endRow - parsed.startRow + 1) * (parsed.endColumn - parsed.startColumn + 1) : undefined;
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

export function formatA1Cell(row: number, column: number): string {
  if (!Number.isInteger(row) || row < 1) {
    throw new Error(`Invalid row number: ${row}`);
  }
  return `${numberToColumnName(column)}${row}`;
}

export function formatA1Address(parsed: ParsedA1Address): string {
  const start = formatA1Cell(parsed.startRow, parsed.startColumn);
  const end = formatA1Cell(parsed.endRow, parsed.endColumn);
  const range = start === end ? start : `${start}:${end}`;
  return parsed.sheetName ? `${quoteSheetName(parsed.sheetName)}!${range}` : range;
}

export function stripSheetName(address: string): string {
  const bang = address.lastIndexOf("!");
  return bang >= 0 ? address.slice(bang + 1) : address;
}

export function rangesOverlap(leftAddress: string, rightAddress: string): boolean {
  const left = tryParseA1Address(stripSheetName(leftAddress));
  const right = tryParseA1Address(stripSheetName(rightAddress));
  if (!left || !right) {
    return false;
  }
  return left.startRow <= right.endRow
    && left.endRow >= right.startRow
    && left.startColumn <= right.endColumn
    && left.endColumn >= right.startColumn;
}

export function unquoteSheetName(sheetName: string | undefined): string | undefined {
  if (!sheetName) {
    return undefined;
  }
  if (sheetName.startsWith("'") && sheetName.endsWith("'")) {
    return sheetName.slice(1, -1).replace(/''/g, "'");
  }
  return sheetName;
}

export function quoteSheetName(sheetName: string): string {
  return /^[A-Za-z_][A-Za-z0-9_]*$/.test(sheetName) ? sheetName : `'${sheetName.replace(/'/g, "''")}'`;
}
