export function matrixCellCount(values: unknown[][]): number {
  return values.reduce((sum, row) => sum + row.length, 0);
}

export function cloneMatrix<T = unknown>(values: T[][]): T[][] {
  return values.map((row) => [...row]);
}

export function rectangularize<T = unknown>(values: T[][], fillValue: T | null = null): Array<Array<T | null>> {
  const width = values.reduce((max, row) => Math.max(max, row.length), 0);
  return values.map((row) => [...row, ...filledCells(width - row.length, fillValue)]);
}

export function padMatrixRows<T = unknown>(values: T[][], rowCount: number, columnCount: number, fillValue: T | null = null): Array<Array<T | null>> {
  const output = values.map((row) => [...row, ...filledCells(Math.max(0, columnCount - row.length), fillValue)]);
  while (output.length < rowCount) {
    output.push(filledCells(columnCount, fillValue));
  }
  return output;
}

function filledCells<T>(count: number, value: T | null): Array<T | null> {
  return Array.from({ length: count }, () => value);
}
