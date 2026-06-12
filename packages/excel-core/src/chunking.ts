export interface MatrixRowChunk<T = unknown> {
  rowOffset: number;
  rows: T[][];
}

export function chunkMatrixRows<T>(matrix: T[][], maxCellsPerChunk: number): Array<MatrixRowChunk<T>> {
  if (matrix.length === 0) {
    return [];
  }
  const columnCount = matrix[0]?.length ?? 0;
  if (columnCount === 0 || matrix.some((row) => row.length !== columnCount)) {
    throw new Error("Matrix must be rectangular and contain at least one column.");
  }
  const rowChunkSize = Math.max(1, Math.floor(maxCellsPerChunk / columnCount));
  const chunks: Array<MatrixRowChunk<T>> = [];
  for (let rowOffset = 0; rowOffset < matrix.length; rowOffset += rowChunkSize) {
    chunks.push({
      rowOffset,
      rows: matrix.slice(rowOffset, rowOffset + rowChunkSize)
    });
  }
  return chunks;
}
