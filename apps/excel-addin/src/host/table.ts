import {
  appendTableRows,
  applyTableFilters,
  clearTableDataKeepFormulas,
  clearTableFilters,
  clearTableSort,
  copyTableStructure,
  createTable,
  getTableInfo,
  listTables,
  readTable,
  reorderTableColumns,
  resizeTable,
  setTableStyle,
  setTableTotalRow,
  sortTable,
  updateTableRows
} from "./executor-core.js";

export const tableHostOperations = {
  listTables: (params: unknown) => listTables((params as { workbookId: string }).workbookId),
  getTableInfo: (params: unknown) => getTableInfo(params as Parameters<typeof getTableInfo>[0]),
  readTable: (params: unknown) => readTable(params as Parameters<typeof readTable>[0]),
  createTable: (params: unknown) => createTable(params as Parameters<typeof createTable>[0]),
  resizeTable: (params: unknown) => resizeTable(params as Parameters<typeof resizeTable>[0]),
  reorderTableColumns: (params: unknown) => reorderTableColumns(params as Parameters<typeof reorderTableColumns>[0]),
  appendTableRows: (params: unknown) => appendTableRows(params as Parameters<typeof appendTableRows>[0]),
  updateTableRows: (params: unknown) => updateTableRows(params as Parameters<typeof updateTableRows>[0]),
  clearTableDataKeepFormulas: (params: unknown) => clearTableDataKeepFormulas(params as Parameters<typeof clearTableDataKeepFormulas>[0]),
  clearTableFilters: (params: unknown) => clearTableFilters(params as Parameters<typeof clearTableFilters>[0]),
  applyTableFilters: (params: unknown) => applyTableFilters(params as Parameters<typeof applyTableFilters>[0]),
  sortTable: (params: unknown) => sortTable(params as Parameters<typeof sortTable>[0]),
  clearTableSort: (params: unknown) => clearTableSort(params as Parameters<typeof clearTableSort>[0]),
  setTableTotalRow: (params: unknown) => setTableTotalRow(params as Parameters<typeof setTableTotalRow>[0]),
  setTableStyle: (params: unknown) => setTableStyle(params as Parameters<typeof setTableStyle>[0]),
  copyTableStructure: (params: unknown) => copyTableStructure(params as Parameters<typeof copyTableStructure>[0])
};
