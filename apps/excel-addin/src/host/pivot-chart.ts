import {
  copyChartFromTemplate,
  copyPivotTableFromTemplate,
  createChart,
  createPivotTable,
  deleteChart,
  deletePivotTable,
  getChartInfo,
  getPivotTableInfo,
  listCharts,
  listPivotTables,
  refreshAllPivotTables,
  refreshChart,
  refreshPivotTable,
  updateChartDataSource
} from "./executor-core.js";

export const pivotChartHostOperations = {
  listPivotTables: (params: unknown) => listPivotTables((params as { workbookId: string }).workbookId),
  getPivotTableInfo: (params: unknown) => getPivotTableInfo(params as Parameters<typeof getPivotTableInfo>[0]),
  createPivotTable: (params: unknown) => createPivotTable(params as Parameters<typeof createPivotTable>[0]),
  refreshPivotTable: (params: unknown) => refreshPivotTable(params as Parameters<typeof refreshPivotTable>[0]),
  refreshAllPivotTables: (params: unknown) => refreshAllPivotTables((params as { workbookId: string }).workbookId),
  copyPivotTableFromTemplate: (params: unknown) => copyPivotTableFromTemplate(params as Parameters<typeof copyPivotTableFromTemplate>[0]),
  deletePivotTable: (params: unknown) => deletePivotTable(params as Parameters<typeof deletePivotTable>[0]),
  listCharts: (params: unknown) => listCharts((params as { workbookId: string }).workbookId),
  getChartInfo: (params: unknown) => getChartInfo(params as Parameters<typeof getChartInfo>[0]),
  createChart: (params: unknown) => createChart(params as Parameters<typeof createChart>[0]),
  updateChartDataSource: (params: unknown) => updateChartDataSource(params as Parameters<typeof updateChartDataSource>[0]),
  copyChartFromTemplate: (params: unknown) => copyChartFromTemplate(params as Parameters<typeof copyChartFromTemplate>[0]),
  refreshChart: (params: unknown) => refreshChart(params as Parameters<typeof refreshChart>[0]),
  deleteChart: (params: unknown) => deleteChart(params as Parameters<typeof deleteChart>[0])
};
