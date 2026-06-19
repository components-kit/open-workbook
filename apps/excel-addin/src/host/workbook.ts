import {
  calculateWorkbook,
  closeWorkbook,
  embedWorkbookLocalConfig,
  exportWorkbookFile,
  getWorkbookInfo,
  getWorkbookMap,
  readWorkbookEmbeddedLocalConfig,
  saveWorkbook,
  snapshotRanges
} from "./executor-core.js";

export const workbookHostOperations = {
  getWorkbookInfo,
  getWorkbookMap,
  calculateWorkbook: (params: unknown) => calculateWorkbook((params as { calculationType?: "full" | "recalculate" }).calculationType),
  saveWorkbook,
  exportWorkbookFile: (params: unknown) => exportWorkbookFile((params as { workbookId: string; sliceSize?: number }).workbookId, (params as { sliceSize?: number }).sliceSize),
  closeWorkbook: (params: unknown) => closeWorkbook((params as { closeBehavior?: "Save" | "SkipSave" }).closeBehavior),
  snapshotRanges: (params: unknown) => snapshotRanges((params as { workbookId: string }).workbookId, (params as { ranges: Parameters<typeof snapshotRanges>[1] }).ranges),
  embedWorkbookLocalConfig: (params: unknown) => embedWorkbookLocalConfig(params as Parameters<typeof embedWorkbookLocalConfig>[0]),
  readWorkbookEmbeddedLocalConfig: (params: unknown) => readWorkbookEmbeddedLocalConfig((params as { workbookId: string }).workbookId)
};
