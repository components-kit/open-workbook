import {
  getActiveWorkbookContext,
  getRuntimeCapabilities,
  getSelection,
  setActiveSheet
} from "./executor-core.js";

export const runtimeHostOperations = {
  ping: (params: unknown) => ({ ok: true, at: new Date().toISOString(), echo: params }),
  getActiveWorkbookContext,
  getRuntimeCapabilities,
  getSelection,
  setActiveSheet: (params: unknown) => setActiveSheet((params as { sheetName: string }).sheetName)
};
