import {
  captureSheetFingerprint,
  captureTemplate,
  repairTemplateConsistency
} from "./executor-core.js";

export const templateHostOperations = {
  captureTemplate: (params: unknown) => captureTemplate(params as Parameters<typeof captureTemplate>[0]),
  captureSheetFingerprint: (params: unknown) => captureSheetFingerprint(params as Parameters<typeof captureSheetFingerprint>[0]),
  repairTemplateConsistency: (params: unknown) => repairTemplateConsistency(params as Parameters<typeof repairTemplateConsistency>[0])
};
