import {
  captureStyleFingerprint,
  copyStyleDimensions,
  copyStyleDimensionsMany
} from "./executor-core.js";

export const styleHostOperations = {
  captureStyleFingerprint: (params: unknown) => captureStyleFingerprint(params as Parameters<typeof captureStyleFingerprint>[0]),
  copyStyleDimensions: (params: unknown) => copyStyleDimensions(params as Parameters<typeof copyStyleDimensions>[0]),
  copyStyleDimensionsMany: (params: unknown) => copyStyleDimensionsMany(params as Parameters<typeof copyStyleDimensionsMany>[0])
};
