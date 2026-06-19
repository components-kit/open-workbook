import { executeBatch } from "./executor-core.js";

export const batchHostOperations = {
  executeBatch: (params: unknown) => executeBatch(params as Parameters<typeof executeBatch>[0])
};
