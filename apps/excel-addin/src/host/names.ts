import {
  createName,
  deleteName,
  getName,
  listNames,
  updateName
} from "./executor-core.js";

export const namesHostOperations = {
  listNames: (params: unknown) => listNames((params as { workbookId: string }).workbookId),
  getName: (params: unknown) => getName(params as Parameters<typeof getName>[0]),
  createName: (params: unknown) => createName(params as Parameters<typeof createName>[0]),
  updateName: (params: unknown) => updateName(params as Parameters<typeof updateName>[0]),
  deleteName: (params: unknown) => deleteName(params as Parameters<typeof deleteName>[0])
};
