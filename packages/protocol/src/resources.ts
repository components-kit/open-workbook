import type { CatalogStatus } from "./tools.js";

export interface ResourceContract {
  uriTemplate: string;
  status: CatalogStatus;
  description: string;
}

export const ResourceCatalog: ResourceContract[] = [
  resource("excel://runtime/status", "stable"),
  resource("excel://workbooks", "stable"),
  resource("excel://workbooks/{workbook_id}/map", "stable"),
  resource("excel://workbooks/{workbook_id}/sheets", "stable"),
  resource("excel://workbooks/{workbook_id}/sheets/{sheet_name}/used-range", "stable"),
  resource("excel://workbooks/{workbook_id}/sheets/{sheet_name}/style-fingerprint", "stable"),
  resource("excel://workbooks/{workbook_id}/sheets/{sheet_name}/formula-patterns", "stable"),
  resource("excel://workbooks/{workbook_id}/tables", "stable"),
  resource("excel://workbooks/{workbook_id}/templates", "stable"),
  resource("excel://workbooks/{workbook_id}/snapshots/{snapshot_id}", "stable"),
  resource("excel://workbooks/{workbook_id}/plans/{plan_id}/diff", "stable"),
  resource("excel://compact/{resource_id}", "stable"),
  resource("excel://agent/contexts/{workbook_context_id}", "stable"),
  resource("excel://agent/operations/{operation_id}", "stable")
];

export const ResourceTemplates = ResourceCatalog.map((resourceContract) => resourceContract.uriTemplate);

export type ResourceTemplate = (typeof ResourceTemplates)[number];

function resource(uriTemplate: string, status: CatalogStatus): ResourceContract {
  return {
    uriTemplate,
    status,
    description: `Open Workbook resource ${uriTemplate}.`
  };
}
