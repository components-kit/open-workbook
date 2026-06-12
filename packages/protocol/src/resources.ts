import type { CatalogStatus } from "./tools.js";

export interface ResourceContract {
  uriTemplate: string;
  status: CatalogStatus;
  description: string;
}

export const ResourceCatalog: ResourceContract[] = [
  resource("excel://runtime/status", "stable"),
  resource("excel://workbooks", "stable"),
  resource("excel://workbooks/{workbook_id}/map", "planned"),
  resource("excel://workbooks/{workbook_id}/sheets", "planned"),
  resource("excel://workbooks/{workbook_id}/sheets/{sheet_name}/used-range", "planned"),
  resource("excel://workbooks/{workbook_id}/sheets/{sheet_name}/style-fingerprint", "planned"),
  resource("excel://workbooks/{workbook_id}/sheets/{sheet_name}/formula-patterns", "planned"),
  resource("excel://workbooks/{workbook_id}/tables", "planned"),
  resource("excel://workbooks/{workbook_id}/templates", "planned"),
  resource("excel://workbooks/{workbook_id}/snapshots/{snapshot_id}", "planned"),
  resource("excel://workbooks/{workbook_id}/plans/{plan_id}/diff", "planned")
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
