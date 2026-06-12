export const ResourceTemplates = [
  "excel://runtime/status",
  "excel://workbooks",
  "excel://workbooks/{workbook_id}/map",
  "excel://workbooks/{workbook_id}/sheets",
  "excel://workbooks/{workbook_id}/templates",
  "excel://workbooks/{workbook_id}/snapshots/{snapshot_id}",
  "excel://workbooks/{workbook_id}/plans/{plan_id}/diff"
] as const;

export type ResourceTemplate = (typeof ResourceTemplates)[number];
