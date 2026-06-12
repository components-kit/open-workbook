import { getExposedToolCatalog, ResourceCatalog } from "@open-workbook/protocol";

export function getToolCatalog(options: { includePreview?: boolean } = {}) {
  return getExposedToolCatalog(options);
}

export function getResourceCatalog() {
  return ResourceCatalog;
}
