import { getExposedToolCatalog, ResourceCatalog } from "@component-kit/open-workbook-protocol";

export function getToolCatalog(options: { includePreview?: boolean } = {}) {
  return getExposedToolCatalog(options);
}

export function getResourceCatalog() {
  return ResourceCatalog;
}
