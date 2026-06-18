import { getExposedToolCatalog, ResourceCatalog } from "@components-kit/open-workbook-protocol";

export function getToolCatalog() {
  return getExposedToolCatalog();
}

export function getResourceCatalog() {
  return ResourceCatalog;
}
