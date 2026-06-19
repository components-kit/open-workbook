import { McpServer, ResourceTemplate } from "@modelcontextprotocol/sdk/server/mcp.js";
import type { PlanId, SnapshotId, WorkbookId } from "@components-kit/open-workbook-protocol";
import type { RuntimeFacade } from "./runtime-facade.js";

export function registerResources(mcp: McpServer, runtime: RuntimeFacade): void {
  registerJsonResource(mcp, "runtime status", "excel://runtime/status", "Runtime connection, collaboration, and capability status.", async () => ({
    status: runtime.getStatus(),
    capabilities: runtime.getCapabilities(),
    collaboration: runtime.getCollaborationStatus()
  }));

  registerJsonResource(mcp, "workbooks", "excel://workbooks", "Open workbook references visible to connected add-ins.", async () => {
    const sessions = runtime.getStatus().sessions;
    return {
      workbooks: sessions.flatMap((session) => (session.activeWorkbook ? [session.activeWorkbook] : [])),
      sessions
    };
  });

  registerJsonTemplateResource(mcp, "workbook map", "excel://workbooks/{workbook_id}/map", "Workbook map with sheets, used ranges, and table names.", async () => runtime.getWorkbookMap());

  registerJsonTemplateResource(mcp, "workbook sheets", "excel://workbooks/{workbook_id}/sheets", "Worksheet list from the workbook map.", async (_uri, variables) => {
    const workbookId = resourceVariable(variables, "workbook_id") as WorkbookId;
    const map = await runtime.getWorkbookMap();
    return {
      ok: (map as { ok?: boolean }).ok,
      workbookId,
      sheets: (map as { map?: { sheets?: unknown[] } }).map?.sheets ?? [],
      source: map
    };
  });

  registerJsonTemplateResource(mcp, "sheet used range", "excel://workbooks/{workbook_id}/sheets/{sheet_name}/used-range", "Used range metadata for one worksheet.", async (_uri, variables) => {
    const workbookId = resourceVariable(variables, "workbook_id") as WorkbookId;
    const sheetName = resourceVariable(variables, "sheet_name");
    const map = await runtime.getWorkbookMap();
    const sheet = (map as { map?: { sheets?: Array<{ name: string; usedRange?: unknown }> } }).map?.sheets?.find((item) => item.name === sheetName);
    return { ok: Boolean(sheet), workbookId, sheetName, usedRange: sheet?.usedRange, source: map };
  });

  registerJsonTemplateResource(mcp, "sheet style fingerprint", "excel://workbooks/{workbook_id}/sheets/{sheet_name}/style-fingerprint", "Style fingerprint for one worksheet used range.", async (_uri, variables) => {
    const workbookId = resourceVariable(variables, "workbook_id") as WorkbookId;
    const sheetName = resourceVariable(variables, "sheet_name");
    return runtime.getStyleFingerprint({ workbookId, sheetName });
  });

  registerJsonTemplateResource(mcp, "sheet formula patterns", "excel://workbooks/{workbook_id}/sheets/{sheet_name}/formula-patterns", "Formula pattern summary for one worksheet used range.", async (_uri, variables) => {
    const workbookId = resourceVariable(variables, "workbook_id") as WorkbookId;
    const sheetName = resourceVariable(variables, "sheet_name");
    const map = await runtime.getWorkbookMap();
    const sheet = (map as { map?: { sheets?: Array<{ name: string; usedRange?: { address?: string } }> } }).map?.sheets?.find((item) => item.name === sheetName);
    const address = sheet?.usedRange?.address;
    if (!address) {
      return { ok: false, workbookId, sheetName, error: { code: "RANGE_INVALID", message: "Sheet used range is unavailable." }, source: map };
    }
    return runtime.readFormulaPatterns({ workbookId, sheetName, address: stripResourceSheetName(address) });
  });

  registerJsonTemplateResource(mcp, "workbook tables", "excel://workbooks/{workbook_id}/tables", "Structured table list for a workbook.", async (_uri, variables) => runtime.listTables(resourceVariable(variables, "workbook_id") as WorkbookId));

  registerJsonTemplateResource(mcp, "workbook templates", "excel://workbooks/{workbook_id}/templates", "Registered Open Workbook templates for a workbook.", async (_uri, variables) => {
    const workbookId = resourceVariable(variables, "workbook_id") as WorkbookId;
    return { ok: true, workbookId, templates: runtime.listTemplates(workbookId) };
  });

  registerJsonTemplateResource(mcp, "workbook snapshot", "excel://workbooks/{workbook_id}/snapshots/{snapshot_id}", "Stored snapshot metadata and payload reference.", async (_uri, variables) => ({
    workbookId: resourceVariable(variables, "workbook_id"),
    ...runtime.getSnapshot(resourceVariable(variables, "snapshot_id") as SnapshotId)
  }));

  registerJsonTemplateResource(mcp, "plan diff", "excel://workbooks/{workbook_id}/plans/{plan_id}/diff", "Stored plan preview diff summary.", async (_uri, variables) => {
    const workbookId = resourceVariable(variables, "workbook_id") as WorkbookId;
    const planId = resourceVariable(variables, "plan_id") as PlanId;
    return runtime.getPlanDiffResource(workbookId, planId);
  });

  registerJsonTemplateResource(mcp, "agent workbook context", "excel://agent/contexts/{workbook_context_id}", "Cached workbook metadata used by the Open Workbook agent workflow.", async (_uri, variables) => runtime.getAgentContextResource(resourceVariable(variables, "workbook_context_id")));

  registerJsonTemplateResource(mcp, "agent pending operation", "excel://agent/operations/{operation_id}", "Pending previewed workbook operation awaiting apply confirmation.", async (_uri, variables) => runtime.getAgentOperationResource(resourceVariable(variables, "operation_id")));
}

function registerJsonResource(
  mcp: McpServer,
  name: string,
  uri: string,
  description: string,
  read: (uri: URL) => unknown | Promise<unknown>
): void {
  mcp.registerResource(name, uri, { title: name, description, mimeType: "application/json" }, async (resourceUri) => jsonResource(resourceUri.toString(), await read(resourceUri)));
}

function registerJsonTemplateResource(
  mcp: McpServer,
  name: string,
  uriTemplate: string,
  description: string,
  read: (uri: URL, variables: Record<string, string | string[]>) => unknown | Promise<unknown>
): void {
  mcp.registerResource(
    name,
    new ResourceTemplate(uriTemplate, { list: undefined }),
    { title: name, description, mimeType: "application/json" },
    async (resourceUri, variables) => jsonResource(resourceUri.toString(), await read(resourceUri, variables as Record<string, string | string[]>))
  );
}

function jsonResource(uri: string, value: unknown) {
  return {
    contents: [
      {
        uri,
        mimeType: "application/json",
        text: JSON.stringify(value, null, 2)
      }
    ]
  };
}

function resourceVariable(variables: Record<string, string | string[]>, name: string): string {
  const value = variables[name];
  return Array.isArray(value) ? value[0] ?? "" : value ?? "";
}

function stripResourceSheetName(address: string): string {
  const bang = address.lastIndexOf("!");
  return bang >= 0 ? address.slice(bang + 1) : address;
}
