#!/usr/bin/env node
import { McpServer, ResourceTemplate } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import * as z from "zod/v4";
import { RuntimeService } from "@open-workbook/backend/runtime";
import { startBackendServer } from "@open-workbook/backend/server";
import type {
  A1Range,
  AgentId,
  BackupId,
  BatchRequest,
  ChartCreateRequest,
  ChartSelector,
  ChartUpdateDataSourceRequest,
  ConflictRecord,
  ExcelOperation,
  FormulaCopyPatternsRequest,
  FormulaFillRequest,
  FormulaPatternRequest,
  LockId,
  LockMode,
  NameCreateRequest,
  NameSelector,
  NameUpdateRequest,
  OperationId,
  PermissionState,
  PivotCompareFingerprintRequest,
  PivotCopyFromTemplateRequest,
  PivotCreateRequest,
  PivotRebuildWithSourceRequest,
  PivotRepairFromTemplateRequest,
  PivotSelector,
  PivotValidateSourceRequest,
  PlanId,
  RegionRegisterRequest,
  RegionSelector,
  SnapshotId,
  TaskId,
  StyleDimension,
  TableAppendRowsRequest,
  TableApplyFiltersRequest,
  TableCopyStructureRequest,
  TableCreateRequest,
  TableResizeRequest,
  TableSelector,
  TableSetStyleRequest,
  TableSetTotalRowRequest,
  TableSortRequest,
  TableUpdateRowsRequest,
  TemplateId,
  TransactionId,
  WorkbookScope,
  WorkbookId,
  WorkbookBackupRetentionRequest,
  WorkbookCreateFileBackupRequest,
  WorkbookRestoreFileBackupRequest,
  WorkbookLocalConfig
} from "@open-workbook/protocol";
import { isToolExposed, makeId } from "@open-workbook/protocol";

type RuntimeFacade = RuntimeService & {
  compileBatch(request: BatchRequest): unknown;
};

const host = process.env.OPEN_WORKBOOK_HOST ?? "127.0.0.1";
const port = Number(process.env.OPEN_WORKBOOK_PORT ?? 37845);
const addinPath = process.env.OPEN_WORKBOOK_ADDIN_PATH ?? "/addin";
const daemonUrl = trimTrailingSlash(readArg("--daemon-url") ?? process.env.OPEN_WORKBOOK_DAEMON_URL ?? `http://${host}:${port}`);
const agentName = readArg("--agent-name") ?? process.env.OPEN_WORKBOOK_AGENT_NAME;
const standalone = hasArg("--standalone") || process.env.OPEN_WORKBOOK_MCP_STANDALONE === "1";
const catalogOptions = {
  includePreview: process.env.OPEN_WORKBOOK_PREVIEW_TOOLS === "1"
};

const STYLE_DIMENSIONS = [
  "columnWidths",
  "rowHeights",
  "borders",
  "fills",
  "fonts",
  "alignment",
  "numberFormats",
  "conditionalFormatting",
  "dataValidation",
  "freezePanes",
  "printSettings",
  "pageLayout",
  "hiddenRowsColumns"
] as const;

const STYLE_COPY_TOOL_DIMENSIONS: Record<string, StyleDimension> = {
  "excel.style.copy_column_widths": "columnWidths",
  "excel.style.copy_row_heights": "rowHeights",
  "excel.style.copy_borders": "borders",
  "excel.style.copy_fills": "fills",
  "excel.style.copy_fonts": "fonts",
  "excel.style.copy_alignment": "alignment",
  "excel.style.copy_number_formats": "numberFormats",
  "excel.style.copy_conditional_formatting": "conditionalFormatting",
  "excel.style.copy_data_validation": "dataValidation",
  "excel.style.copy_freeze_panes": "freezePanes",
  "excel.style.copy_print_settings": "printSettings",
  "excel.style.copy_page_layout": "pageLayout",
  "excel.style.copy_hidden_rows_columns": "hiddenRowsColumns"
};

const runtime = await createRuntimeFacade();

const server = new McpServer({
  name: "open-workbook",
  version: "0.1.0"
});

registerRuntimeTools(server);
registerWorkbookTools(server);
registerBackupTools(server);
registerSheetTools(server);
registerRangeTools(server);
registerBatchTools(server);
registerPlanTools(server);
registerTemplateTools(server);
registerStyleTools(server);
registerFormulaTools(server);
registerTableTools(server);
registerFilterTools(server);
registerSortTools(server);
registerPivotTools(server);
registerChartTools(server);
registerNamesTools(server);
registerRegionTools(server);
registerTaskTools(server);
registerCollaborationTools(server);
registerLockTools(server);
registerConflictTools(server);
registerTransactionTools(server);
registerPermissionsTools(server);
registerCleanTools(server);
registerValidateTools(server);
registerRepairTools(server);
registerSnapshotTools(server);
registerDiffTools(server);
registerEventTools(server);
registerResources(server);
registerPrompts(server);

await server.connect(new StdioServerTransport());

async function createRuntimeFacade(): Promise<RuntimeFacade> {
  if (!standalone && await daemonAvailable(daemonUrl)) {
    const proxy = createDaemonRuntimeProxy(daemonUrl) as RuntimeFacade;
    const registration = await proxy.registerAgent({ agentName, clientType: "mcp", pid: process.pid });
    const registeredAgent = (registration as { agent?: { agentId?: string } }).agent;
    console.error(`open-workbook MCP adapter connected to ${daemonUrl}${registeredAgent?.agentId ? ` as ${registeredAgent.agentId}` : ""}`);
    return proxy;
  }

  const localRuntime = new RuntimeService() as RuntimeFacade;
  if (agentName !== undefined) {
    localRuntime.registerAgent({ agentName, clientType: "mcp", pid: process.pid });
  }
  await startBackendServer(localRuntime, { host, port, addinPath });
  console.error(`open-workbook MCP standalone backend listening on ws://${host}:${port}${addinPath}`);
  return localRuntime;
}

async function daemonAvailable(baseUrl: string): Promise<boolean> {
  try {
    const response = await fetch(`${baseUrl}/status`);
    return response.ok;
  } catch {
    return false;
  }
}

function createDaemonRuntimeProxy(baseUrl: string): unknown {
  const call = async (method: string, args: unknown[]) => {
    const response = await fetch(`${baseUrl}/rpc`, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({ method, args })
    });
    const payload = await response.json() as { ok: boolean; result?: unknown; error?: unknown };
    if (!response.ok || !payload.ok) {
      throw new Error(JSON.stringify(payload.error ?? { code: "OPERATION_FAILED", message: `Daemon RPC failed: ${method}` }));
    }
    return payload.result;
  };
  return new Proxy(
    {},
    {
      get(_target, property) {
        if (typeof property !== "string") {
          return undefined;
        }
        return (...args: unknown[]) => call(property, args);
      }
    }
  );
}

function registerResources(mcp: McpServer): void {
  registerJsonResource(mcp, "runtime status", "excel://runtime/status", "Runtime connection, collaboration, and capability status.", async (uri) => ({
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

  registerJsonTemplateResource(
    mcp,
    "workbook map",
    "excel://workbooks/{workbook_id}/map",
    "Workbook map with sheets, used ranges, and table names.",
    async (_uri, variables) => runtime.getWorkbookMap()
  );

  registerJsonTemplateResource(
    mcp,
    "workbook sheets",
    "excel://workbooks/{workbook_id}/sheets",
    "Worksheet list from the workbook map.",
    async (_uri, variables) => {
      const workbookId = resourceVariable(variables, "workbook_id") as WorkbookId;
      const map = await runtime.getWorkbookMap();
      return {
        ok: (map as { ok?: boolean }).ok,
        workbookId,
        sheets: (map as { map?: { sheets?: unknown[] } }).map?.sheets ?? [],
        source: map
      };
    }
  );

  registerJsonTemplateResource(
    mcp,
    "sheet used range",
    "excel://workbooks/{workbook_id}/sheets/{sheet_name}/used-range",
    "Used range metadata for one worksheet.",
    async (_uri, variables) => {
      const workbookId = resourceVariable(variables, "workbook_id") as WorkbookId;
      const sheetName = resourceVariable(variables, "sheet_name");
      const map = await runtime.getWorkbookMap();
      const sheet = (map as { map?: { sheets?: Array<{ name: string; usedRange?: unknown }> } }).map?.sheets?.find((item) => item.name === sheetName);
      return {
        ok: Boolean(sheet),
        workbookId,
        sheetName,
        usedRange: sheet?.usedRange,
        source: map
      };
    }
  );

  registerJsonTemplateResource(
    mcp,
    "sheet style fingerprint",
    "excel://workbooks/{workbook_id}/sheets/{sheet_name}/style-fingerprint",
    "Style fingerprint for one worksheet used range.",
    async (_uri, variables) => {
      const workbookId = resourceVariable(variables, "workbook_id") as WorkbookId;
      const sheetName = resourceVariable(variables, "sheet_name");
      return runtime.getStyleFingerprint({ workbookId, sheetName });
    }
  );

  registerJsonTemplateResource(
    mcp,
    "sheet formula patterns",
    "excel://workbooks/{workbook_id}/sheets/{sheet_name}/formula-patterns",
    "Formula pattern summary for one worksheet used range.",
    async (_uri, variables) => {
      const workbookId = resourceVariable(variables, "workbook_id") as WorkbookId;
      const sheetName = resourceVariable(variables, "sheet_name");
      const map = await runtime.getWorkbookMap();
      const sheet = (map as { map?: { sheets?: Array<{ name: string; usedRange?: { address?: string } }> } }).map?.sheets?.find((item) => item.name === sheetName);
      const address = sheet?.usedRange?.address;
      if (!address) {
        return {
          ok: false,
          workbookId,
          sheetName,
          error: { code: "RANGE_INVALID", message: "Sheet used range is unavailable." },
          source: map
        };
      }
      return runtime.readFormulaPatterns({ workbookId, sheetName, address: stripResourceSheetName(address) });
    }
  );

  registerJsonTemplateResource(
    mcp,
    "workbook tables",
    "excel://workbooks/{workbook_id}/tables",
    "Structured table list for a workbook.",
    async (_uri, variables) => runtime.listTables(resourceVariable(variables, "workbook_id") as WorkbookId)
  );

  registerJsonTemplateResource(
    mcp,
    "workbook templates",
    "excel://workbooks/{workbook_id}/templates",
    "Registered Open Workbook templates for a workbook.",
    async (_uri, variables) => ({
      ok: true,
      workbookId: resourceVariable(variables, "workbook_id"),
      templates: runtime.listTemplates(resourceVariable(variables, "workbook_id") as WorkbookId)
    })
  );

  registerJsonTemplateResource(
    mcp,
    "workbook snapshot",
    "excel://workbooks/{workbook_id}/snapshots/{snapshot_id}",
    "Stored snapshot metadata and payload reference.",
    async (_uri, variables) => {
      const snapshot = runtime.getSnapshot(resourceVariable(variables, "snapshot_id") as SnapshotId);
      return {
        workbookId: resourceVariable(variables, "workbook_id"),
        ...snapshot
      };
    }
  );

  registerJsonTemplateResource(
    mcp,
    "plan diff",
    "excel://workbooks/{workbook_id}/plans/{plan_id}/diff",
    "Stored plan preview diff summary.",
    async (_uri, variables) => {
      const workbookId = resourceVariable(variables, "workbook_id") as WorkbookId;
      const planId = resourceVariable(variables, "plan_id") as PlanId;
      return runtime.getPlanDiffResource(workbookId, planId);
    }
  );
}

function registerJsonResource(
  mcp: McpServer,
  name: string,
  uri: string,
  description: string,
  read: (uri: URL) => unknown | Promise<unknown>
): void {
  mcp.registerResource(
    name,
    uri,
    {
      title: name,
      description,
      mimeType: "application/json"
    },
    async (resourceUri) => jsonResource(resourceUri.toString(), await read(resourceUri))
  );
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
    {
      title: name,
      description,
      mimeType: "application/json"
    },
    async (resourceUri, variables) => jsonResource(resourceUri.toString(), await read(resourceUri, variables as Record<string, string | string[]>))
  );
}

function registerPrompts(mcp: McpServer): void {
  const promptArgs = {
    workbookId: z.string().optional(),
    sheetName: z.string().optional(),
    templateId: z.string().optional(),
    targetSheetName: z.string().optional(),
    goal: z.string().optional()
  };

  registerWorkflowPrompt(
    mcp,
    "excel.prompts.create_next_month_sheet",
    "Create next month sheet",
    "Plan and safely create a next-period worksheet from an existing template or previous-period sheet.",
    promptArgs,
    (args) => [
      "Create a next-period worksheet without damaging formulas, formatting, filters, tables, print layout, or named regions.",
      promptContext(args),
      "Workflow:",
      "1. Read `excel.runtime.get_active_context`, then inspect `excel.workbook.get_workbook_map` and `excel.template.list`.",
      "2. Prefer a registered template. If no template is registered, call `excel.template.detect_templates` and ask the user to confirm the source sheet.",
      "3. Use `excel.plan.create` and `excel.plan.preview` before mutation.",
      "4. Use `excel.template.create_sheet_from_template` or `excel.template.create_sheet_from_previous_period`.",
      "5. Clear only declared data regions with `excel.template.clear_data_regions` or `excel.range.clear_values_keep_format`.",
      "6. Validate with `excel.template.validate_sheet_against_template`, `excel.formula.validate_against_template`, `excel.style.validate_consistency`, and `excel.validate.no_formula_errors`.",
      "7. Commit only after validation is clean or after discussing warnings with the user."
    ]
  );

  registerWorkflowPrompt(
    mcp,
    "excel.prompts.clean_current_sheet",
    "Clean current sheet",
    "Clean worksheet data while preserving workbook structure, styling, formulas, filters, and templates.",
    promptArgs,
    (args) => [
      "Clean the current worksheet conservatively. Do not overwrite formulas, templates, filters, styling, or hidden layout areas.",
      promptContext(args),
      "Workflow:",
      "1. Read active context, selection, used range, tables, filters, formulas, and style fingerprint.",
      "2. Identify data-entry regions using registered regions, table data bodies, or template data regions.",
      "3. Preview transformations with read-only tools first: header detection, trim/normalize, parse dates/numbers, duplicate/outlier checks.",
      "4. Create a plan and preview it. Apply only scoped range/table operations.",
      "5. Prefer `excel.table.update_rows`, `excel.region.write_values`, or `excel.range.write_values` with format preservation.",
      "6. Re-run table/filter/style/formula validation and summarize exactly what changed."
    ]
  );

  registerWorkflowPrompt(
    mcp,
    "excel.prompts.fix_formula_errors",
    "Fix formula errors",
    "Diagnose formula errors, compare against template patterns, and repair only after preview and validation.",
    promptArgs,
    (args) => [
      "Fix formula errors without converting formulas to values unless the user explicitly asks.",
      promptContext(args),
      "Workflow:",
      "1. Locate errors with `excel.formula.find_errors` and `excel.validate.no_formula_errors`.",
      "2. Read formula patterns and dependency graph with `excel.formula.read_patterns`, `excel.formula.get_dependency_graph`, `trace_precedents`, and `trace_dependents`.",
      "3. If a template exists, compare with `excel.formula.validate_against_template`.",
      "4. Create a repair plan using `excel.formula.repair_patterns`, `fill_down`, `fill_right`, or explicit `range.write_formulas`.",
      "5. Preview, apply, recalculate, and re-run formula validation before reporting success."
    ]
  );

  registerWorkflowPrompt(
    mcp,
    "excel.prompts.format_like_template",
    "Format like template",
    "Repair styling and layout consistency using registered template fingerprints.",
    promptArgs,
    (args) => [
      "Make the target sheet look like the template while preserving current data values.",
      promptContext(args),
      "Workflow:",
      "1. Read template registry and current style fingerprints.",
      "2. Compare with `excel.style.compare_fingerprint` and `excel.style.validate_consistency`.",
      "3. Ask before changing structure-level layout such as hidden rows/columns, freeze panes, print settings, or page layout.",
      "4. Use granular style copy tools or `excel.style.repair_consistency` for confirmed dimensions.",
      "5. Validate styles, formulas, tables, filters, and print layout after applying."
    ]
  );

  registerWorkflowPrompt(
    mcp,
    "excel.prompts.validate_report_before_saving",
    "Validate report before saving",
    "Run workbook/report validation before saving or handing a file back to the user.",
    promptArgs,
    (args) => [
      "Validate the report before saving. Do not save if validation finds material formula, reference, template, or unintended-change issues.",
      promptContext(args),
      "Workflow:",
      "1. Create or refresh a snapshot if one is available for unintended-change checks.",
      "2. Run workbook, sheet, template, formula, style, table, filter, print-layout, broken-reference, formula-error, and unintended-change validators.",
      "3. Summarize issues by severity and affected range/table/sheet.",
      "4. Repair only with explicit scoped tools and backups.",
      "5. Save with `excel.workbook.save` only when errors are clean or the user confirms known warnings."
    ]
  );

  registerWorkflowPrompt(
    mcp,
    "excel.prompts.create_summary_report",
    "Create summary report",
    "Create a summary/report sheet from existing workbook data with safe planning and validation.",
    promptArgs,
    (args) => [
      "Create a summary report from existing workbook data without disturbing source sheets.",
      promptContext(args),
      "Workflow:",
      "1. Map workbook sheets, tables, names, regions, filters, PivotTables, and charts.",
      "2. Ask the user which metrics/groupings/date ranges matter if not obvious.",
      "3. Prefer creating a new sheet from a template or copying a previous report sheet.",
      "4. Use table reads, formulas, PivotTables, and charts through planned operations.",
      "5. Preview and apply via plan/batch; never write directly outside the target report regions.",
      "6. Validate formulas, style consistency, tables, charts, and no unintended source changes."
    ]
  );

  for (const name of [
    "excel.prompts.import_receipts_to_table",
    "excel.prompts.import_invoices_to_table",
    "excel.prompts.reconcile_statement",
    "excel.prompts.create_driver_payroll",
    "excel.prompts.import_fuel_slips",
    "excel.prompts.calculate_fuel_consumption",
    "excel.prompts.create_customer_transport_report",
    "excel.prompts.reconcile_job_payments"
  ]) {
    registerUnsupportedPrompt(mcp, name);
  }
}

function registerWorkflowPrompt(
  mcp: McpServer,
  name: string,
  title: string,
  description: string,
  argsSchema: Record<string, z.ZodTypeAny>,
  body: (args: Record<string, unknown>) => string[]
): void {
  mcp.registerPrompt(
    name,
    {
      title,
      description,
      argsSchema
    },
    (args) => ({
      description,
      messages: [
        {
          role: "user",
          content: {
            type: "text",
            text: body(args as Record<string, unknown>).filter(Boolean).join("\n")
          }
        }
      ]
    })
  );
}

function registerUnsupportedPrompt(mcp: McpServer, name: string): void {
  mcp.registerPrompt(
    name,
    {
      title: name.replace(/^excel\.prompts\./, "").replace(/_/g, " "),
      description: "This vertical workflow is intentionally unsupported in the current Open Workbook runtime.",
      argsSchema: {
        goal: z.string().optional()
      }
    },
    () => ({
      description: "Unsupported vertical workflow.",
      messages: [
        {
          role: "user",
          content: {
            type: "text",
            text: [
              "This prompt is intentionally unsupported in the current Open Workbook runtime.",
              "OCR and vertical reconciliation workflows are out of scope for now.",
              "Use generic table, range, template, validation, and cleaning tools with explicit user-provided data instead."
            ].join("\n")
          }
        }
      ]
    })
  );
}

function promptContext(args: Record<string, unknown>): string {
  const entries = Object.entries(args).filter(([, value]) => value !== undefined && value !== "");
  if (entries.length === 0) {
    return "";
  }
  return `Context: ${entries.map(([key, value]) => `${key}=${String(value)}`).join(", ")}`;
}

function registerRuntimeTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.runtime.get_status",
    {
      title: "Get Excel runtime status",
      description: "Return backend, Excel add-in, and optional native file bridge health status.",
      inputSchema: {
        probeFileBridge: z.boolean().optional()
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ probeFileBridge }: { probeFileBridge?: boolean }) => jsonResult(probeFileBridge ? await runtime.getStatusWithFileBridgeProbe() : runtime.getStatus())
  );

  registerMcpTool(
    mcp,
    "excel.runtime.connect_addin",
    {
      title: "Get add-in connection instructions",
      description: "Return the local backend WebSocket URL that the Excel add-in should connect to.",
      inputSchema: {},
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async () => jsonResult(runtime.connectAddinInfo())
  );

  registerMcpTool(
    mcp,
    "excel.runtime.disconnect_addin",
    {
      title: "Disconnect active add-in",
      description: "Close the active Excel add-in session.",
      inputSchema: {},
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async () => jsonResult(runtime.disconnectActiveAddin())
  );

  registerMcpTool(
    mcp,
    "excel.runtime.ping_addin",
    {
      title: "Ping active add-in",
      description: "Ping the active Excel add-in session.",
      inputSchema: {},
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async () => jsonResult(await runtime.pingAddin())
  );

  registerMcpTool(
    mcp,
    "excel.runtime.get_capabilities",
    {
      title: "Get Open Workbook capabilities",
      description: "Return complete tool/resource/prompt catalog status and runtime capability metadata.",
      inputSchema: {
        includePreview: z.boolean().optional()
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ includePreview }: { includePreview?: boolean }) =>
      jsonResult(runtime.getCapabilities(includePreview === undefined ? {} : { includePreview }))
  );

  registerMcpTool(
    mcp,
    "excel.runtime.get_active_context",
    {
      title: "Get active Excel context",
      description: "Return active workbook context from the connected Excel add-in.",
      inputSchema: {},
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async () => jsonResult(await runtime.getActiveContext())
  );

  registerMcpTool(
    mcp,
    "excel.runtime.get_selection",
    {
      title: "Get active Excel selection",
      description: "Return the current selected range from the active Excel add-in.",
      inputSchema: {},
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async () => jsonResult(await runtime.getSelection())
  );

  registerMcpTool(
    mcp,
    "excel.runtime.set_active_workbook",
    {
      title: "Set active workbook session",
      description: "Select the active connected workbook session by workbook ID or workbook name.",
      inputSchema: {
        workbookIdOrName: z.string()
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ workbookIdOrName }: { workbookIdOrName: string }) => jsonResult(runtime.setActiveWorkbook(workbookIdOrName))
  );

  registerMcpTool(
    mcp,
    "excel.runtime.set_active_sheet",
    {
      title: "Set active worksheet",
      description: "Activate a worksheet in the active connected workbook.",
      inputSchema: {
        sheetName: z.string()
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ sheetName }: { sheetName: string }) => jsonResult(await runtime.setActiveSheet(sheetName))
  );
}

function registerWorkbookTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.workbook.list_open_workbooks",
    {
      title: "List open Excel workbooks",
      description: "List workbooks currently visible to connected add-ins.",
      inputSchema: {},
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async () => {
      const sessions = runtime.getStatus().sessions;
      return jsonResult({
        workbooks: sessions.flatMap((session) => (session.activeWorkbook ? [session.activeWorkbook] : []))
      });
    }
  );

  registerMcpTool(
    mcp,
    "excel.workbook.get_workbook_info",
    {
      title: "Get workbook info",
      description: "Return active workbook metadata from Excel.",
      inputSchema: {},
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async () => jsonResult(await runtime.getWorkbookInfo())
  );

  registerMcpTool(
    mcp,
    "excel.workbook.get_workbook_map",
    {
      title: "Get workbook map",
      description: "Return worksheets, used ranges, and table names for the active workbook.",
      inputSchema: {},
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async () => jsonResult(await runtime.getWorkbookMap())
  );

  registerMcpTool(
    mcp,
    "excel.workbook.snapshot",
    {
      title: "Create workbook snapshot",
      description: "Capture a restorable snapshot of used ranges or specific ranges.",
      inputSchema: snapshotInputSchema(),
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ workbookId, reason, ranges }: { workbookId: string; reason?: string; ranges?: A1Range[] }) =>
      jsonResult(await runtime.createWorkbookSnapshot(snapshotRequest(workbookId, reason, ranges)))
  );

  registerMcpTool(
    mcp,
    "excel.workbook.refresh_snapshot",
    {
      title: "Refresh workbook snapshot",
      description: "Capture a fresh snapshot over the same ranges as an existing snapshot.",
      inputSchema: {
        snapshotId: z.string(),
        reason: z.string().optional()
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ snapshotId, reason }: { snapshotId: string; reason?: string }) => {
      const existing = await runtime.getSnapshot(snapshotId as SnapshotId);
      if (!existing.ok || !("snapshot" in existing)) {
        return jsonResult(existing);
      }
      return jsonResult(
        await runtime.createWorkbookSnapshot({
          workbookId: existing.snapshot.workbookId,
          reason: reason ?? `Refresh snapshot ${snapshotId}`,
          ranges: existing.snapshot.affectedRanges
        })
      );
    }
  );

  registerMcpTool(
    mcp,
    "excel.workbook.get_snapshot",
    {
      title: "Get workbook snapshot",
      description: "Return a captured workbook snapshot by ID.",
      inputSchema: {
        snapshotId: z.string()
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ snapshotId }: { snapshotId: string }) => jsonResult(runtime.getSnapshot(snapshotId as SnapshotId))
  );

  registerMcpTool(
    mcp,
    "excel.workbook.detect_external_changes",
    {
      title: "Detect external workbook changes",
      description: "Compare a stored snapshot with the current workbook state over the same ranges.",
      inputSchema: {
        workbookId: z.string(),
        snapshotId: z.string()
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ workbookId, snapshotId }: { workbookId: string; snapshotId: string }) =>
      jsonResult(await runtime.detectExternalChanges({ workbookId: workbookId as WorkbookId, snapshotId: snapshotId as SnapshotId }))
  );

  registerMcpTool(
    mcp,
    "excel.workbook.calculate",
    {
      title: "Calculate workbook",
      description: "Recalculate the active workbook.",
      inputSchema: {
        workbookId: z.string(),
        calculationType: z.enum(["full", "recalculate"]).optional()
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ workbookId, calculationType }: { workbookId: string; calculationType?: "full" | "recalculate" }) =>
      jsonResult(await runtime.calculateWorkbook(workbookId as WorkbookId, calculationType))
  );

  registerMcpTool(
    mcp,
    "excel.workbook.save",
    {
      title: "Save workbook",
      description: "Save the active workbook.",
      inputSchema: {
        workbookId: z.string()
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ workbookId }: { workbookId: string }) => jsonResult(await runtime.saveWorkbook(workbookId as WorkbookId))
  );

  registerMcpTool(
    mcp,
    "excel.workbook.save_as",
    {
      title: "Save workbook as",
      description: "Save the workbook through the native file bridge when configured, otherwise report Save As capability status.",
      inputSchema: {
        workbookId: z.string(),
        targetPath: z.string().optional()
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        openWorldHint: false
      }
    },
    async ({ workbookId, targetPath }: { workbookId: string; targetPath?: string }) =>
      jsonResult(await runtime.saveWorkbookAs(workbookId as WorkbookId, targetPath))
  );

  registerMcpTool(
    mcp,
    "excel.workbook.create_backup",
    {
      title: "Create workbook backup",
      description: "Create a session-restorable workbook backup snapshot.",
      inputSchema: snapshotInputSchema(),
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ workbookId, reason, ranges }: { workbookId: string; reason?: string; ranges?: A1Range[] }) =>
      jsonResult(await runtime.createWorkbookBackup(snapshotRequest(workbookId, reason, ranges)))
  );

  registerMcpTool(
    mcp,
    "excel.workbook.restore_backup",
    {
      title: "Restore workbook backup",
      description: "Restore a backup captured by Open Workbook.",
      inputSchema: {
        backupId: z.string(),
        confirmationToken: z.string().optional()
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        openWorldHint: false
      }
    },
    async ({ backupId, confirmationToken }: { backupId: string; confirmationToken?: string }) =>
      jsonResult(await runtime.restoreBackup(backupId as BackupId, confirmationToken))
  );

  registerMcpTool(
    mcp,
    "excel.workbook.export_copy",
    {
      title: "Export workbook copy",
      description: "Create a persistent snapshot backup and report .xlsx export capability status.",
      inputSchema: {
        workbookId: z.string(),
        reason: z.string().optional(),
        targetPath: z.string().optional(),
        ranges: z
          .array(
            z.object({
              workbookId: z.string(),
              sheetName: z.string(),
              address: z.string()
            })
          )
          .optional()
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ workbookId, reason, targetPath, ranges }: { workbookId: string; reason?: string; targetPath?: string; ranges?: A1Range[] }) => {
      const request: { workbookId: WorkbookId; reason?: string; targetPath?: string; ranges?: A1Range[] } = {
        workbookId: workbookId as WorkbookId
      };
      if (reason !== undefined) {
        request.reason = reason;
      }
      if (targetPath !== undefined) {
        request.targetPath = targetPath;
      }
      if (ranges !== undefined) {
        request.ranges = ranges;
      }
      return jsonResult(await runtime.exportWorkbookCopy(request));
    }
  );

  registerMcpTool(
    mcp,
    "excel.workbook.export_local_config",
    {
      title: "Export workbook local config",
      description: "Export Open Workbook templates, registered regions, and workbook permission metadata as portable JSON.",
      inputSchema: {
        workbookId: z.string(),
        includePermissions: z.boolean().optional()
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ workbookId, includePermissions }: { workbookId: string; includePermissions?: boolean }) => {
      const options: { includePermissions?: boolean } = {};
      if (includePermissions !== undefined) {
        options.includePermissions = includePermissions;
      }
      return jsonResult(await runtime.exportWorkbookLocalConfig(workbookId as WorkbookId, options));
    }
  );

  registerMcpTool(
    mcp,
    "excel.workbook.import_local_config",
    {
      title: "Import workbook local config",
      description: "Import portable Open Workbook templates, registered regions, and workbook permission metadata into the local daemon registry.",
      inputSchema: {
        workbookId: z.string(),
        config: z.object({
          version: z.literal(1),
          workbookId: z.string(),
          exportedAt: z.string(),
          source: z.literal("open-workbook-local-config"),
          templates: z.array(z.record(z.string(), z.unknown())),
          regions: z.array(z.any()),
          permissions: z.any().optional()
        }),
        includeTemplates: z.boolean().optional(),
        includeRegions: z.boolean().optional(),
        includePermissions: z.boolean().optional(),
        overwrite: z.boolean().optional()
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({
      workbookId,
      config,
      includeTemplates,
      includeRegions,
      includePermissions,
      overwrite
    }: {
      workbookId: string;
      config: WorkbookLocalConfig;
      includeTemplates?: boolean;
      includeRegions?: boolean;
      includePermissions?: boolean;
      overwrite?: boolean;
    }) => {
      const request = {
        workbookId: workbookId as WorkbookId,
        config
      } as {
        workbookId: WorkbookId;
        config: WorkbookLocalConfig;
        includeTemplates?: boolean;
        includeRegions?: boolean;
        includePermissions?: boolean;
        overwrite?: boolean;
      };
      if (includeTemplates !== undefined) {
        request.includeTemplates = includeTemplates;
      }
      if (includeRegions !== undefined) {
        request.includeRegions = includeRegions;
      }
      if (includePermissions !== undefined) {
        request.includePermissions = includePermissions;
      }
      if (overwrite !== undefined) {
        request.overwrite = overwrite;
      }
      return jsonResult(await runtime.importWorkbookLocalConfig(request));
    }
  );

  registerMcpTool(
    mcp,
    "excel.workbook.embed_local_config",
    {
      title: "Embed workbook local config",
      description: "Write Open Workbook template, region, and permission metadata into the workbook custom XML part when the Excel host supports it.",
      inputSchema: {
        workbookId: z.string(),
        includePermissions: z.boolean().optional()
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ workbookId, includePermissions }: { workbookId: string; includePermissions?: boolean }) => {
      const options: { includePermissions?: boolean } = {};
      if (includePermissions !== undefined) {
        options.includePermissions = includePermissions;
      }
      return jsonResult(await runtime.embedWorkbookLocalConfig(workbookId as WorkbookId, options));
    }
  );

  registerMcpTool(
    mcp,
    "excel.workbook.read_embedded_local_config",
    {
      title: "Read embedded workbook local config",
      description: "Read Open Workbook local config metadata from the workbook custom XML part when present.",
      inputSchema: {
        workbookId: z.string()
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ workbookId }: { workbookId: string }) => jsonResult(await runtime.readWorkbookEmbeddedLocalConfig(workbookId as WorkbookId))
  );

  registerMcpTool(
    mcp,
    "excel.workbook.import_embedded_local_config",
    {
      title: "Import embedded workbook local config",
      description: "Read workbook custom XML metadata and import it into the local daemon registry.",
      inputSchema: {
        workbookId: z.string(),
        includeTemplates: z.boolean().optional(),
        includeRegions: z.boolean().optional(),
        includePermissions: z.boolean().optional(),
        overwrite: z.boolean().optional()
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({
      workbookId,
      includeTemplates,
      includeRegions,
      includePermissions,
      overwrite
    }: {
      workbookId: string;
      includeTemplates?: boolean;
      includeRegions?: boolean;
      includePermissions?: boolean;
      overwrite?: boolean;
    }) => {
      const request: {
        workbookId: WorkbookId;
        includeTemplates?: boolean;
        includeRegions?: boolean;
        includePermissions?: boolean;
        overwrite?: boolean;
      } = { workbookId: workbookId as WorkbookId };
      if (includeTemplates !== undefined) {
        request.includeTemplates = includeTemplates;
      }
      if (includeRegions !== undefined) {
        request.includeRegions = includeRegions;
      }
      if (includePermissions !== undefined) {
        request.includePermissions = includePermissions;
      }
      if (overwrite !== undefined) {
        request.overwrite = overwrite;
      }
      return jsonResult(await runtime.importWorkbookEmbeddedLocalConfig(request));
    }
  );

  registerMcpTool(
    mcp,
    "excel.workbook.close",
    {
      title: "Close workbook",
      description: "Close the active workbook through Office.js.",
      inputSchema: {
        workbookId: z.string(),
        closeBehavior: z.enum(["Save", "SkipSave"]).optional()
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        openWorldHint: false
      }
    },
    async ({ workbookId, closeBehavior }: { workbookId: string; closeBehavior?: "Save" | "SkipSave" }) =>
      jsonResult(await runtime.closeWorkbook(workbookId as WorkbookId, closeBehavior))
  );
}

function registerBackupTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.backup.create_file",
    {
      title: "Create file backup",
      description: "Create a verified full .xlsx file backup using native SaveCopyAs or the supported Office.js file export fallback.",
      inputSchema: {
        workbookId: z.string(),
        reason: z.string().optional(),
        targetPath: z.string().optional(),
        mode: z.enum(["export-copy", "save-copy-as"]).optional(),
        pin: z.boolean().optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => {
      const request: WorkbookCreateFileBackupRequest = { workbookId: args.workbookId as WorkbookId };
      if (args.reason !== undefined) request.reason = args.reason;
      if (args.targetPath !== undefined) request.targetPath = args.targetPath;
      if (args.mode !== undefined) request.mode = args.mode;
      if (args.pin !== undefined) request.pin = args.pin;
      return jsonResult(await runtime.createFileBackup(request));
    }
  );

  registerMcpTool(
    mcp,
    "excel.backup.list",
    {
      title: "List file backups",
      description: "List durable full-file workbook backups.",
      inputSchema: { workbookId: z.string().optional() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ workbookId }: { workbookId?: string }) => jsonResult(runtime.listFileBackups(workbookId as WorkbookId | undefined))
  );

  registerMcpTool(
    mcp,
    "excel.backup.get",
    {
      title: "Get file backup",
      description: "Return one durable full-file workbook backup manifest.",
      inputSchema: { backupId: z.string() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ backupId }: { backupId: string }) => jsonResult(runtime.getFileBackup(backupId as BackupId))
  );

  registerMcpTool(
    mcp,
    "excel.backup.verify",
    {
      title: "Verify file backup",
      description: "Verify that a full-file backup exists and still matches its checksum.",
      inputSchema: { backupId: z.string() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ backupId }: { backupId: string }) => jsonResult(await runtime.verifyFileBackup(backupId as BackupId))
  );

  registerMcpTool(
    mcp,
    "excel.backup.restore_file",
    {
      title: "Restore file backup",
      description: "Restore a full-file backup. open-as-new is safe; destructive modes require confirmation and native bridge support.",
      inputSchema: {
        workbookId: z.string(),
        backupId: z.string(),
        mode: z.enum(["open-as-new", "replace-open-workbook", "restore-into-open-workbook"]).optional(),
        restoreTargetPath: z.string().optional(),
        confirmationToken: z.string().optional(),
        force: z.boolean().optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => {
      const request: WorkbookRestoreFileBackupRequest = {
        workbookId: args.workbookId as WorkbookId,
        backupId: args.backupId as BackupId
      };
      if (args.mode !== undefined) request.mode = args.mode;
      if (args.restoreTargetPath !== undefined) request.restoreTargetPath = args.restoreTargetPath;
      if (args.confirmationToken !== undefined) request.confirmationToken = args.confirmationToken;
      if (args.force !== undefined) request.force = args.force;
      return jsonResult(await runtime.restoreFileBackup(request));
    }
  );

  registerMcpTool(
    mcp,
    "excel.backup.delete",
    {
      title: "Delete file backup",
      description: "Delete an unpinned durable full-file backup and its file payload when possible.",
      inputSchema: { backupId: z.string() },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async ({ backupId }: { backupId: string }) => jsonResult(runtime.deleteFileBackup(backupId as BackupId))
  );

  registerMcpTool(
    mcp,
    "excel.backup.prune",
    {
      title: "Prune file backups",
      description: "Prune unpinned durable file backups by age or per-workbook retention count.",
      inputSchema: {
        workbookId: z.string().optional(),
        maxAgeDays: z.number().optional(),
        maxBackupsPerWorkbook: z.number().optional(),
        dryRun: z.boolean().optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => {
      const request: WorkbookBackupRetentionRequest = {};
      if (args.workbookId !== undefined) request.workbookId = args.workbookId as WorkbookId;
      if (args.maxAgeDays !== undefined) request.maxAgeDays = args.maxAgeDays;
      if (args.maxBackupsPerWorkbook !== undefined) request.maxBackupsPerWorkbook = args.maxBackupsPerWorkbook;
      if (args.dryRun !== undefined) request.dryRun = args.dryRun;
      return jsonResult(runtime.pruneFileBackups(request));
    }
  );

  registerMcpTool(
    mcp,
    "excel.backup.pin",
    {
      title: "Pin file backup",
      description: "Prevent a durable file backup from being pruned or deleted.",
      inputSchema: { backupId: z.string() },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async ({ backupId }: { backupId: string }) => jsonResult(runtime.pinFileBackup(backupId as BackupId, true))
  );

  registerMcpTool(
    mcp,
    "excel.backup.unpin",
    {
      title: "Unpin file backup",
      description: "Allow a durable file backup to be pruned or deleted.",
      inputSchema: { backupId: z.string() },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async ({ backupId }: { backupId: string }) => jsonResult(runtime.pinFileBackup(backupId as BackupId, false))
  );
}

function registerSheetTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.sheet.list",
    {
      title: "List worksheets",
      description: "List worksheets in the active workbook.",
      inputSchema: {},
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async () => {
      const result = await runtime.getWorkbookMap();
      const map = "map" in result ? result.map as { sheets?: unknown[] } : undefined;
      return jsonResult({ ok: result.ok, sheets: map?.sheets ?? [], result });
    }
  );

  registerMcpTool(
    mcp,
    "excel.sheet.get_info",
    {
      title: "Get worksheet info",
      description: "Return worksheet metadata by sheet name.",
      inputSchema: { sheetName: z.string() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ sheetName }: { sheetName: string }) => jsonResult(await selectSheetInfo(sheetName))
  );

  registerMcpTool(
    mcp,
    "excel.sheet.get_used_range",
    {
      title: "Get worksheet used range",
      description: "Return the used range for a worksheet.",
      inputSchema: { sheetName: z.string() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ sheetName }: { sheetName: string }) => {
      const info = await selectSheetInfo(sheetName);
      return jsonResult({ ok: info.ok, usedRange: info.sheet?.usedRange, sheet: info.sheet });
    }
  );

  registerSheetOperation(mcp, "excel.sheet.create", {
    workbookId: z.string(),
    sheetName: z.string(),
    activate: z.boolean().optional()
  }, (args) => ({
    kind: "sheet.create",
    workbookId: args.workbookId as WorkbookId,
    sheetName: args.sheetName,
    activate: args.activate,
    destructiveLevel: "structure",
    reason: "MCP sheet create"
  }));

  registerSheetOperation(mcp, "excel.sheet.copy", {
    workbookId: z.string(),
    sourceSheetName: z.string(),
    newSheetName: z.string(),
    activate: z.boolean().optional()
  }, (args) => ({
    kind: "sheet.copy",
    workbookId: args.workbookId as WorkbookId,
    sourceSheetName: args.sourceSheetName,
    newSheetName: args.newSheetName,
    activate: args.activate,
    destructiveLevel: "structure",
    reason: "MCP sheet copy"
  }));

  registerSheetOperation(mcp, "excel.sheet.rename", {
    workbookId: z.string(),
    sheetName: z.string(),
    newSheetName: z.string()
  }, (args) => ({
    kind: "sheet.rename",
    workbookId: args.workbookId as WorkbookId,
    sheetName: args.sheetName,
    newSheetName: args.newSheetName,
    destructiveLevel: "structure",
    reason: "MCP sheet rename"
  }));

  for (const [name, destructiveLevel] of [
    ["excel.sheet.delete", "structure"],
    ["excel.sheet.hide", "structure"],
    ["excel.sheet.unhide", "structure"],
    ["excel.sheet.protect", "structure"],
    ["excel.sheet.unprotect", "structure"],
    ["excel.sheet.clear", "structure"]
  ] as const) {
    registerSheetOperation(mcp, name, {
      workbookId: z.string(),
      sheetName: z.string(),
      password: z.string().optional(),
      applyTo: z.enum(["all", "contents", "formats"]).optional()
    }, (args) => ({
      kind: name.replace("excel.", "") as ExcelOperation["kind"],
      workbookId: args.workbookId as WorkbookId,
      sheetName: args.sheetName,
      password: args.password,
      applyTo: args.applyTo,
      destructiveLevel,
      reason: `MCP ${name}`
    }));
  }

  registerSheetOperation(mcp, "excel.sheet.set_tab_color", {
    workbookId: z.string(),
    sheetName: z.string(),
    color: z.string()
  }, (args) => ({
    kind: "sheet.set_tab_color",
    workbookId: args.workbookId as WorkbookId,
    sheetName: args.sheetName,
    color: args.color,
    destructiveLevel: "format",
    reason: "MCP sheet tab color"
  }));
}

function registerRangeTools(mcp: McpServer): void {
  const readSchema = {
    workbookId: z.string(),
    sheetName: z.string(),
    address: z.string()
  };

  for (const name of [
    "excel.range.read_values",
    "excel.range.read_formulas",
    "excel.range.read_number_formats",
    "excel.range.read_display_text",
    "excel.range.read_styles"
  ]) {
    registerMcpTool(
      mcp,
      name,
      {
        title: name.replace(/^excel\./, "").replace(/\./g, " "),
        description: "Read a range facet using the full range snapshot path.",
        inputSchema: readSchema,
        annotations: {
          readOnlyHint: true,
          destructiveHint: false,
          openWorldHint: false
        }
      },
      async ({ workbookId, sheetName, address }: { workbookId: string; sheetName: string; address: string }) =>
        jsonResult(await readRangeSnapshot(workbookId, sheetName, address))
    );
  }

  for (const [name, method] of [
    ["excel.range.read_hyperlinks", "range.read_hyperlinks"],
    ["excel.range.read_comments", "range.read_comments"],
    ["excel.range.read_notes", "range.read_notes"],
    ["excel.range.read_merged_cells", "range.read_merged_cells"],
    ["excel.range.read_data_validation", "range.read_data_validation"],
    ["excel.range.read_conditional_formatting", "range.read_conditional_formatting"],
    ["excel.range.find_blank_cells", "range.find_blank_cells"],
    ["excel.range.find_errors", "range.find_errors"]
  ] as const) {
    registerMcpTool(
      mcp,
      name,
      {
        title: name.replace(/^excel\./, "").replace(/\./g, " "),
        description: "Read advanced range metadata from the connected Excel add-in.",
        inputSchema: readSchema,
        annotations: {
          readOnlyHint: true,
          destructiveHint: false,
          openWorldHint: false
        }
      },
      async (args: any) => jsonResult(await runtime.readRangeMetadata(method, rangeMetadataRequest(args)))
    );
  }

  registerMcpTool(
    mcp,
    "excel.range.search",
    {
      title: "Search Excel range",
      description: "Search a worksheet for text and return matching range areas.",
      inputSchema: {
        ...readSchema,
        text: z.string(),
        completeMatch: z.boolean().optional(),
        matchCase: z.boolean().optional(),
        searchDirection: z.enum(["Forward", "Backwards"]).optional()
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async (args: any) => jsonResult(await runtime.readRangeMetadata("range.search", rangeSearchRequest(args)))
  );

  registerMcpTool(
    mcp,
    "excel.range.read_full",
    {
      title: "Read full Excel range state",
      description: "Read values, formulas, text, number formats, and basic style fingerprint data for a range.",
      inputSchema: {
        workbookId: z.string(),
        sheetName: z.string(),
        address: z.string(),
        includeStyles: z.boolean().optional(),
        includeFormulas: z.boolean().optional()
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({
      workbookId,
      sheetName,
      address,
      includeStyles,
      includeFormulas
    }: {
      workbookId: string;
      sheetName: string;
      address: string;
      includeStyles?: boolean;
      includeFormulas?: boolean;
    }) => {
      const operation: ExcelOperation = {
        kind: "range.read_full",
        operationId: makeId<OperationId>("op"),
        workbookId: workbookId as WorkbookId,
        destructiveLevel: "none",
        reason: "MCP range read",
        target: {
          workbookId: workbookId as WorkbookId,
          sheetName,
          address
        }
      };
      if (includeStyles !== undefined) {
        operation.includeStyles = includeStyles;
      }
      if (includeFormulas !== undefined) {
        operation.includeFormulas = includeFormulas;
      }
      return jsonResult(
        await runtime.applyBatch({
          workbookId: workbookId as WorkbookId,
          mode: "apply",
          operations: [operation]
        })
      );
    }
  );

  registerRangeOperation(mcp, "excel.range.write_values", {
    workbookId: z.string(),
    sheetName: z.string(),
    address: z.string(),
    values: z.array(z.array(z.any()))
  }, (args) => ({
    kind: "range.write_values",
    workbookId: args.workbookId as WorkbookId,
    target: targetFromArgs(args),
    values: args.values,
    preserveFormats: true,
    destructiveLevel: "values",
    reason: "MCP range write values"
  }));

  registerRangeOperation(mcp, "excel.range.write_formulas", {
    workbookId: z.string(),
    sheetName: z.string(),
    address: z.string(),
    formulas: z.array(z.array(z.string().nullable()))
  }, (args) => ({
    kind: "range.write_formulas",
    workbookId: args.workbookId as WorkbookId,
    target: targetFromArgs(args),
    formulas: args.formulas,
    preserveFormats: true,
    destructiveLevel: "values",
    reason: "MCP range write formulas"
  }));

  registerRangeOperation(mcp, "excel.range.write_number_formats", {
    workbookId: z.string(),
    sheetName: z.string(),
    address: z.string(),
    numberFormat: z.array(z.array(z.string()))
  }, (args) => ({
    kind: "range.write_number_formats",
    workbookId: args.workbookId as WorkbookId,
    target: targetFromArgs(args),
    numberFormat: args.numberFormat,
    preserveValues: true,
    destructiveLevel: "format",
    reason: "MCP range write number formats"
  }));

  registerRangeOperation(mcp, "excel.range.write_styles", {
    workbookId: z.string(),
    sheetName: z.string(),
    address: z.string(),
    style: z.record(z.string(), z.any())
  }, (args) => ({
    kind: "range.write_styles",
    workbookId: args.workbookId as WorkbookId,
    target: targetFromArgs(args),
    style: args.style,
    preserveValues: true,
    destructiveLevel: "format",
    reason: "MCP range write styles"
  }));

  for (const name of ["excel.range.clear", "excel.range.clear_values", "excel.range.clear_formats", "excel.range.clear_values_keep_format"] as const) {
    registerRangeOperation(mcp, name, {
      workbookId: z.string(),
      sheetName: z.string(),
      address: z.string(),
      applyTo: z.enum(["all", "contents", "formats", "hyperlinks"]).optional()
    }, (args) => ({
      kind: name.replace("excel.", "") as ExcelOperation["kind"],
      workbookId: args.workbookId as WorkbookId,
      target: targetFromArgs(args),
      applyTo: args.applyTo,
      destructiveLevel: name.includes("format") && !name.includes("keep_format") ? "format" : "values",
      reason: `MCP ${name}`
    }));
  }

  for (const name of [
    "excel.range.insert_rows",
    "excel.range.delete_rows",
    "excel.range.insert_columns",
    "excel.range.delete_columns",
    "excel.range.autofit_columns",
    "excel.range.autofit_rows",
    "excel.range.merge",
    "excel.range.unmerge"
  ] as const) {
    registerRangeOperation(mcp, name, {
      workbookId: z.string(),
      sheetName: z.string(),
      address: z.string(),
      across: z.boolean().optional()
    }, (args) => ({
      kind: name.replace("excel.", "") as ExcelOperation["kind"],
      workbookId: args.workbookId as WorkbookId,
      target: targetFromArgs(args),
      across: args.across,
      destructiveLevel: name.includes("autofit") ? "format" : "structure",
      reason: `MCP ${name}`
    }));
  }

  for (const name of ["excel.range.copy", "excel.range.move"] as const) {
    registerRangeOperation(mcp, name, {
      workbookId: z.string(),
      sourceSheetName: z.string(),
      sourceAddress: z.string(),
      targetSheetName: z.string(),
      targetAddress: z.string(),
      copyType: z.enum(["all", "values", "formats", "formulas"]).optional()
    }, (args) => ({
      kind: name.replace("excel.", "") as ExcelOperation["kind"],
      workbookId: args.workbookId as WorkbookId,
      source: {
        workbookId: args.workbookId as WorkbookId,
        sheetName: args.sourceSheetName,
        address: args.sourceAddress
      },
      target: {
        workbookId: args.workbookId as WorkbookId,
        sheetName: args.targetSheetName,
        address: args.targetAddress
      },
      copyType: args.copyType,
      destructiveLevel: name.endsWith(".move") ? "values" : "none",
      reason: `MCP ${name}`
    }));
  }
}

function registerBatchTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.batch.validate",
    {
      title: "Validate Excel batch",
      description: "Compile and validate a batch without sending it to Excel.",
      inputSchema: {
        workbookId: z.string(),
        operations: z.array(z.any())
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ workbookId, operations }: { workbookId: string; operations: unknown[] }) => {
      const request: BatchRequest = {
        workbookId: workbookId as WorkbookId,
        mode: "validate",
        operations: operations as ExcelOperation[]
      };
      return jsonResult(await runtime.compileBatch(request));
    }
  );

  registerMcpTool(
    mcp,
    "excel.batch.dry_run",
    {
      title: "Dry-run Excel batch",
      description: "Compile a batch and report backups, touched ranges, and estimated changes.",
      inputSchema: {
        workbookId: z.string(),
        operations: z.array(z.any())
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ workbookId, operations }: { workbookId: string; operations: unknown[] }) => {
      const request: BatchRequest = {
        workbookId: workbookId as WorkbookId,
        mode: "dry_run",
        operations: operations as ExcelOperation[]
      };
      return jsonResult(await runtime.compileBatch(request));
    }
  );

  registerMcpTool(
    mcp,
    "excel.batch.apply",
    {
      title: "Apply Excel batch",
      description: "Apply a batch through snapshots, backups, target conflict checks, and Office.js execution.",
      inputSchema: {
        workbookId: z.string(),
        operations: z.array(z.any()),
        confirmationToken: z.string().optional(),
        expectedTargetFingerprints: z.array(z.any()).optional(),
        agentId: z.string().optional(),
        agentName: z.string().optional(),
        taskId: z.string().optional(),
        role: z.string().optional()
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        openWorldHint: false
      }
    },
    async ({
      workbookId,
      operations,
      confirmationToken,
      expectedTargetFingerprints,
      agentId,
      agentName,
      taskId,
      role
    }: {
      workbookId: string;
      operations: unknown[];
      confirmationToken?: string;
      expectedTargetFingerprints?: unknown[];
      agentId?: string;
      agentName?: string;
      taskId?: string;
      role?: string;
    }) => {
      const request: BatchRequest = {
        workbookId: workbookId as WorkbookId,
        mode: "apply",
        operations: operations as ExcelOperation[]
      };
      if (confirmationToken !== undefined) {
        request.confirmationToken = confirmationToken;
      }
      if (expectedTargetFingerprints !== undefined) {
        request.expectedTargetFingerprints = expectedTargetFingerprints as NonNullable<
          BatchRequest["expectedTargetFingerprints"]
        >;
      }
      if (agentId !== undefined) {
        request.agentId = agentId as AgentId;
      }
      if (agentName !== undefined) {
        request.agentName = agentName;
      }
      if (taskId !== undefined) {
        request.taskId = taskId as TaskId;
      }
      if (role !== undefined) {
        request.role = role;
      }
      return jsonResult(await runtime.applyBatch(request));
    }
  );
}

function registerPlanTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.plan.create",
    {
      title: "Create Excel plan",
      description: "Create a reversible plan from proposed Excel operations.",
      inputSchema: {
        workbookId: z.string(),
        goal: z.string(),
        operations: z.array(z.any()),
        agentId: z.string().optional(),
        agentName: z.string().optional(),
        taskId: z.string().optional(),
        role: z.string().optional()
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({
      workbookId,
      goal,
      operations,
      agentId,
      agentName,
      taskId,
      role
    }: {
      workbookId: string;
      goal: string;
      operations: unknown[];
      agentId?: string;
      agentName?: string;
      taskId?: string;
      role?: string;
    }) =>
      jsonResult(
        runtime.createPlan({
          workbookId: workbookId as WorkbookId,
          goal,
          operations: operations as ExcelOperation[],
          agentId: agentId as AgentId | undefined,
          agentName,
          taskId: taskId as TaskId | undefined,
          role
        })
      )
  );

  registerMcpTool(
    mcp,
    "excel.plan.preview",
    {
      title: "Preview Excel plan",
      description: "Preview a plan and capture target-region fingerprints when Excel is connected.",
      inputSchema: {
        planId: z.string()
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ planId }: { planId: string }) => jsonResult(await runtime.previewPlan(planId as PlanId))
  );

  registerMcpTool(
    mcp,
    "excel.plan.refresh_preview",
    {
      title: "Refresh Excel plan preview",
      description: "Refresh plan target fingerprints only when target ranges have not changed since preview.",
      inputSchema: {
        planId: z.string()
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ planId }: { planId: string }) => jsonResult(await runtime.refreshPlanPreview(planId as PlanId))
  );

  registerMcpTool(
    mcp,
    "excel.plan.rebase",
    {
      title: "Rebase Excel plan",
      description: "Safely rebase a plan by refreshing fingerprints when target ranges are unchanged.",
      inputSchema: {
        planId: z.string()
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ planId }: { planId: string }) => jsonResult(await runtime.rebasePlan(planId as PlanId))
  );

  registerMcpTool(
    mcp,
    "excel.plan.apply",
    {
      title: "Apply Excel plan",
      description: "Apply a previewed plan if target-region fingerprints still match.",
      inputSchema: {
        planId: z.string(),
        confirmationToken: z.string().optional()
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        openWorldHint: false
      }
    },
    async ({ planId, confirmationToken }: { planId: string; confirmationToken?: string }) =>
      jsonResult(await runtime.applyPlan(planId as PlanId, confirmationToken))
  );

  registerMcpTool(
    mcp,
    "excel.plan.rollback",
    {
      title: "Rollback Excel plan",
      description: "Rollback an applied plan using captured region snapshots and created-sheet cleanup.",
      inputSchema: {
        planId: z.string(),
        confirmationToken: z.string().optional()
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        openWorldHint: false
      }
    },
    async ({ planId, confirmationToken }: { planId: string; confirmationToken?: string }) =>
      jsonResult(await runtime.rollbackPlan(planId as PlanId, confirmationToken))
  );
}

function registerTemplateTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.template.register",
    {
      title: "Register Excel template",
      description: "Capture and register a sheet template fingerprint for style, formula, and layout preservation.",
      inputSchema: {
        workbookId: z.string(),
        name: z.string(),
        scope: z.enum(["workbook", "local"]).default("workbook"),
        sourceSheetName: z.string(),
        dataRegions: z.array(z.string()).default([])
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({
      workbookId,
      name,
      scope,
      sourceSheetName,
      dataRegions
    }: {
      workbookId: string;
      name: string;
      scope: "workbook" | "local";
      sourceSheetName: string;
      dataRegions: string[];
    }) =>
      jsonResult(
        await runtime.registerTemplate({
          workbookId: workbookId as WorkbookId,
          name,
          scope,
          sourceSheetName,
          dataRegions
        })
      )
  );

  registerMcpTool(
    mcp,
    "excel.template.get",
    {
      title: "Get Excel template",
      description: "Return a registered template including fingerprint payload.",
      inputSchema: {
        templateId: z.string()
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ templateId }: { templateId: string }) => jsonResult(runtime.getTemplate(templateId as TemplateId))
  );

  registerMcpTool(
    mcp,
    "excel.template.unregister",
    {
      title: "Unregister Excel template",
      description: "Remove a registered template from the local runtime registry.",
      inputSchema: {
        templateId: z.string()
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ templateId }: { templateId: string }) => jsonResult(runtime.unregisterTemplate(templateId as TemplateId))
  );

  registerMcpTool(
    mcp,
    "excel.template.detect_templates",
    {
      title: "Detect Excel templates",
      description: "Return candidate template sheets from the active workbook.",
      inputSchema: {
        workbookId: z.string()
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ workbookId }: { workbookId: string }) => jsonResult(await runtime.detectTemplates(workbookId as WorkbookId))
  );

  registerMcpTool(
    mcp,
    "excel.template.infer_regions",
    {
      title: "Infer Excel template regions",
      description: "Return declared and inferred data regions for a registered template.",
      inputSchema: {
        templateId: z.string()
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ templateId }: { templateId: string }) => jsonResult(runtime.inferTemplateRegions(templateId as TemplateId))
  );

  registerMcpTool(
    mcp,
    "excel.template.create_sheet_from_template",
    {
      title: "Create sheet from template",
      description: "Copy a registered template sheet and clear declared data regions.",
      inputSchema: {
        workbookId: z.string(),
        templateId: z.string(),
        newSheetName: z.string(),
        clearDataRegions: z.boolean().default(true)
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        openWorldHint: false
      }
    },
    async ({
      workbookId,
      templateId,
      newSheetName,
      clearDataRegions
    }: {
      workbookId: string;
      templateId: string;
      newSheetName: string;
      clearDataRegions: boolean;
    }) =>
      jsonResult(
        await applySingleOperation(workbookId, {
          kind: "template.create_sheet_from_template",
          workbookId: workbookId as WorkbookId,
          templateId: templateId as TemplateId,
          newSheetName,
          clearDataRegions,
          destructiveLevel: "structure",
          reason: "MCP create sheet from template"
        })
      )
  );

  registerMcpTool(
    mcp,
    "excel.template.validate_sheet_against_template",
    {
      title: "Validate sheet against template",
      description: "Compare a target sheet against a registered template fingerprint.",
      inputSchema: {
        workbookId: z.string(),
        templateId: z.string(),
        targetSheetName: z.string()
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ workbookId, templateId, targetSheetName }: { workbookId: string; templateId: string; targetSheetName: string }) =>
      jsonResult(
        await runtime.validateSheetAgainstTemplate({
          workbookId: workbookId as WorkbookId,
          templateId: templateId as TemplateId,
          targetSheetName
        })
      )
  );

  registerMcpTool(
    mcp,
    "excel.template.clear_data_regions",
    {
      title: "Clear template data regions",
      description: "Clear declared data regions on a target sheet while preserving formats.",
      inputSchema: {
        workbookId: z.string(),
        templateId: z.string(),
        targetSheetName: z.string()
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        openWorldHint: false
      }
    },
    async ({ workbookId, templateId, targetSheetName }: { workbookId: string; templateId: string; targetSheetName: string }) => {
      const templateResult = runtime.getTemplate(templateId as TemplateId);
      if (!templateResult.ok || !("template" in templateResult)) {
        return jsonResult(templateResult);
      }
      const operations: ExcelOperation[] = templateResult.template.dataRegions.map((address) => ({
        kind: "range.clear_values_keep_format",
        operationId: makeId<OperationId>("op"),
        workbookId: workbookId as WorkbookId,
        destructiveLevel: "values",
        reason: `Clear data region from template ${templateId}`,
        target: {
          workbookId: workbookId as WorkbookId,
          sheetName: targetSheetName,
          address
        }
      }));
      return jsonResult(await runtime.applyBatch({ workbookId: workbookId as WorkbookId, mode: "apply", operations }));
    }
  );

  registerMcpTool(
    mcp,
    "excel.template.fill_regions",
    {
      title: "Fill template regions",
      description: "Write values to target sheet regions while preserving existing formats.",
      inputSchema: {
        workbookId: z.string(),
        targetSheetName: z.string(),
        regions: z.array(
          z.object({
            address: z.string(),
            values: z.array(z.array(z.any()))
          })
        )
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        openWorldHint: false
      }
    },
    async ({ workbookId, targetSheetName, regions }: { workbookId: string; targetSheetName: string; regions: Array<{ address: string; values: unknown[][] }> }) => {
      const operations: ExcelOperation[] = regions.map((region) => ({
        kind: "range.write_values",
        operationId: makeId<OperationId>("op"),
        workbookId: workbookId as WorkbookId,
        destructiveLevel: "values",
        reason: "Fill template region",
        target: {
          workbookId: workbookId as WorkbookId,
          sheetName: targetSheetName,
          address: region.address
        },
        values: region.values as any,
        preserveFormats: true
      }));
      return jsonResult(await runtime.applyBatch({ workbookId: workbookId as WorkbookId, mode: "apply", operations }));
    }
  );

  registerMcpTool(
    mcp,
    "excel.template.repair_sheet_from_template",
    {
      title: "Repair sheet from template",
      description: "Repair target sheet styles, formulas, layout, or data regions from a registered template.",
      inputSchema: templateRepairSchema(),
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        openWorldHint: false
      }
    },
    async (args: any) => jsonResult(await repairTemplateFromArgs(args))
  );

  registerMcpTool(
    mcp,
    "excel.template.list",
    {
      title: "List Excel templates",
      description: "List local and workbook-scoped registered templates.",
      inputSchema: {
        workbookId: z.string().optional()
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ workbookId }: { workbookId?: string }) => jsonResult(runtime.listTemplates(workbookId as WorkbookId | undefined))
  );
}

function registerStyleTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.style.get_fingerprint",
    {
      title: "Get style fingerprint",
      description: "Capture a granular style fingerprint for a sheet or address.",
      inputSchema: {
        workbookId: z.string(),
        sheetName: z.string(),
        address: z.string().optional(),
        maxCellSamples: z.number().int().positive().optional()
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) =>
      jsonResult(
        await runtime.getStyleFingerprint({
          workbookId: args.workbookId as WorkbookId,
          sheetName: args.sheetName,
          ...(args.address !== undefined ? { address: args.address } : {}),
          ...(args.maxCellSamples !== undefined ? { maxCellSamples: args.maxCellSamples } : {})
        })
      )
  );

  registerMcpTool(
    mcp,
    "excel.style.compare_fingerprint",
    {
      title: "Compare style fingerprints",
      description: "Compare source and target style fingerprints by dimension.",
      inputSchema: styleCompareSchema(),
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.compareStyleFingerprints(styleCompareRequest(args)))
  );

  registerMcpTool(
    mcp,
    "excel.style.copy_from_template",
    {
      title: "Copy style from template",
      description: "Repair only styles from a registered template.",
      inputSchema: templateRepairSchema(),
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => jsonResult(await repairTemplateFromArgs({ ...args, repair: ["styles"] }))
  );

  registerMcpTool(
    mcp,
    "excel.style.apply_style",
    {
      title: "Apply range style",
      description: "Apply direct fill, font, alignment, row height, and column width properties to a range.",
      inputSchema: {
        workbookId: z.string(),
        sheetName: z.string(),
        address: z.string(),
        style: z.record(z.string(), z.unknown())
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => {
      const operation: ExcelOperation = {
        kind: "range.write_styles",
        operationId: makeId<OperationId>("op"),
        workbookId: args.workbookId as WorkbookId,
        destructiveLevel: "format",
        reason: "MCP style apply",
        target: {
          workbookId: args.workbookId as WorkbookId,
          sheetName: args.sheetName,
          address: args.address
        },
        style: args.style,
        preserveValues: true
      };
      return jsonResult(await runtime.applyBatch({ workbookId: args.workbookId as WorkbookId, mode: "apply", operations: [operation] }));
    }
  );

  for (const name of ["excel.style.validate_consistency", "excel.style.repair_consistency"] as const) {
    registerMcpTool(
      mcp,
      name,
      {
        title: name.replace(/^excel\./, "").replace(/\./g, " "),
        description: "Validate or repair style consistency against a registered template.",
        inputSchema: templateRepairSchema(),
        annotations: { readOnlyHint: name.includes("validate"), destructiveHint: name.includes("repair"), openWorldHint: false }
      },
      async (args: any) =>
        jsonResult(
          name.includes("validate")
            ? await runtime.validateSheetAgainstTemplate({
                workbookId: args.workbookId as WorkbookId,
                templateId: args.templateId as TemplateId,
                targetSheetName: args.targetSheetName
              })
            : await repairTemplateFromArgs({ ...args, repair: ["styles"] })
        )
    );
  }

  for (const [name, dimension] of Object.entries(STYLE_COPY_TOOL_DIMENSIONS)) {
    registerMcpTool(
      mcp,
      name,
      {
        title: name.replace(/^excel\.style\./, "").replace(/_/g, " "),
        description: `Copy ${dimension} styling from one sheet/range to another with backup and post-copy validation.`,
        inputSchema: styleCopySchema(),
        annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
      },
      async (args: any) => jsonResult(await runtime.copyStyleDimensions({ ...styleCopyRequest(args), dimensions: [dimension] }))
    );
  }

  registerMcpTool(
    mcp,
    "excel.style.get_theme",
    {
      title: "Get workbook theme",
      description: "Report workbook theme capability status.",
      inputSchema: {
        workbookId: z.string()
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ workbookId }: { workbookId: string }) =>
      jsonResult({
        ok: false,
        workbookId,
        error: {
          code: "CAPABILITY_UNAVAILABLE",
          message: "Office.js does not expose a safe cross-platform workbook theme snapshot in the current implementation.",
          retryable: false
        }
      })
  );

  registerMcpTool(
    mcp,
    "excel.style.apply_theme",
    {
      title: "Apply workbook theme",
      description: "Report workbook theme apply capability status.",
      inputSchema: {
        workbookId: z.string(),
        theme: z.record(z.string(), z.unknown())
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async ({ workbookId }: { workbookId: string }) =>
      jsonResult({
        ok: false,
        workbookId,
        error: {
          code: "CAPABILITY_UNAVAILABLE",
          message: "Theme apply is not enabled until workbook theme snapshots can be captured and replayed deterministically.",
          retryable: false
        }
      })
  );
}

function styleCompareSchema() {
  return {
    workbookId: z.string(),
    sourceSheetName: z.string(),
    targetSheetName: z.string(),
    sourceAddress: z.string().optional(),
    targetAddress: z.string().optional(),
    dimensions: z.array(z.enum(STYLE_DIMENSIONS)).optional(),
    maxCellSamples: z.number().int().positive().optional()
  };
}

function styleCopySchema() {
  return {
    workbookId: z.string(),
    sourceSheetName: z.string(),
    targetSheetName: z.string(),
    sourceAddress: z.string().optional(),
    targetAddress: z.string().optional()
  };
}

function styleCompareRequest(args: any): {
  workbookId: WorkbookId;
  sourceSheetName: string;
  targetSheetName: string;
  sourceAddress?: string;
  targetAddress?: string;
  dimensions?: StyleDimension[];
  maxCellSamples?: number;
} {
  return {
    workbookId: args.workbookId as WorkbookId,
    sourceSheetName: args.sourceSheetName,
    targetSheetName: args.targetSheetName,
    ...(args.sourceAddress !== undefined ? { sourceAddress: args.sourceAddress } : {}),
    ...(args.targetAddress !== undefined ? { targetAddress: args.targetAddress } : {}),
    ...(args.dimensions !== undefined ? { dimensions: args.dimensions as StyleDimension[] } : {}),
    ...(args.maxCellSamples !== undefined ? { maxCellSamples: args.maxCellSamples } : {})
  };
}

function styleCopyRequest(args: any): {
  workbookId: WorkbookId;
  sourceSheetName: string;
  targetSheetName: string;
  sourceAddress?: string;
  targetAddress?: string;
  dimensions: StyleDimension[];
} {
  return {
    workbookId: args.workbookId as WorkbookId,
    sourceSheetName: args.sourceSheetName,
    targetSheetName: args.targetSheetName,
    ...(args.sourceAddress !== undefined ? { sourceAddress: args.sourceAddress } : {}),
    ...(args.targetAddress !== undefined ? { targetAddress: args.targetAddress } : {}),
    dimensions: []
  };
}

function registerFormulaTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.formula.read_patterns",
    {
      title: "Read formula patterns",
      description: "Capture formulas, R1C1 pattern hashes, and pattern groups for a sheet or range.",
      inputSchema: formulaPatternSchema(),
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.readFormulaPatterns(formulaPatternRequest(args)))
  );

  registerMcpTool(
    mcp,
    "excel.formula.copy_patterns",
    {
      title: "Copy formula patterns",
      description: "Copy formulas from a source sheet/range to a target sheet/range with backup and validation.",
      inputSchema: formulaCopySchema(),
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.copyFormulaPatterns(formulaCopyRequest(args)))
  );

  for (const [name, direction] of [
    ["excel.formula.fill_down", "down"],
    ["excel.formula.fill_right", "right"]
  ] as const) {
    registerMcpTool(
      mcp,
      name,
      {
        title: name.replace(/^excel\.formula\./, "").replace(/_/g, " "),
        description: `Fill formulas ${direction} using R1C1 pattern semantics.`,
        inputSchema: formulaFillSchema(),
        annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
      },
      async (args: any) => jsonResult(await runtime.fillFormulaPattern({ ...formulaFillRequest(args), direction }))
    );
  }

  registerMcpTool(
    mcp,
    "excel.formula.validate",
    {
      title: "Validate formulas",
      description: "Validate formula errors in a workbook, sheet, or range.",
      inputSchema: validationRangeSchema(),
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) =>
      jsonResult(
        await runtime.validateFormulas({
          workbookId: args.workbookId as WorkbookId,
          ...(args.sheetName !== undefined ? { sheetName: args.sheetName } : {}),
          ...(args.address !== undefined ? { address: args.address } : {})
        })
      )
  );

  registerMcpTool(
    mcp,
    "excel.formula.validate_against_template",
    {
      title: "Validate formulas against template",
      description: "Validate formula consistency against a registered template.",
      inputSchema: templateRepairSchema(),
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) =>
      jsonResult(
        await runtime.validateSheetAgainstTemplate({
          workbookId: args.workbookId as WorkbookId,
          templateId: args.templateId as TemplateId,
          targetSheetName: args.targetSheetName
        })
      )
  );

  registerMcpTool(
    mcp,
    "excel.formula.repair_patterns",
    {
      title: "Repair formula patterns",
      description: "Repair formulas from a registered template.",
      inputSchema: templateRepairSchema(),
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => jsonResult(await repairTemplateFromArgs({ ...args, repair: ["formulas"] }))
  );

  registerMcpTool(
    mcp,
    "excel.formula.find_errors",
    {
      title: "Find formula errors",
      description: "Find cells with formula errors in a sheet/range.",
      inputSchema: validationRangeSchema(),
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) =>
      jsonResult(
        await runtime.validateFormulas({
          workbookId: args.workbookId as WorkbookId,
          ...(args.sheetName !== undefined ? { sheetName: args.sheetName } : {}),
          ...(args.address !== undefined ? { address: args.address } : {})
        })
      )
  );

  registerMcpTool(
    mcp,
    "excel.formula.find_circular_references",
    {
      title: "Find circular references",
      description: "Report circular-reference detection capability status.",
      inputSchema: validationRangeSchema(),
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(capabilityUnavailable(args.workbookId, "FORMULA_CIRCULAR_REFERENCES_UNAVAILABLE", "Office.js does not expose deterministic circular-reference enumeration in this runtime yet."))
  );

  registerMcpTool(
    mcp,
    "excel.formula.get_dependency_graph",
    {
      title: "Get formula dependency graph",
      description: "Parse formulas in a sheet or range and return precedent dependency edges.",
      inputSchema: formulaPatternSchema(),
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.getFormulaDependencyGraph(formulaPatternRequest(args)))
  );

  registerMcpTool(
    mcp,
    "excel.formula.trace_precedents",
    {
      title: "Trace formula precedents",
      description: "Parse a formula cell and return referenced precedent ranges.",
      inputSchema: {
        workbookId: z.string(),
        sheetName: z.string(),
        address: z.string()
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.traceFormulaPrecedents(formulaPatternRequest(args)))
  );

  registerMcpTool(
    mcp,
    "excel.formula.trace_dependents",
    {
      title: "Trace formula dependents",
      description: "Parse formulas on a sheet and return formula cells that depend on the target range.",
      inputSchema: {
        workbookId: z.string(),
        sheetName: z.string(),
        address: z.string()
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.traceFormulaDependents(formulaPatternRequest(args)))
  );

  registerMcpTool(
    mcp,
    "excel.formula.convert_to_values",
    {
      title: "Convert formulas to values",
      description: "Replace formulas in a sheet/range with their current calculated values, with backup.",
      inputSchema: formulaPatternSchema(),
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.convertFormulasToValues(formulaPatternRequest(args)))
  );

  registerMcpTool(
    mcp,
    "excel.formula.recalculate",
    {
      title: "Recalculate formulas",
      description: "Run workbook recalculation.",
      inputSchema: {
        workbookId: z.string(),
        calculationType: z.enum(["full", "recalculate"]).optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async ({ workbookId, calculationType }: { workbookId: string; calculationType?: "full" | "recalculate" }) =>
      jsonResult(await runtime.calculateWorkbook(workbookId as WorkbookId, calculationType))
  );

  registerMcpTool(
    mcp,
    "excel.formula.explain",
    {
      title: "Explain formula",
      description: "Return a lightweight parse summary for a formula string or the first formula in a range.",
      inputSchema: {
        workbookId: z.string(),
        sheetName: z.string().optional(),
        address: z.string().optional(),
        formula: z.string().optional()
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await explainFormula(args))
  );
}

function formulaPatternSchema() {
  return {
    workbookId: z.string(),
    sheetName: z.string(),
    address: z.string().optional()
  };
}

function formulaCopySchema() {
  return {
    workbookId: z.string(),
    sourceSheetName: z.string(),
    targetSheetName: z.string(),
    sourceAddress: z.string().optional(),
    targetAddress: z.string().optional()
  };
}

function formulaFillSchema() {
  return {
    workbookId: z.string(),
    sheetName: z.string(),
    sourceAddress: z.string(),
    targetAddress: z.string()
  };
}

function validationRangeSchema() {
  return {
    workbookId: z.string(),
    sheetName: z.string().optional(),
    address: z.string().optional()
  };
}

function formulaPatternRequest(args: { workbookId: string; sheetName: string; address?: string }): FormulaPatternRequest {
  return {
    workbookId: args.workbookId as WorkbookId,
    sheetName: args.sheetName,
    ...(args.address !== undefined ? { address: args.address } : {})
  };
}

function formulaCopyRequest(args: {
  workbookId: string;
  sourceSheetName: string;
  targetSheetName: string;
  sourceAddress?: string;
  targetAddress?: string;
}): FormulaCopyPatternsRequest {
  return {
    workbookId: args.workbookId as WorkbookId,
    sourceSheetName: args.sourceSheetName,
    targetSheetName: args.targetSheetName,
    ...(args.sourceAddress !== undefined ? { sourceAddress: args.sourceAddress } : {}),
    ...(args.targetAddress !== undefined ? { targetAddress: args.targetAddress } : {})
  };
}

function formulaFillRequest(args: { workbookId: string; sheetName: string; sourceAddress: string; targetAddress: string }): FormulaFillRequest {
  return {
    workbookId: args.workbookId as WorkbookId,
    sheetName: args.sheetName,
    sourceAddress: args.sourceAddress,
    targetAddress: args.targetAddress,
    direction: "down"
  };
}

function capabilityUnavailable(workbookId: string, code: string, message: string) {
  return {
    ok: false,
    workbookId,
    error: {
      code: "CAPABILITY_UNAVAILABLE",
      message,
      retryable: false,
      details: { reasonCode: code }
    }
  };
}

async function explainFormula(args: { workbookId: string; sheetName?: string; address?: string; formula?: string }) {
  let formula = args.formula;
  if (!formula && args.sheetName && args.address) {
    const result = await runtime.readFormulaPatterns({
      workbookId: args.workbookId as WorkbookId,
      sheetName: args.sheetName,
      address: args.address
    });
    if (!result.ok) {
      return result;
    }
    formula = result.patterns.cells[0]?.formula;
  }
  if (!formula) {
    return {
      ok: false,
      error: {
        code: "FORMULA_REQUIRED",
        message: "Provide a formula string or a sheetName/address containing at least one formula.",
        retryable: false
      }
    };
  }
  const normalized = formula.startsWith("=") ? formula : `=${formula}`;
  const functions = [...normalized.matchAll(/\b([A-Z][A-Z0-9_.]*)\s*\(/gi)].map((match) => match[1]!.toUpperCase());
  const references = [...normalized.matchAll(/(?:'[^']+'|[A-Z_][A-Z0-9_ ]*!|\b)?\$?[A-Z]{1,3}\$?\d+(?::\$?[A-Z]{1,3}\$?\d+)?/gi)].map((match) => match[0]);
  return {
    ok: true,
    formula: normalized,
    summary: {
      functions: [...new Set(functions)],
      references: [...new Set(references)],
      hasExternalReference: /\[[^\]]+\]/.test(normalized),
      hasStructuredReference: /\[[#@\w ,:[\]]+\]/.test(normalized),
      hasVolatileFunction: functions.some((fn) => ["NOW", "TODAY", "RAND", "RANDBETWEEN", "OFFSET", "INDIRECT"].includes(fn))
    }
  };
}

function registerTableTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.table.list",
    {
      title: "List Excel tables",
      description: "List structured tables in the active workbook.",
      inputSchema: { workbookId: z.string() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ workbookId }: { workbookId: string }) => jsonResult(await runtime.listTables(workbookId as WorkbookId))
  );

  registerMcpTool(
    mcp,
    "excel.table.get_info",
    {
      title: "Get Excel table info",
      description: "Return table range, columns, style, filter, and sort metadata.",
      inputSchema: tableSelectorSchema(),
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.getTableInfo(tableSelector(args)))
  );

  registerMcpTool(
    mcp,
    "excel.table.read",
    {
      title: "Read Excel table",
      description: "Read table headers, values, formulas, text, number formats, and metadata.",
      inputSchema: tableSelectorSchema(),
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.readTable(tableSelector(args)))
  );

  registerMcpTool(
    mcp,
    "excel.table.create",
    {
      title: "Create Excel table",
      description: "Create a structured table from a range, optionally writing values first.",
      inputSchema: {
        workbookId: z.string(),
        sheetName: z.string(),
        address: z.string(),
        tableName: z.string().optional(),
        hasHeaders: z.boolean().default(true),
        values: z.array(z.array(z.any())).optional(),
        style: z.string().optional(),
        showTotals: z.boolean().optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.createTable(tableCreateRequest(args)))
  );

  registerMcpTool(
    mcp,
    "excel.table.resize",
    {
      title: "Resize Excel table",
      description: "Resize a structured table to a new range.",
      inputSchema: { ...tableSelectorSchema(), address: z.string() },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.resizeTable({ ...tableSelector(args), address: args.address } as TableResizeRequest))
  );

  registerMcpTool(
    mcp,
    "excel.table.append_rows",
    {
      title: "Append Excel table rows",
      description: "Append one or more rows to a structured table.",
      inputSchema: {
        ...tableSelectorSchema(),
        values: z.array(z.array(z.any())),
        index: z.number().int().optional(),
        alwaysInsert: z.boolean().optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) =>
      jsonResult(
        await runtime.appendTableRows({
          ...tableSelector(args),
          values: args.values,
          index: args.index,
          alwaysInsert: args.alwaysInsert
        } as TableAppendRowsRequest)
      )
  );

  registerMcpTool(
    mcp,
    "excel.table.update_rows",
    {
      title: "Update Excel table rows",
      description: "Update table rows by zero-based table-row index.",
      inputSchema: {
        ...tableSelectorSchema(),
        rows: z.array(z.object({ index: z.number().int().min(0), values: z.array(z.any()) }))
      },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.updateTableRows({ ...tableSelector(args), rows: args.rows } as TableUpdateRowsRequest))
  );

  registerMcpTool(
    mcp,
    "excel.table.clear_data_keep_formulas",
    {
      title: "Clear table data keep formulas",
      description: "Clear constants in the table data body while preserving formulas.",
      inputSchema: tableSelectorSchema(),
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.clearTableDataKeepFormulas(tableSelector(args)))
  );

  registerMcpTool(
    mcp,
    "excel.table.clear_filters",
    {
      title: "Clear table filters",
      description: "Clear all filters on a structured table.",
      inputSchema: tableSelectorSchema(),
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.clearTableFilters(tableSelector(args)))
  );

  registerMcpTool(
    mcp,
    "excel.table.apply_filters",
    {
      title: "Apply table filters",
      description: "Apply Office.js filter criteria to table columns.",
      inputSchema: tableFilterSchema(),
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.applyTableFilters({ ...tableSelector(args), filters: args.filters } as TableApplyFiltersRequest))
  );

  registerMcpTool(
    mcp,
    "excel.table.preserve_filters",
    {
      title: "Preserve table filters",
      description: "Reapply provided filter criteria to a table.",
      inputSchema: tableFilterSchema(),
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.applyTableFilters({ ...tableSelector(args), filters: args.filters } as TableApplyFiltersRequest))
  );

  registerMcpTool(
    mcp,
    "excel.table.sort",
    {
      title: "Sort Excel table",
      description: "Apply table sort fields.",
      inputSchema: tableSortSchema(),
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) =>
      jsonResult(await runtime.sortTable({ ...tableSelector(args), fields: args.fields, matchCase: args.matchCase, method: args.method } as TableSortRequest))
  );

  registerMcpTool(
    mcp,
    "excel.table.set_total_row",
    {
      title: "Set table total row",
      description: "Show or hide the table total row.",
      inputSchema: { ...tableSelectorSchema(), showTotals: z.boolean() },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.setTableTotalRow({ ...tableSelector(args), showTotals: args.showTotals } as TableSetTotalRowRequest))
  );

  registerMcpTool(
    mcp,
    "excel.table.set_style",
    {
      title: "Set table style",
      description: "Apply an Excel table style.",
      inputSchema: { ...tableSelectorSchema(), style: z.string() },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.setTableStyle({ ...tableSelector(args), style: args.style } as TableSetStyleRequest))
  );

  registerMcpTool(
    mcp,
    "excel.table.copy_structure",
    {
      title: "Copy table structure",
      description: "Copy table headers and optional style/totals to a new target table.",
      inputSchema: {
        ...tableSelectorSchema(),
        targetSheetName: z.string(),
        targetAddress: z.string(),
        newTableName: z.string().optional(),
        includeStyle: z.boolean().optional(),
        includeTotals: z.boolean().optional(),
        includeFilters: z.boolean().optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) =>
      jsonResult(
        await runtime.copyTableStructure({
          ...tableSelector(args),
          targetSheetName: args.targetSheetName,
          targetAddress: args.targetAddress,
          newTableName: args.newTableName,
          includeStyle: args.includeStyle,
          includeTotals: args.includeTotals,
          includeFilters: args.includeFilters
        } as TableCopyStructureRequest)
      )
  );

  registerMcpTool(
    mcp,
    "excel.table.validate_against_template",
    {
      title: "Validate table against template",
      description: "Compare current table metadata with a registered template table fingerprint.",
      inputSchema: { ...tableSelectorSchema(), templateId: z.string() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.validateTableAgainstTemplate({ ...tableSelector(args), templateId: args.templateId as TemplateId }))
  );
}

function registerFilterTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.filter.get_filters",
    {
      title: "Get table filters",
      description: "Return filter metadata for a structured table.",
      inputSchema: tableSelectorSchema(),
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => {
      const result = await runtime.getTableInfo(tableSelector(args));
      const info = (result as { info?: { filters?: unknown } }).info;
      return jsonResult({ ok: Boolean(info), filters: info?.filters, result });
    }
  );

  registerMcpTool(
    mcp,
    "excel.filter.apply",
    {
      title: "Apply table filters",
      description: "Apply filters to a structured table.",
      inputSchema: tableFilterSchema(),
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.applyTableFilters({ ...tableSelector(args), filters: args.filters } as TableApplyFiltersRequest))
  );

  registerMcpTool(
    mcp,
    "excel.filter.clear",
    {
      title: "Clear table filters",
      description: "Clear table filters.",
      inputSchema: tableSelectorSchema(),
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.clearTableFilters(tableSelector(args)))
  );

  registerMcpTool(
    mcp,
    "excel.filter.preserve_from_template",
    {
      title: "Preserve filters from template",
      description: "Apply provided template filter criteria to a table.",
      inputSchema: tableFilterSchema(),
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.applyTableFilters({ ...tableSelector(args), filters: args.filters } as TableApplyFiltersRequest))
  );

  registerMcpTool(
    mcp,
    "excel.filter.validate",
    {
      title: "Validate table filters",
      description: "Return current filter state for validation by the agent.",
      inputSchema: tableSelectorSchema(),
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.getTableInfo(tableSelector(args)))
  );
}

function registerSortTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.sort.apply",
    {
      title: "Apply table sort",
      description: "Apply sort fields to a structured table.",
      inputSchema: tableSortSchema(),
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) =>
      jsonResult(await runtime.sortTable({ ...tableSelector(args), fields: args.fields, matchCase: args.matchCase, method: args.method } as TableSortRequest))
  );

  registerMcpTool(
    mcp,
    "excel.sort.clear",
    {
      title: "Clear table sort",
      description: "Clear sort state on a structured table.",
      inputSchema: tableSelectorSchema(),
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.clearTableSort(tableSelector(args)))
  );

  registerMcpTool(
    mcp,
    "excel.sort.preserve_from_template",
    {
      title: "Preserve sort from template",
      description: "Apply provided template sort fields to a table.",
      inputSchema: tableSortSchema(),
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) =>
      jsonResult(await runtime.sortTable({ ...tableSelector(args), fields: args.fields, matchCase: args.matchCase, method: args.method } as TableSortRequest))
  );
}

function registerPivotTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.pivot.list",
    {
      title: "List PivotTables",
      description: "List PivotTables in the active workbook.",
      inputSchema: { workbookId: z.string() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ workbookId }: { workbookId: string }) => jsonResult(await runtime.listPivotTables(workbookId as WorkbookId))
  );

  registerMcpTool(
    mcp,
    "excel.pivot.get_info",
    {
      title: "Get PivotTable info",
      description: "Return PivotTable metadata and source details where Office.js exposes them.",
      inputSchema: pivotSelectorSchema(),
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.getPivotTableInfo(pivotSelector(args)))
  );

  registerMcpTool(
    mcp,
    "excel.pivot.create",
    {
      title: "Create PivotTable",
      description: "Create a PivotTable from a source range or structured table at a destination cell.",
      inputSchema: {
        workbookId: z.string(),
        pivotTableName: z.string(),
        sourceSheetName: z.string().optional(),
        sourceAddress: z.string().optional(),
        sourceTableName: z.string().optional(),
        destinationSheetName: z.string(),
        destinationAddress: z.string(),
        rowFields: z.array(z.string()).optional(),
        columnFields: z.array(z.string()).optional(),
        filterFields: z.array(z.string()).optional(),
        dataFields: z
          .array(
            z.object({
              sourceFieldName: z.string(),
              name: z.string().optional(),
              summarizeBy: z.string().optional(),
              numberFormat: z.string().optional()
            })
          )
          .optional(),
        layout: z.record(z.string(), z.unknown()).optional(),
        refresh: z.boolean().optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.createPivotTable(pivotCreateRequest(args)))
  );

  registerMcpTool(
    mcp,
    "excel.pivot.refresh",
    {
      title: "Refresh PivotTable",
      description: "Refresh one PivotTable.",
      inputSchema: pivotSelectorSchema(),
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.refreshPivotTable(pivotSelector(args)))
  );

  registerMcpTool(
    mcp,
    "excel.pivot.refresh_all",
    {
      title: "Refresh all PivotTables",
      description: "Refresh all PivotTables in the workbook.",
      inputSchema: { workbookId: z.string() },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async ({ workbookId }: { workbookId: string }) => jsonResult(await runtime.refreshAllPivotTables(workbookId as WorkbookId))
  );

  registerMcpTool(
    mcp,
    "excel.pivot.update_source",
    {
      title: "Update PivotTable source",
      description: "Return current support status for PivotTable source reassignment.",
      inputSchema: {
        workbookId: z.string(),
        pivotTableName: z.string(),
        sourceSheetName: z.string().optional(),
        sourceAddress: z.string().optional(),
        sourceTableName: z.string().optional(),
        destinationSheetName: z.string(),
        destinationAddress: z.string()
      },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => jsonResult(runtime.updatePivotSource(pivotCreateRequest(args)))
  );

  registerMcpTool(
    mcp,
    "excel.pivot.copy_from_template",
    {
      title: "Copy PivotTable from template",
      description: "Replay deterministic PivotTable metadata and layout from a template PivotTable through the transaction-backed add-in path.",
      inputSchema: {
        ...pivotSelectorSchema(),
        templatePivotTableName: z.string(),
        templateId: z.string().optional(),
        dimensions: z.array(z.enum(["metadata", "layout", "fields", "dataFields", "numberFormats", "filters", "refresh"])).optional(),
        strict: z.boolean().optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => {
      const request: PivotCopyFromTemplateRequest = {
        ...pivotSelector(args),
        templatePivotTableName: args.templatePivotTableName
      };
      if (args.templateId !== undefined) {
        request.templateId = args.templateId as TemplateId;
      }
      if (args.dimensions !== undefined) {
        request.dimensions = args.dimensions;
      }
      if (args.strict !== undefined) {
        request.strict = args.strict;
      }
      return jsonResult(await runtime.copyPivotFromTemplate(request));
    }
  );

  registerMcpTool(
    mcp,
    "excel.pivot.delete",
    {
      title: "Delete PivotTable",
      description: "Delete a PivotTable through a transaction-backed backup path.",
      inputSchema: pivotSelectorSchema(),
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.deletePivotTable(pivotSelector(args)))
  );

  registerMcpTool(
    mcp,
    "excel.pivot.validate_source",
    {
      title: "Validate PivotTable source",
      description: "Validate PivotTable source metadata and optional expected source/layout fields.",
      inputSchema: {
        ...pivotSelectorSchema(),
        expectedFields: z.array(z.string()).optional(),
        expectedRowFields: z.array(z.string()).optional(),
        expectedColumnFields: z.array(z.string()).optional(),
        expectedFilterFields: z.array(z.string()).optional(),
        expectedDataFields: z.array(z.string()).optional(),
        expectedDataFieldSettings: z
          .array(
            z.object({
              sourceFieldName: z.string().optional(),
              name: z.string().optional(),
              summarizeBy: z.string().optional(),
              numberFormat: z.string().optional()
            })
          )
          .optional(),
        expectedLayout: z.record(z.string(), z.unknown()).optional()
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.validatePivotSource(pivotValidateSourceRequest(args)))
  );

  registerMcpTool(
    mcp,
    "excel.pivot.get_capability_matrix",
    {
      title: "Get PivotTable capability matrix",
      description: "Return deterministic PivotTable support status for the active Excel host.",
      inputSchema: { workbookId: z.string().optional() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ workbookId }: { workbookId?: string }) => jsonResult(runtime.getPivotCapabilityMatrix(workbookId as WorkbookId | undefined))
  );

  registerMcpTool(
    mcp,
    "excel.pivot.get_fingerprint",
    {
      title: "Get PivotTable fingerprint",
      description: "Capture a deterministic PivotTable fingerprint from metadata Office.js exposes.",
      inputSchema: pivotSelectorSchema(),
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.getPivotFingerprint(pivotSelector(args)))
  );

  registerMcpTool(
    mcp,
    "excel.pivot.compare_fingerprint",
    {
      title: "Compare PivotTable fingerprint",
      description: "Compare two PivotTable fingerprints and return deterministic differences.",
      inputSchema: { ...pivotSelectorSchema(), targetPivotTableName: z.string() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => {
      const request: PivotCompareFingerprintRequest = {
        ...pivotSelector(args),
        targetPivotTableName: args.targetPivotTableName
      };
      return jsonResult(await runtime.comparePivotFingerprint(request));
    }
  );

  registerMcpTool(
    mcp,
    "excel.pivot.diff",
    {
      title: "Diff PivotTables",
      description: "Return a PivotTable diff based on captured fingerprint dimensions.",
      inputSchema: { ...pivotSelectorSchema(), targetPivotTableName: z.string() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => {
      const request: PivotCompareFingerprintRequest = {
        ...pivotSelector(args),
        targetPivotTableName: args.targetPivotTableName
      };
      return jsonResult(await runtime.diffPivotTables(request));
    }
  );

  registerMcpTool(
    mcp,
    "excel.pivot.repair_from_template",
    {
      title: "Repair PivotTable from template",
      description: "Repair a target PivotTable by replaying deterministic metadata from a template PivotTable, then diffing the result.",
      inputSchema: {
        ...pivotSelectorSchema(),
        templatePivotTableName: z.string(),
        templateId: z.string().optional(),
        dimensions: z.array(z.enum(["metadata", "layout", "fields", "dataFields", "numberFormats", "filters", "refresh"])).optional(),
        strict: z.boolean().optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => {
      const request: PivotRepairFromTemplateRequest = {
        ...pivotSelector(args),
        templatePivotTableName: args.templatePivotTableName
      };
      if (args.templateId !== undefined) request.templateId = args.templateId as TemplateId;
      if (args.dimensions !== undefined) request.dimensions = args.dimensions;
      if (args.strict !== undefined) request.strict = args.strict;
      return jsonResult(await runtime.repairPivotFromTemplate(request));
    }
  );

  registerMcpTool(
    mcp,
    "excel.pivot.rebuild_with_source",
    {
      title: "Rebuild PivotTable with source",
      description: "Create a new PivotTable from the desired source and optionally replay a template PivotTable.",
      inputSchema: {
        workbookId: z.string(),
        pivotTableName: z.string(),
        sourceSheetName: z.string().optional(),
        sourceAddress: z.string().optional(),
        sourceTableName: z.string().optional(),
        destinationSheetName: z.string(),
        destinationAddress: z.string(),
        rowFields: z.array(z.string()).optional(),
        columnFields: z.array(z.string()).optional(),
        filterFields: z.array(z.string()).optional(),
        dataFields: z
          .array(
            z.object({
              sourceFieldName: z.string(),
              name: z.string().optional(),
              summarizeBy: z.string().optional(),
              numberFormat: z.string().optional()
            })
          )
          .optional(),
        layout: z.record(z.string(), z.unknown()).optional(),
        refresh: z.boolean().optional(),
        templatePivotTableName: z.string().optional(),
        replaceExisting: z.boolean().optional(),
        strict: z.boolean().optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => {
      const request = pivotCreateRequest(args) as PivotRebuildWithSourceRequest;
      if (args.templatePivotTableName !== undefined) request.templatePivotTableName = args.templatePivotTableName;
      if (args.replaceExisting !== undefined) request.replaceExisting = args.replaceExisting;
      if (args.strict !== undefined) request.strict = args.strict;
      return jsonResult(await runtime.rebuildPivotWithSource(request));
    }
  );
}

function registerChartTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.chart.list",
    {
      title: "List charts",
      description: "List charts across worksheets.",
      inputSchema: { workbookId: z.string() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ workbookId }: { workbookId: string }) => jsonResult(await runtime.listCharts(workbookId as WorkbookId))
  );

  registerMcpTool(
    mcp,
    "excel.chart.get_info",
    {
      title: "Get chart info",
      description: "Return chart metadata.",
      inputSchema: chartSelectorSchema(),
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.getChartInfo(chartSelector(args)))
  );

  registerMcpTool(
    mcp,
    "excel.chart.create",
    {
      title: "Create chart",
      description: "Create an Excel chart from a source range.",
      inputSchema: chartCreateSchema(),
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.createChart(chartCreateRequest(args)))
  );

  registerMcpTool(
    mcp,
    "excel.chart.update_data_source",
    {
      title: "Update chart data source",
      description: "Reset a chart data source to a new range.",
      inputSchema: { ...chartSelectorSchema(), sourceAddress: z.string(), seriesBy: z.enum(["Auto", "Columns", "Rows"]).optional() },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.updateChartDataSource(chartUpdateDataSourceRequest(args)))
  );

  registerMcpTool(
    mcp,
    "excel.chart.copy_from_template",
    {
      title: "Copy chart from template",
      description: "Copy deterministic chart metadata from a template chart to a target chart.",
      inputSchema: { ...chartSelectorSchema(), templateSheetName: z.string(), templateChartName: z.string() },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) =>
      jsonResult(
        await runtime.copyChartFromTemplate({
          ...chartSelector(args),
          templateSheetName: args.templateSheetName,
          templateChartName: args.templateChartName
        })
      )
  );

  registerMcpTool(
    mcp,
    "excel.chart.refresh",
    {
      title: "Refresh chart",
      description: "Return current chart metadata; charts update from their source data through Excel.",
      inputSchema: chartSelectorSchema(),
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.refreshChart(chartSelector(args)))
  );

  registerMcpTool(
    mcp,
    "excel.chart.delete",
    {
      title: "Delete chart",
      description: "Delete a chart from a worksheet.",
      inputSchema: chartSelectorSchema(),
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.deleteChart(chartSelector(args)))
  );

  registerMcpTool(
    mcp,
    "excel.chart.validate_against_template",
    {
      title: "Validate chart against template",
      description: "Validate target and template chart metadata availability.",
      inputSchema: { ...chartSelectorSchema(), templateSheetName: z.string().optional(), templateChartName: z.string().optional() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.validateChartAgainstTemplate(chartTemplateValidationRequest(args)))
  );
}

function registerNamesTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.names.list",
    {
      title: "List named items",
      description: "List workbook and worksheet scoped Excel named items.",
      inputSchema: { workbookId: z.string() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ workbookId }: { workbookId: string }) => jsonResult(await runtime.listNames(workbookId as WorkbookId))
  );

  registerMcpTool(
    mcp,
    "excel.names.get",
    {
      title: "Get named item",
      description: "Get one workbook or worksheet scoped Excel named item.",
      inputSchema: nameSelectorSchema(),
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.getName(nameSelector(args)))
  );

  registerMcpTool(
    mcp,
    "excel.names.create",
    {
      title: "Create named item",
      description: "Create a workbook or worksheet scoped Excel name for a formula or range reference.",
      inputSchema: nameMutationSchema(),
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.createName(nameCreateRequest(args)))
  );

  registerMcpTool(
    mcp,
    "excel.names.update",
    {
      title: "Update named item",
      description: "Update a named item's formula/reference, comment, or visibility.",
      inputSchema: nameMutationSchema(),
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.updateName(nameUpdateRequest(args)))
  );

  registerMcpTool(
    mcp,
    "excel.names.delete",
    {
      title: "Delete named item",
      description: "Delete a workbook or worksheet scoped named item.",
      inputSchema: nameSelectorSchema(),
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.deleteName(nameSelector(args)))
  );
}

function registerRegionTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.region.detect",
    {
      title: "Detect workbook regions",
      description: "Return registered regions plus named-range and used-range region candidates.",
      inputSchema: { workbookId: z.string() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ workbookId }: { workbookId: string }) => jsonResult(await runtime.detectRegions(workbookId as WorkbookId))
  );

  registerMcpTool(
    mcp,
    "excel.region.register",
    {
      title: "Register workbook region",
      description: "Register a reusable workbook region, optionally creating a matching Excel named range.",
      inputSchema: {
        workbookId: z.string(),
        name: z.string(),
        sheetName: z.string(),
        address: z.string(),
        kind: z.enum(["data", "header", "formula", "output", "template", "table", "named-range", "other"]).optional(),
        description: z.string().optional(),
        templateId: z.string().optional(),
        createNamedRange: z.boolean().optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.registerRegion(regionRegisterRequest(args)))
  );

  registerMcpTool(
    mcp,
    "excel.region.list",
    {
      title: "List workbook regions",
      description: "List registered workbook regions.",
      inputSchema: { workbookId: z.string() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ workbookId }: { workbookId: string }) => jsonResult(runtime.listRegions(workbookId as WorkbookId))
  );

  registerMcpTool(
    mcp,
    "excel.region.get",
    {
      title: "Get workbook region",
      description: "Get a registered region or resolve an Excel named range as a region.",
      inputSchema: regionSelectorSchema(),
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.getRegion(regionSelector(args)))
  );

  registerMcpTool(
    mcp,
    "excel.region.clear_values",
    {
      title: "Clear region values",
      description: "Clear values in a registered region while preserving formats.",
      inputSchema: regionSelectorSchema(),
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.clearRegionValues(regionSelector(args)))
  );

  registerMcpTool(
    mcp,
    "excel.region.write_values",
    {
      title: "Write region values",
      description: "Write values to a registered region while preserving formats.",
      inputSchema: { ...regionSelectorSchema(), values: z.array(z.array(z.any())) },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.writeRegionValues({ ...regionSelector(args), values: args.values }))
  );

  registerMcpTool(
    mcp,
    "excel.region.fill",
    {
      title: "Fill workbook region",
      description: "Optionally clear then write values to a registered region while preserving formats.",
      inputSchema: { ...regionSelectorSchema(), values: z.array(z.array(z.any())), clearFirst: z.boolean().optional() },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => {
      const request: RegionSelector & { values: unknown[][]; clearFirst?: boolean } = { ...regionSelector(args), values: args.values };
      if (args.clearFirst !== undefined) {
        request.clearFirst = args.clearFirst;
      }
      return jsonResult(await runtime.fillRegion(request));
    }
  );
}

function registerTaskTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.task.create",
    {
      title: "Create Excel task",
      description: "Create a multi-agent workbook task with optional scope and assigned agent.",
      inputSchema: {
        workbookId: z.string(),
        goal: z.string(),
        role: z.string().optional(),
        priority: z.enum(["low", "normal", "high"]).optional(),
        assignedAgentId: z.string().optional(),
        allowedScopes: z.array(z.any()).optional(),
        dependencies: z.array(z.string()).optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) =>
      jsonResult(
        runtime.createTask({
          workbookId: args.workbookId as WorkbookId,
          goal: args.goal,
          role: args.role,
          priority: args.priority,
          assignedAgentId: args.assignedAgentId as AgentId | undefined,
          allowedScopes: args.allowedScopes,
          dependencies: args.dependencies as TaskId[] | undefined
        })
      )
  );

  registerMcpTool(
    mcp,
    "excel.task.claim",
    {
      title: "Claim Excel task",
      description: "Assign a task to an agent.",
      inputSchema: { taskId: z.string(), agentId: z.string() },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async ({ taskId, agentId }: { taskId: string; agentId: string }) => jsonResult(runtime.claimTask(taskId as TaskId, agentId as AgentId))
  );

  registerMcpTool(
    mcp,
    "excel.task.update",
    {
      title: "Update Excel task",
      description: "Update task metadata, status, scope, or assignment.",
      inputSchema: {
        taskId: z.string(),
        goal: z.string().optional(),
        role: z.string().optional(),
        priority: z.enum(["low", "normal", "high"]).optional(),
        status: z.enum(["open", "claimed", "planning", "queued", "applying", "blocked", "completed", "failed", "cancelled"]).optional(),
        progress: z.number().min(0).max(100).optional(),
        currentStep: z.string().optional(),
        blockers: z.array(z.any()).optional(),
        assignedAgentId: z.string().optional(),
        allowedScopes: z.array(z.any()).optional(),
        dependencies: z.array(z.string()).optional(),
        errorMessage: z.string().optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => {
      const { taskId, ...patch } = args;
      return jsonResult(runtime.updateTask(taskId as TaskId, patch));
    }
  );

  registerMcpTool(
    mcp,
    "excel.task.set_progress",
    {
      title: "Set Excel task progress",
      description: "Update task progress and the current step shown in collaboration status.",
      inputSchema: {
        taskId: z.string(),
        progress: z.number().min(0).max(100),
        currentStep: z.string().optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async ({ taskId, progress, currentStep }: { taskId: string; progress: number; currentStep?: string }) =>
      jsonResult(runtime.setTaskProgress(taskId as TaskId, progress, currentStep))
  );

  registerMcpTool(
    mcp,
    "excel.task.add_blocker",
    {
      title: "Add Excel task blocker",
      description: "Add an open blocker, warning, or informational note to a task.",
      inputSchema: {
        taskId: z.string(),
        severity: z.enum(["info", "warning", "blocked"]),
        message: z.string(),
        scope: z.any().optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) =>
      jsonResult(
        runtime.addTaskBlocker(args.taskId as TaskId, {
          severity: args.severity,
          message: args.message,
          scope: args.scope
        })
      )
  );

  registerMcpTool(
    mcp,
    "excel.task.resolve_blocker",
    {
      title: "Resolve Excel task blocker",
      description: "Mark a task blocker as resolved.",
      inputSchema: {
        taskId: z.string(),
        blockerId: z.string()
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async ({ taskId, blockerId }: { taskId: string; blockerId: string }) => jsonResult(runtime.resolveTaskBlocker(taskId as TaskId, blockerId))
  );

  registerMcpTool(
    mcp,
    "excel.task.evaluate_schedule",
    {
      title: "Evaluate Excel task schedule",
      description: "Evaluate task readiness against dependencies, blockers, and active locks.",
      inputSchema: {
        workbookId: z.string().optional(),
        apply: z.boolean().optional(),
        lockMode: z.enum(["read", "write_values", "write_formulas", "write_styles", "format_layout", "table", "chart", "pivot", "structure", "workbook"]).optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) =>
      jsonResult(
        runtime.evaluateTaskSchedule({
          workbookId: args.workbookId as WorkbookId | undefined,
          apply: args.apply,
          lockMode: args.lockMode
        })
      )
  );

  registerMcpTool(
    mcp,
    "excel.task.resume_ready",
    {
      title: "Resume ready Excel tasks",
      description: "Apply scheduler decisions so blocked tasks with cleared dependencies can resume.",
      inputSchema: {
        workbookId: z.string().optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async ({ workbookId }: { workbookId?: string }) => jsonResult(runtime.resumeReadyTasks(workbookId as WorkbookId | undefined))
  );

  for (const [name, status] of [
    ["excel.task.complete", "completed"],
    ["excel.task.fail", "failed"],
    ["excel.task.cancel", "cancelled"]
  ] as const) {
    registerMcpTool(
      mcp,
      name,
      {
        title: name.replace(/^excel\./, ""),
        description: `Mark a task as ${status}.`,
        inputSchema: { taskId: z.string(), errorMessage: z.string().optional() },
        annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
      },
      async ({ taskId, errorMessage }: { taskId: string; errorMessage?: string }) =>
        jsonResult(runtime.updateTask(taskId as TaskId, status === "failed" ? { status, errorMessage } : { status }))
    );
  }

  registerMcpTool(
    mcp,
    "excel.task.list",
    {
      title: "List Excel tasks",
      description: "List multi-agent workbook tasks.",
      inputSchema: { workbookId: z.string().optional() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ workbookId }: { workbookId?: string }) => jsonResult(runtime.listTasks(workbookId as WorkbookId | undefined))
  );

  registerMcpTool(
    mcp,
    "excel.task.get",
    {
      title: "Get Excel task",
      description: "Return a task by ID.",
      inputSchema: { taskId: z.string() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ taskId }: { taskId: string }) => jsonResult(runtime.getTask(taskId as TaskId))
  );
}

function registerCollaborationTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.collab.get_status",
    {
      title: "Get collaboration status",
      description: "Return agents, tasks, locks, transactions, conflicts, and recent collaboration events.",
      inputSchema: { workbookId: z.string().optional() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ workbookId }: { workbookId?: string }) => jsonResult(runtime.getCollaborationStatus(workbookId as WorkbookId | undefined))
  );

  for (const [name, method] of [
    ["excel.collab.list_agents", "listAgents"],
    ["excel.collab.list_tasks", "listTasks"],
    ["excel.collab.list_locks", "listLocks"],
    ["excel.collab.list_transactions", "listTransactions"],
    ["excel.collab.get_conflicts", "listConflicts"]
  ] as const) {
    registerMcpTool(
      mcp,
      name,
      {
        title: name.replace(/^excel\./, ""),
        description: "Return collaboration runtime state.",
        inputSchema: { workbookId: z.string().optional() },
        annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
      },
      async ({ workbookId }: { workbookId?: string }) => jsonResult(runtime[method](workbookId as WorkbookId | undefined))
    );
  }

  registerMcpTool(
    mcp,
    "excel.collab.get_recent_events",
    {
      title: "Get recent collaboration events",
      description: "Return recent collaboration events from the shared runtime.",
      inputSchema: { workbookId: z.string().optional() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ workbookId }: { workbookId?: string }) => jsonResult(runtime.getCollaborationStatus(workbookId as WorkbookId | undefined).events)
  );
}

function registerLockTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.lock.get_policy",
    {
      title: "Get lock lease policy",
      description: "Return runtime lock TTL and manual-lock policy.",
      inputSchema: {},
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async () => jsonResult(runtime.getLockPolicy())
  );

  registerMcpTool(
    mcp,
    "excel.lock.set_policy",
    {
      title: "Set lock lease policy",
      description: "Update runtime lock TTL and manual-lock policy.",
      inputSchema: {
        defaultTtlMs: z.number().int().positive().optional(),
        transactionTtlMs: z.number().int().positive().optional(),
        maxTtlMs: z.number().int().positive().optional(),
        allowManualLocks: z.boolean().optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(runtime.setLockPolicy(args))
  );

  registerMcpTool(
    mcp,
    "excel.lock.acquire",
    {
      title: "Acquire Excel lock",
      description: "Acquire explicit workbook/sheet/range/object locks for multi-agent planning or guarded work.",
      inputSchema: {
        workbookId: z.string(),
        scopes: z.array(z.any()),
        mode: z.enum(["read", "write_values", "write_formulas", "write_styles", "format_layout", "table", "chart", "pivot", "structure", "workbook"]),
        reason: z.string(),
        ownerAgentId: z.string().optional(),
        taskId: z.string().optional(),
        ttlMs: z.number().int().positive().optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) =>
      jsonResult(
        runtime.acquireLocks({
          workbookId: args.workbookId as WorkbookId,
          scopes: args.scopes as WorkbookScope[],
          mode: args.mode as LockMode,
          reason: args.reason,
          ownerAgentId: args.ownerAgentId as AgentId | undefined,
          taskId: args.taskId as TaskId | undefined,
          ttlMs: args.ttlMs
        })
      )
  );

  registerMcpTool(
    mcp,
    "excel.lock.renew",
    {
      title: "Renew Excel locks",
      description: "Extend active lock leases up to the runtime max TTL.",
      inputSchema: {
        lockIds: z.array(z.string()),
        ttlMs: z.number().int().positive().optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async ({ lockIds, ttlMs }: { lockIds: string[]; ttlMs?: number }) => jsonResult(runtime.renewLocks(lockIds as LockId[], ttlMs))
  );

  registerMcpTool(
    mcp,
    "excel.lock.release",
    {
      title: "Release Excel locks",
      description: "Release active lock leases.",
      inputSchema: {
        lockIds: z.array(z.string())
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async ({ lockIds }: { lockIds: string[] }) => jsonResult(runtime.releaseLocks(lockIds as LockId[]))
  );
}

function registerConflictTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.conflict.get_guidance",
    {
      title: "Get conflict guidance",
      description: "Return actionable conflict-resolution guidance for recent runtime conflicts.",
      inputSchema: { workbookId: z.string().optional() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ workbookId }: { workbookId?: string }) => jsonResult(runtime.getConflictGuidance(workbookId as WorkbookId | undefined))
  );

  registerMcpTool(
    mcp,
    "excel.conflict.explain",
    {
      title: "Explain conflict",
      description: "Return actionable resolution guidance for a supplied conflict record.",
      inputSchema: {
        conflict: z.any()
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ conflict }: { conflict: ConflictRecord }) => jsonResult(runtime.explainConflict(conflict))
  );

  registerMcpTool(
    mcp,
    "excel.conflict.get_telemetry",
    {
      title: "Get conflict telemetry",
      description: "Summarize repeated contention, hot scopes, tasks, agents, and wait outcomes.",
      inputSchema: {
        workbookId: z.string().optional(),
        windowSize: z.number().int().positive().optional()
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ workbookId, windowSize }: { workbookId?: string; windowSize?: number }) =>
      jsonResult(runtime.getConflictTelemetry(workbookId as WorkbookId | undefined, windowSize))
  );

  registerMcpTool(
    mcp,
    "excel.conflict.clear_telemetry",
    {
      title: "Clear conflict telemetry",
      description: "Clear conflict telemetry for one workbook or the whole runtime.",
      inputSchema: {
        workbookId: z.string().optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async ({ workbookId }: { workbookId?: string }) => jsonResult(runtime.clearConflictTelemetry(workbookId as WorkbookId | undefined))
  );
}

function registerTransactionTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.transaction.list",
    {
      title: "List Excel transactions",
      description: "List serialized workbook transactions.",
      inputSchema: { workbookId: z.string().optional() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ workbookId }: { workbookId?: string }) => jsonResult(runtime.listTransactions(workbookId as WorkbookId | undefined))
  );

  registerMcpTool(
    mcp,
    "excel.transaction.get",
    {
      title: "Get Excel transaction",
      description: "Return one serialized workbook transaction.",
      inputSchema: { transactionId: z.string() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ transactionId }: { transactionId: string }) => jsonResult(runtime.getTransaction(transactionId as TransactionId))
  );

  registerMcpTool(
    mcp,
    "excel.transaction.preview_rollback",
    {
      title: "Preview transaction rollback",
      description: "Check whether a transaction can be rolled back without overwriting later work.",
      inputSchema: { transactionId: z.string() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ transactionId }: { transactionId: string }) => jsonResult(runtime.previewTransactionRollback(transactionId as TransactionId))
  );

  registerMcpTool(
    mcp,
    "excel.transaction.rollback",
    {
      title: "Rollback transaction",
      description: "Rollback a transaction only when rollback preview has no later-overlap conflicts.",
      inputSchema: { transactionId: z.string(), confirmationToken: z.string().optional() },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async ({ transactionId, confirmationToken }: { transactionId: string; confirmationToken?: string }) =>
      jsonResult(await runtime.rollbackTransaction(transactionId as TransactionId, confirmationToken))
  );

  registerMcpTool(
    mcp,
    "excel.transaction.preview_rollback_chain",
    {
      title: "Preview transaction rollback chain",
      description: "Find later dependent transactions that must be rolled back newest-first with the target transaction.",
      inputSchema: { transactionId: z.string() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ transactionId }: { transactionId: string }) => jsonResult(runtime.previewTransactionRollbackChain(transactionId as TransactionId))
  );

  registerMcpTool(
    mcp,
    "excel.transaction.rollback_chain",
    {
      title: "Rollback transaction chain",
      description: "Rollback a confirmed related transaction chain newest-first.",
      inputSchema: { transactionId: z.string(), confirmationToken: z.string().optional() },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async ({ transactionId, confirmationToken }: { transactionId: string; confirmationToken?: string }) =>
      jsonResult(await runtime.rollbackTransactionChain(transactionId as TransactionId, confirmationToken))
  );
}

function registerPermissionsTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.permissions.get",
    {
      title: "Get permissions",
      description: "Return current Open Workbook permission policy, scope, and locked regions.",
      inputSchema: {},
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async () => jsonResult(runtime.getPermissions())
  );

  registerMcpTool(
    mcp,
    "excel.permissions.set",
    {
      title: "Set permissions",
      description: "Update Open Workbook permission policy.",
      inputSchema: permissionSetSchema(),
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(runtime.setPermissions(permissionUpdate(args)))
  );

  registerMcpTool(
    mcp,
    "excel.permissions.require_confirmation",
    {
      title: "Require confirmation",
      description: "Set destructive levels that require a confirmation token before apply.",
      inputSchema: { levels: z.array(z.enum(["none", "values", "format", "structure", "workbook"])) },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async ({ levels }: { levels: PermissionState["requireConfirmationFor"] }) => jsonResult(runtime.requireConfirmation(levels))
  );

  registerMcpTool(
    mcp,
    "excel.permissions.set_scope",
    {
      title: "Set permission scope",
      description: "Restrict mutations to a workbook, sheet names, or registered region names.",
      inputSchema: {
        workbookId: z.string().optional(),
        sheetNames: z.array(z.string()).optional(),
        regionNames: z.array(z.string()).optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(runtime.setPermissionScope(permissionScope(args)))
  );

  registerMcpTool(
    mcp,
    "excel.permissions.allow_destructive_actions",
    {
      title: "Allow destructive actions",
      description: "Allow or block structure/workbook destructive actions.",
      inputSchema: { allow: z.boolean() },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async ({ allow }: { allow: boolean }) => jsonResult(runtime.allowDestructiveActions(allow))
  );

  registerMcpTool(
    mcp,
    "excel.permissions.allow_macro_execution",
    {
      title: "Allow macro execution",
      description: "Record macro execution permission. Macro execution is not implemented by the Office.js runtime.",
      inputSchema: { allow: z.boolean() },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async ({ allow }: { allow: boolean }) => jsonResult(runtime.allowMacroExecution(allow))
  );

  registerMcpTool(
    mcp,
    "excel.permissions.lock_regions",
    {
      title: "Lock regions",
      description: "Block future writes that overlap registered regions.",
      inputSchema: {
        workbookId: z.string(),
        regions: z.array(z.object({ regionName: z.string(), reason: z.string().optional() }))
      },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.lockRegions({ workbookId: args.workbookId as WorkbookId, regions: args.regions }))
  );

  registerMcpTool(
    mcp,
    "excel.permissions.unlock_regions",
    {
      title: "Unlock regions",
      description: "Unlock all or selected locked regions for a workbook.",
      inputSchema: { workbookId: z.string(), regionNames: z.array(z.string()).optional() },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => {
      const request: { workbookId: WorkbookId; regionNames?: string[] } = { workbookId: args.workbookId as WorkbookId };
      if (args.regionNames !== undefined) {
        request.regionNames = args.regionNames;
      }
      return jsonResult(runtime.unlockRegions(request));
    }
  );
}

function registerCleanTools(mcp: McpServer): void {
  const rangeSchema = cleanRangeSchema();

  registerMcpTool(
    mcp,
    "excel.clean.detect_header_row",
    {
      title: "Detect header row",
      description: "Score likely header rows in a range without mutating the workbook.",
      inputSchema: { ...rangeSchema, maxRows: z.number().int().min(1).max(100).optional() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.cleanDetectHeaderRow(cleanRangeArgs(args, { maxRows: args.maxRows })))
  );

  registerMcpTool(
    mcp,
    "excel.clean.normalize_headers",
    {
      title: "Normalize headers",
      description: "Normalize a header row to lowercase snake_case and deduplicate names.",
      inputSchema: { ...rangeSchema, headerRowIndex: z.number().int().min(0).optional() },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.cleanNormalizeHeaders(cleanRangeArgs(args, { headerRowIndex: args.headerRowIndex })))
  );

  registerMcpTool(
    mcp,
    "excel.clean.trim_whitespace",
    {
      title: "Trim whitespace",
      description: "Trim and collapse whitespace in string cells.",
      inputSchema: rangeSchema,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.cleanTrimWhitespace(cleanRangeArgs(args)))
  );

  registerMcpTool(
    mcp,
    "excel.clean.remove_duplicates",
    {
      title: "Remove duplicate rows",
      description: "Compact duplicate rows in-place and blank removed rows to preserve range shape.",
      inputSchema: { ...rangeSchema, hasHeader: z.boolean().optional(), keyColumns: z.array(z.number().int().min(0)).optional() },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.cleanRemoveDuplicates(cleanRangeArgs(args, { hasHeader: args.hasHeader, keyColumns: args.keyColumns })))
  );

  for (const [name, method] of [
    ["excel.clean.parse_dates", "cleanParseDates"],
    ["excel.clean.parse_numbers", "cleanParseNumbers"],
    ["excel.clean.standardize_currency", "cleanStandardizeCurrency"]
  ] as const) {
    registerMcpTool(
      mcp,
      name,
      {
        title: name.replace(/^excel\./, "").replace(/\./g, " "),
        description: "Parse and standardize cell values in a range.",
        inputSchema: rangeSchema,
        annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
      },
      async (args: any) => jsonResult(await runtime[method](cleanRangeArgs(args)))
    );
  }

  registerMcpTool(
    mcp,
    "excel.clean.fill_missing_values",
    {
      title: "Fill missing values",
      description: "Fill blank cells using a fixed value, zero, previous value, or next value.",
      inputSchema: { ...rangeSchema, strategy: z.enum(["value", "zero", "previous", "next"]).optional(), value: z.any().optional() },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.cleanFillMissingValues(cleanRangeArgs(args, { strategy: args.strategy, value: args.value })))
  );

  registerMcpTool(
    mcp,
    "excel.clean.split_column",
    {
      title: "Split column",
      description: "Split one source column by delimiter and write the output to a target range.",
      inputSchema: { ...rangeSchema, columnIndex: z.number().int().min(0), delimiter: z.string().optional(), targetAddress: z.string() },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) =>
      jsonResult(await runtime.cleanSplitColumn(cleanRangeArgs(args, { columnIndex: args.columnIndex, delimiter: args.delimiter, targetAddress: args.targetAddress })))
  );

  registerMcpTool(
    mcp,
    "excel.clean.merge_columns",
    {
      title: "Merge columns",
      description: "Merge selected source columns and write the output to a target range.",
      inputSchema: { ...rangeSchema, columnIndexes: z.array(z.number().int().min(0)), separator: z.string().optional(), targetAddress: z.string() },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) =>
      jsonResult(await runtime.cleanMergeColumns(cleanRangeArgs(args, { columnIndexes: args.columnIndexes, separator: args.separator, targetAddress: args.targetAddress })))
  );

  registerMcpTool(
    mcp,
    "excel.clean.detect_outliers",
    {
      title: "Detect outliers",
      description: "Detect numeric outliers using z-score threshold without mutating the workbook.",
      inputSchema: { ...rangeSchema, columnIndex: z.number().int().min(0).optional(), threshold: z.number().positive().optional() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.cleanDetectOutliers(cleanRangeArgs(args, { columnIndex: args.columnIndex, threshold: args.threshold })))
  );

  registerMcpTool(
    mcp,
    "excel.clean.fuzzy_match",
    {
      title: "Fuzzy match",
      description: "Compare cell text to lookup values and return similarity matches without mutating the workbook.",
      inputSchema: { ...rangeSchema, lookupValues: z.array(z.string()), threshold: z.number().min(0).max(1).optional() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.cleanFuzzyMatch(cleanRangeArgs(args, { lookupValues: args.lookupValues, threshold: args.threshold })))
  );
}

function registerValidateTools(mcp: McpServer): void {
  const validationRangeSchema = {
    workbookId: z.string(),
    sheetName: z.string().optional(),
    address: z.string().optional()
  };

  registerMcpTool(
    mcp,
    "excel.validate.workbook",
    {
      title: "Validate workbook",
      description: "Run workbook-level health checks, including map availability and formula-error scanning over used ranges.",
      inputSchema: { workbookId: z.string() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ workbookId }: { workbookId: string }) => jsonResult(await runtime.validateWorkbook({ workbookId: workbookId as WorkbookId }))
  );

  registerMcpTool(
    mcp,
    "excel.validate.sheet",
    {
      title: "Validate sheet",
      description: "Validate one worksheet's used range and formula-error state.",
      inputSchema: { workbookId: z.string(), sheetName: z.string() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ workbookId, sheetName }: { workbookId: string; sheetName: string }) =>
      jsonResult(await runtime.validateSheet({ workbookId: workbookId as WorkbookId, sheetName }))
  );

  registerMcpTool(
    mcp,
    "excel.validate.template_consistency",
    {
      title: "Validate template consistency",
      description: "Compare a target sheet against a registered template fingerprint.",
      inputSchema: { workbookId: z.string(), templateId: z.string(), targetSheetName: z.string() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) =>
      jsonResult(
        await runtime.validateTemplateConsistency({
          workbookId: args.workbookId as WorkbookId,
          templateId: args.templateId as TemplateId,
          targetSheetName: args.targetSheetName
        })
      )
  );

  registerMcpTool(
    mcp,
    "excel.validate.formulas",
    {
      title: "Validate formulas",
      description: "Find formula errors in a workbook, sheet, or explicit range.",
      inputSchema: validationRangeSchema,
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.validateFormulas(validationRangeArgs(args)))
  );

  registerMcpTool(
    mcp,
    "excel.validate.styles",
    {
      title: "Validate styles",
      description: "Capture style fingerprint data or compare styles against a registered template.",
      inputSchema: {
        workbookId: z.string(),
        sheetName: z.string().optional(),
        templateId: z.string().optional(),
        targetSheetName: z.string().optional()
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => {
      const request: { workbookId: WorkbookId; templateId?: TemplateId; targetSheetName?: string; sheetName?: string } = {
        workbookId: args.workbookId as WorkbookId
      };
      if (args.sheetName !== undefined) {
        request.sheetName = args.sheetName;
      }
      if (args.templateId !== undefined) {
        request.templateId = args.templateId as TemplateId;
      }
      if (args.targetSheetName !== undefined) {
        request.targetSheetName = args.targetSheetName;
      }
      return jsonResult(await runtime.validateStyles(request));
    }
  );

  registerMcpTool(
    mcp,
    "excel.validate.tables",
    {
      title: "Validate tables",
      description: "Inspect structured table metadata, optionally against a registered template.",
      inputSchema: { workbookId: z.string(), tableName: z.string().optional(), templateId: z.string().optional() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => {
      const request: { workbookId: WorkbookId; tableName?: string; templateId?: TemplateId } = {
        workbookId: args.workbookId as WorkbookId
      };
      if (args.tableName !== undefined) {
        request.tableName = args.tableName;
      }
      if (args.templateId !== undefined) {
        request.templateId = args.templateId as TemplateId;
      }
      return jsonResult(await runtime.validateTables(request));
    }
  );

  registerMcpTool(
    mcp,
    "excel.validate.filters",
    {
      title: "Validate filters",
      description: "Inspect current table filter metadata for one table or the workbook table list.",
      inputSchema: { workbookId: z.string(), tableName: z.string().optional() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => {
      const request: { workbookId: WorkbookId; tableName?: string } = { workbookId: args.workbookId as WorkbookId };
      if (args.tableName !== undefined) {
        request.tableName = args.tableName;
      }
      return jsonResult(await runtime.validateFilters(request));
    }
  );

  registerMcpTool(
    mcp,
    "excel.validate.print_layout",
    {
      title: "Validate print layout",
      description: "Return print-layout validation status and current capability limitations.",
      inputSchema: { workbookId: z.string(), templateId: z.string().optional(), targetSheetName: z.string().optional() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => {
      const request: { workbookId: WorkbookId; templateId?: TemplateId; targetSheetName?: string } = {
        workbookId: args.workbookId as WorkbookId
      };
      if (args.templateId !== undefined) {
        request.templateId = args.templateId as TemplateId;
      }
      if (args.targetSheetName !== undefined) {
        request.targetSheetName = args.targetSheetName;
      }
      return jsonResult(runtime.validatePrintLayout(request));
    }
  );

  registerMcpTool(
    mcp,
    "excel.validate.no_broken_references",
    {
      title: "Validate no broken references",
      description: "Search used ranges for #REF! broken-reference markers.",
      inputSchema: validationRangeSchema,
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.validateNoBrokenReferences(validationRangeArgs(args)))
  );

  registerMcpTool(
    mcp,
    "excel.validate.no_formula_errors",
    {
      title: "Validate no formula errors",
      description: "Assert that formula-error cells are absent from a workbook, sheet, or explicit range.",
      inputSchema: validationRangeSchema,
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => jsonResult(await runtime.validateNoFormulaErrors(validationRangeArgs(args)))
  );

  registerMcpTool(
    mcp,
    "excel.validate.no_unintended_changes",
    {
      title: "Validate no unintended changes",
      description: "Compare snapshots or detect changes since a snapshot.",
      inputSchema: {
        workbookId: z.string(),
        snapshotId: z.string().optional(),
        leftSnapshotId: z.string().optional(),
        rightSnapshotId: z.string().optional()
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async (args: any) => {
      const request: { workbookId: WorkbookId; snapshotId?: SnapshotId; leftSnapshotId?: SnapshotId; rightSnapshotId?: SnapshotId } = {
        workbookId: args.workbookId as WorkbookId
      };
      if (args.snapshotId !== undefined) {
        request.snapshotId = args.snapshotId as SnapshotId;
      }
      if (args.leftSnapshotId !== undefined) {
        request.leftSnapshotId = args.leftSnapshotId as SnapshotId;
      }
      if (args.rightSnapshotId !== undefined) {
        request.rightSnapshotId = args.rightSnapshotId as SnapshotId;
      }
      return jsonResult(await runtime.validateNoUnintendedChanges(request));
    }
  );
}

function registerRepairTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.repair.style_from_template",
    {
      title: "Repair style from template",
      description: "Repair target sheet styles from a registered template with a rollback backup.",
      inputSchema: templateRepairSchema(),
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) =>
      jsonResult(
        await runtime.repairStyleFromTemplate({
          workbookId: args.workbookId as WorkbookId,
          templateId: args.templateId as TemplateId,
          targetSheetName: args.targetSheetName
        })
      )
  );

  registerMcpTool(
    mcp,
    "excel.repair.formulas_from_template",
    {
      title: "Repair formulas from template",
      description: "Repair target sheet formulas from a registered template with a rollback backup.",
      inputSchema: templateRepairSchema(),
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) =>
      jsonResult(
        await runtime.repairFormulasFromTemplate({
          workbookId: args.workbookId as WorkbookId,
          templateId: args.templateId as TemplateId,
          targetSheetName: args.targetSheetName
        })
      )
  );

  registerMcpTool(
    mcp,
    "excel.repair.filters_from_template",
    {
      title: "Repair filters from template",
      description: "Return current filter repair capability status for a registered template.",
      inputSchema: templateRepairSchema(),
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => {
      const request: { workbookId: WorkbookId; templateId?: TemplateId; targetSheetName?: string } = {
        workbookId: args.workbookId as WorkbookId
      };
      if (args.templateId !== undefined) {
        request.templateId = args.templateId as TemplateId;
      }
      if (args.targetSheetName !== undefined) {
        request.targetSheetName = args.targetSheetName;
      }
      return jsonResult(runtime.repairFiltersFromTemplate(request));
    }
  );

  registerMcpTool(
    mcp,
    "excel.repair.table_structure",
    {
      title: "Repair table structure",
      description: "Copy table headers and optional style/totals to a target range with rollback backup.",
      inputSchema: {
        ...tableSelectorSchema(),
        targetSheetName: z.string(),
        targetAddress: z.string(),
        newTableName: z.string().optional(),
        includeStyle: z.boolean().optional(),
        includeTotals: z.boolean().optional(),
        includeFilters: z.boolean().optional()
      },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) =>
      jsonResult(
        await runtime.repairTableStructure({
          ...tableSelector(args),
          targetSheetName: args.targetSheetName,
          targetAddress: args.targetAddress,
          newTableName: args.newTableName,
          includeStyle: args.includeStyle,
          includeTotals: args.includeTotals,
          includeFilters: args.includeFilters
        } as TableCopyStructureRequest)
      )
  );

  for (const [name, repair] of [
    ["excel.repair.print_layout", "repairPrintLayout"],
    ["excel.repair.named_ranges", "repairNamedRanges"],
    ["excel.repair.formula_errors", "repairFormulaErrors"],
    ["excel.repair.merged_cells", "repairMergedCells"]
  ] as const) {
    registerMcpTool(
      mcp,
      name,
      {
        title: name.replace(/^excel\./, "").replace(/\./g, " "),
        description: "Return current repair capability status for this repair category.",
        inputSchema: { workbookId: z.string() },
        annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
      },
      async ({ workbookId }: { workbookId: string }) => jsonResult(runtime[repair]({ workbookId: workbookId as WorkbookId }))
    );
  }
}

function registerSnapshotTools(mcp: McpServer): void {
  for (const name of ["excel.snapshot.create", "excel.snapshot.refresh"] as const) {
    registerMcpTool(
      mcp,
      name,
      {
        title: name.replace(/^excel\./, "").replace(/\./g, " "),
        description: "Create or refresh a workbook snapshot.",
        inputSchema: name.endsWith(".refresh")
          ? { snapshotId: z.string(), reason: z.string().optional() }
          : snapshotInputSchema(),
        annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
      },
      async (args: any) => {
        if (name.endsWith(".refresh")) {
          const existing = await runtime.getSnapshot(args.snapshotId as SnapshotId);
          if (!existing.ok || !("snapshot" in existing)) {
            return jsonResult(existing);
          }
          return jsonResult(
            await runtime.createWorkbookSnapshot({
              workbookId: existing.snapshot.workbookId,
              reason: args.reason ?? `Refresh snapshot ${args.snapshotId}`,
              ranges: existing.snapshot.affectedRanges
            })
          );
        }
        return jsonResult(
          await runtime.createWorkbookSnapshot(snapshotRequest(args.workbookId, args.reason, args.ranges))
        );
      }
    );
  }

  registerMcpTool(
    mcp,
    "excel.snapshot.get",
    {
      title: "Get snapshot",
      description: "Return a stored workbook snapshot.",
      inputSchema: { snapshotId: z.string() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ snapshotId }: { snapshotId: string }) => jsonResult(runtime.getSnapshot(snapshotId as SnapshotId))
  );

  registerMcpTool(
    mcp,
    "excel.snapshot.list",
    {
      title: "List snapshots",
      description: "List stored snapshots for a workbook.",
      inputSchema: { workbookId: z.string() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ workbookId }: { workbookId: string }) => jsonResult(runtime.listSnapshots(workbookId as WorkbookId))
  );

  registerMcpTool(
    mcp,
    "excel.snapshot.compare",
    {
      title: "Compare snapshots",
      description: "Compare two workbook snapshots.",
      inputSchema: {
        leftSnapshotId: z.string(),
        rightSnapshotId: z.string()
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ leftSnapshotId, rightSnapshotId }: { leftSnapshotId: string; rightSnapshotId: string }) =>
      jsonResult(runtime.compareSnapshots(leftSnapshotId as SnapshotId, rightSnapshotId as SnapshotId))
  );

  registerMcpTool(
    mcp,
    "excel.snapshot.invalidate",
    {
      title: "Invalidate snapshot",
      description: "Mark a snapshot as stale without deleting it.",
      inputSchema: { snapshotId: z.string() },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async ({ snapshotId }: { snapshotId: string }) => jsonResult(runtime.invalidateSnapshot(snapshotId as SnapshotId))
  );

  registerMcpTool(
    mcp,
    "excel.snapshot.delete",
    {
      title: "Delete snapshot",
      description: "Delete a stored snapshot.",
      inputSchema: { snapshotId: z.string() },
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async ({ snapshotId }: { snapshotId: string }) => jsonResult(runtime.deleteSnapshot(snapshotId as SnapshotId))
  );
}

function registerDiffTools(mcp: McpServer): void {
  for (const name of ["excel.diff.create", "excel.diff.summarize", "excel.diff.get_details", "excel.diff.export_json"] as const) {
    registerMcpTool(
      mcp,
      name,
      {
        title: name.replace(/^excel\./, "").replace(/\./g, " "),
        description: "Create or return a diff between two stored snapshots.",
        inputSchema: {
          leftSnapshotId: z.string(),
          rightSnapshotId: z.string()
        },
        annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
      },
      async ({ leftSnapshotId, rightSnapshotId }: { leftSnapshotId: string; rightSnapshotId: string }) => {
        const diff = runtime.compareSnapshots(leftSnapshotId as SnapshotId, rightSnapshotId as SnapshotId);
        return jsonResult(name.endsWith("export_json") ? { ok: true, json: JSON.stringify(diff, null, 2) } : diff);
      }
    );
  }

  registerMcpTool(
    mcp,
    "excel.diff.export_html",
    {
      title: "Export diff HTML",
      description: "Return a small HTML representation of a snapshot diff.",
      inputSchema: {
        leftSnapshotId: z.string(),
        rightSnapshotId: z.string()
      },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ leftSnapshotId, rightSnapshotId }: { leftSnapshotId: string; rightSnapshotId: string }) => {
      const diff = runtime.compareSnapshots(leftSnapshotId as SnapshotId, rightSnapshotId as SnapshotId);
      const escaped = escapeHtml(JSON.stringify(diff, null, 2));
      return jsonResult({ ok: true, html: `<html><body><pre>${escaped}</pre></body></html>` });
    }
  );
}

function registerEventTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.events.subscribe",
    {
      title: "Subscribe to Excel events",
      description: "Enable recent add-in event capture.",
      inputSchema: {},
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async () => jsonResult(runtime.subscribeEvents())
  );

  registerMcpTool(
    mcp,
    "excel.events.unsubscribe",
    {
      title: "Unsubscribe from Excel events",
      description: "Disable recent add-in event capture.",
      inputSchema: {},
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async () => jsonResult(runtime.unsubscribeEvents())
  );

  registerMcpTool(
    mcp,
    "excel.events.get_recent",
    {
      title: "Get recent Excel events",
      description: "Return recent add-in events observed by the backend.",
      inputSchema: { limit: z.number().int().positive().max(250).optional() },
      annotations: { readOnlyHint: true, destructiveHint: false, openWorldHint: false }
    },
    async ({ limit }: { limit?: number }) => jsonResult(runtime.getRecentEvents(limit))
  );

  registerMcpTool(
    mcp,
    "excel.events.clear",
    {
      title: "Clear Excel events",
      description: "Clear the in-memory event log.",
      inputSchema: {},
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async () => jsonResult(runtime.clearEvents())
  );

  registerMcpTool(
    mcp,
    "excel.events.set_debounce",
    {
      title: "Set Excel event debounce",
      description: "Set the event debounce preference stored by the backend.",
      inputSchema: { debounceMs: z.number().int().min(0).max(60000) },
      annotations: { readOnlyHint: false, destructiveHint: false, openWorldHint: false }
    },
    async ({ debounceMs }: { debounceMs: number }) => jsonResult(runtime.setEventDebounce(debounceMs))
  );
}

async function readRangeSnapshot(workbookId: string, sheetName: string, address: string) {
  const operation: ExcelOperation = {
    kind: "range.read_full",
    operationId: makeId<OperationId>("op"),
    workbookId: workbookId as WorkbookId,
    destructiveLevel: "none",
    reason: "MCP range read",
    target: {
      workbookId: workbookId as WorkbookId,
      sheetName,
      address
    }
  };
  return runtime.applyBatch({
    workbookId: workbookId as WorkbookId,
    mode: "apply",
    operations: [operation]
  });
}

function templateRepairSchema() {
  return {
    workbookId: z.string(),
    templateId: z.string(),
    targetSheetName: z.string(),
    repair: z.array(z.enum(["styles", "formulas", "dataRegions", "layout"])).optional()
  };
}

function tableSelectorSchema() {
  return {
    workbookId: z.string(),
    tableName: z.string()
  };
}

function tableSelector(args: { workbookId: string; tableName: string }): TableSelector {
  return {
    workbookId: args.workbookId as WorkbookId,
    tableName: args.tableName
  };
}

function tableCreateRequest(args: any): TableCreateRequest {
  const request: TableCreateRequest = {
    workbookId: args.workbookId as WorkbookId,
    sheetName: args.sheetName,
    address: args.address,
    hasHeaders: args.hasHeaders
  };
  if (args.tableName !== undefined) {
    request.tableName = args.tableName;
  }
  if (args.values !== undefined) {
    request.values = args.values;
  }
  if (args.style !== undefined) {
    request.style = args.style;
  }
  if (args.showTotals !== undefined) {
    request.showTotals = args.showTotals;
  }
  return request;
}

function tableFilterSchema() {
  return {
    ...tableSelectorSchema(),
    filters: z.array(
      z.object({
        column: z.union([z.string(), z.number().int().min(0)]),
        criteria: z.record(z.string(), z.any())
      })
    )
  };
}

function tableSortSchema() {
  return {
    ...tableSelectorSchema(),
    fields: z.array(
      z.object({
        key: z.number().int().min(0),
        ascending: z.boolean().optional(),
        sortOn: z.enum(["Value", "CellColor", "FontColor", "Icon"]).optional(),
        color: z.string().optional(),
        dataOption: z.enum(["Normal", "TextAsNumber"]).optional()
      })
    ),
    matchCase: z.boolean().optional(),
    method: z.enum(["PinYin", "StrokeCount"]).optional()
  };
}

function nameSelectorSchema() {
  return {
    workbookId: z.string(),
    name: z.string(),
    sheetName: z.string().optional()
  };
}

function nameMutationSchema() {
  return {
    ...nameSelectorSchema(),
    reference: z.string().optional(),
    formula: z.string().optional(),
    comment: z.string().optional(),
    visible: z.boolean().optional()
  };
}

function nameSelector(args: { workbookId: string; name: string; sheetName?: string }): NameSelector {
  const request: NameSelector = {
    workbookId: args.workbookId as WorkbookId,
    name: args.name
  };
  if (args.sheetName !== undefined) {
    request.sheetName = args.sheetName;
  }
  return request;
}

function nameCreateRequest(args: { workbookId: string; name: string; sheetName?: string; reference?: string; formula?: string; comment?: string; visible?: boolean }): NameCreateRequest {
  const request: NameCreateRequest = nameSelector(args);
  if (args.reference !== undefined) {
    request.reference = args.reference;
  }
  if (args.formula !== undefined) {
    request.formula = args.formula;
  }
  if (args.comment !== undefined) {
    request.comment = args.comment;
  }
  if (args.visible !== undefined) {
    request.visible = args.visible;
  }
  return request;
}

function nameUpdateRequest(args: { workbookId: string; name: string; sheetName?: string; reference?: string; formula?: string; comment?: string; visible?: boolean }): NameUpdateRequest {
  return nameCreateRequest(args);
}

function regionSelectorSchema() {
  return {
    workbookId: z.string(),
    regionName: z.string()
  };
}

function regionSelector(args: { workbookId: string; regionName: string }): RegionSelector {
  return {
    workbookId: args.workbookId as WorkbookId,
    regionName: args.regionName
  };
}

function pivotSelectorSchema() {
  return {
    workbookId: z.string(),
    pivotTableName: z.string()
  };
}

function pivotSelector(args: { workbookId: string; pivotTableName: string }): PivotSelector {
  return {
    workbookId: args.workbookId as WorkbookId,
    pivotTableName: args.pivotTableName
  };
}

function pivotValidateSourceRequest(args: {
  workbookId: string;
  pivotTableName: string;
  expectedFields?: string[];
  expectedRowFields?: string[];
  expectedColumnFields?: string[];
  expectedFilterFields?: string[];
  expectedDataFields?: string[];
  expectedDataFieldSettings?: PivotValidateSourceRequest["expectedDataFieldSettings"];
  expectedLayout?: PivotValidateSourceRequest["expectedLayout"];
}): PivotValidateSourceRequest {
  return {
    ...pivotSelector(args),
    ...(args.expectedFields !== undefined ? { expectedFields: args.expectedFields } : {}),
    ...(args.expectedRowFields !== undefined ? { expectedRowFields: args.expectedRowFields } : {}),
    ...(args.expectedColumnFields !== undefined ? { expectedColumnFields: args.expectedColumnFields } : {}),
    ...(args.expectedFilterFields !== undefined ? { expectedFilterFields: args.expectedFilterFields } : {}),
    ...(args.expectedDataFields !== undefined ? { expectedDataFields: args.expectedDataFields } : {}),
    ...(args.expectedDataFieldSettings !== undefined ? { expectedDataFieldSettings: args.expectedDataFieldSettings } : {}),
    ...(args.expectedLayout !== undefined ? { expectedLayout: args.expectedLayout } : {})
  };
}

function pivotCreateRequest(args: {
  workbookId: string;
  pivotTableName: string;
  sourceSheetName?: string;
  sourceAddress?: string;
  sourceTableName?: string;
  destinationSheetName: string;
  destinationAddress: string;
  rowFields?: string[];
  columnFields?: string[];
  filterFields?: string[];
  dataFields?: PivotCreateRequest["dataFields"];
  layout?: PivotCreateRequest["layout"];
  refresh?: boolean;
}): PivotCreateRequest {
  const request: PivotCreateRequest = {
    workbookId: args.workbookId as WorkbookId,
    pivotTableName: args.pivotTableName,
    destinationSheetName: args.destinationSheetName,
    destinationAddress: args.destinationAddress
  };
  if (args.sourceSheetName !== undefined) {
    request.sourceSheetName = args.sourceSheetName;
  }
  if (args.sourceAddress !== undefined) {
    request.sourceAddress = args.sourceAddress;
  }
  if (args.sourceTableName !== undefined) {
    request.sourceTableName = args.sourceTableName;
  }
  if (args.rowFields !== undefined) {
    request.rowFields = args.rowFields;
  }
  if (args.columnFields !== undefined) {
    request.columnFields = args.columnFields;
  }
  if (args.filterFields !== undefined) {
    request.filterFields = args.filterFields;
  }
  if (args.dataFields !== undefined) {
    request.dataFields = args.dataFields;
  }
  if (args.layout !== undefined) {
    request.layout = args.layout;
  }
  if (args.refresh !== undefined) {
    request.refresh = args.refresh;
  }
  return request;
}

function chartSelectorSchema() {
  return {
    workbookId: z.string(),
    sheetName: z.string(),
    chartName: z.string()
  };
}

function chartSelector(args: { workbookId: string; sheetName: string; chartName: string }): ChartSelector {
  return {
    workbookId: args.workbookId as WorkbookId,
    sheetName: args.sheetName,
    chartName: args.chartName
  };
}

function chartCreateSchema() {
  return {
    workbookId: z.string(),
    sheetName: z.string(),
    chartName: z.string().optional(),
    sourceAddress: z.string(),
    chartType: z.string(),
    seriesBy: z.enum(["Auto", "Columns", "Rows"]).optional(),
    title: z.string().optional(),
    position: z.object({ startCell: z.string(), endCell: z.string().optional() }).optional(),
    style: z.number().int().optional()
  };
}

function chartCreateRequest(args: {
  workbookId: string;
  sheetName: string;
  chartName?: string;
  sourceAddress: string;
  chartType: string;
  seriesBy?: "Auto" | "Columns" | "Rows";
  title?: string;
  position?: { startCell: string; endCell?: string };
  style?: number;
}): ChartCreateRequest {
  const request: ChartCreateRequest = {
    workbookId: args.workbookId as WorkbookId,
    sheetName: args.sheetName,
    sourceAddress: args.sourceAddress,
    chartType: args.chartType
  };
  if (args.chartName !== undefined) {
    request.chartName = args.chartName;
  }
  if (args.seriesBy !== undefined) {
    request.seriesBy = args.seriesBy;
  }
  if (args.title !== undefined) {
    request.title = args.title;
  }
  if (args.position !== undefined) {
    request.position = args.position;
  }
  if (args.style !== undefined) {
    request.style = args.style;
  }
  return request;
}

function chartUpdateDataSourceRequest(args: {
  workbookId: string;
  sheetName: string;
  chartName: string;
  sourceAddress: string;
  seriesBy?: "Auto" | "Columns" | "Rows";
}): ChartUpdateDataSourceRequest {
  const request: ChartUpdateDataSourceRequest = {
    ...chartSelector(args),
    sourceAddress: args.sourceAddress
  };
  if (args.seriesBy !== undefined) {
    request.seriesBy = args.seriesBy;
  }
  return request;
}

function chartTemplateValidationRequest(args: {
  workbookId: string;
  sheetName: string;
  chartName: string;
  templateSheetName?: string;
  templateChartName?: string;
}): ChartSelector & { templateSheetName?: string; templateChartName?: string } {
  const request: ChartSelector & { templateSheetName?: string; templateChartName?: string } = chartSelector(args);
  if (args.templateSheetName !== undefined) {
    request.templateSheetName = args.templateSheetName;
  }
  if (args.templateChartName !== undefined) {
    request.templateChartName = args.templateChartName;
  }
  return request;
}

function regionRegisterRequest(args: {
  workbookId: string;
  name: string;
  sheetName: string;
  address: string;
  kind?: RegionRegisterRequest["kind"];
  description?: string;
  templateId?: string;
  createNamedRange?: boolean;
}): RegionRegisterRequest {
  const request: RegionRegisterRequest = {
    workbookId: args.workbookId as WorkbookId,
    name: args.name,
    sheetName: args.sheetName,
    address: args.address
  };
  if (args.kind !== undefined) {
    request.kind = args.kind;
  }
  if (args.description !== undefined) {
    request.description = args.description;
  }
  if (args.templateId !== undefined) {
    request.templateId = args.templateId as TemplateId;
  }
  if (args.createNamedRange !== undefined) {
    request.createNamedRange = args.createNamedRange;
  }
  return request;
}

function permissionSetSchema() {
  return {
    allowWrites: z.boolean().optional(),
    allowDestructiveActions: z.boolean().optional(),
    allowWorkbookActions: z.boolean().optional(),
    allowMacroExecution: z.boolean().optional(),
    requireConfirmationFor: z.array(z.enum(["none", "values", "format", "structure", "workbook"])).optional(),
    scope: z
      .object({
        workbookId: z.string().optional(),
        sheetNames: z.array(z.string()).optional(),
        regionNames: z.array(z.string()).optional()
      })
      .optional()
  };
}

function permissionUpdate(args: any): Partial<PermissionState> {
  const update: Partial<PermissionState> = {};
  for (const key of ["allowWrites", "allowDestructiveActions", "allowWorkbookActions", "allowMacroExecution"] as const) {
    if (args[key] !== undefined) {
      update[key] = args[key];
    }
  }
  if (args.requireConfirmationFor !== undefined) {
    update.requireConfirmationFor = args.requireConfirmationFor;
  }
  if (args.scope !== undefined) {
    update.scope = permissionScope(args.scope);
  }
  return update;
}

function permissionScope(args: { workbookId?: string; sheetNames?: string[]; regionNames?: string[] }): PermissionState["scope"] {
  const scope: PermissionState["scope"] = {};
  if (args.workbookId !== undefined) {
    scope.workbookId = args.workbookId as WorkbookId;
  }
  if (args.sheetNames !== undefined) {
    scope.sheetNames = args.sheetNames;
  }
  if (args.regionNames !== undefined) {
    scope.regionNames = args.regionNames;
  }
  return scope;
}

function cleanRangeSchema() {
  return {
    workbookId: z.string(),
    sheetName: z.string(),
    address: z.string()
  };
}

function cleanRangeArgs<T extends Record<string, unknown> = Record<string, never>>(
  args: { workbookId: string; sheetName: string; address: string },
  extra?: T
): { workbookId: WorkbookId; sheetName: string; address: string } & T {
  return {
    workbookId: args.workbookId as WorkbookId,
    sheetName: args.sheetName,
    address: args.address,
    ...compactOptional(extra)
  } as { workbookId: WorkbookId; sheetName: string; address: string } & T;
}

function compactOptional<T extends Record<string, unknown>>(input: T | undefined): Partial<T> {
  if (!input) {
    return {};
  }
  return Object.fromEntries(Object.entries(input).filter(([, value]) => value !== undefined)) as Partial<T>;
}

function validationRangeArgs(args: { workbookId: string; sheetName?: string; address?: string }): {
  workbookId: WorkbookId;
  sheetName?: string;
  address?: string;
} {
  const request: { workbookId: WorkbookId; sheetName?: string; address?: string } = {
    workbookId: args.workbookId as WorkbookId
  };
  if (args.sheetName !== undefined) {
    request.sheetName = args.sheetName;
  }
  if (args.address !== undefined) {
    request.address = args.address;
  }
  return request;
}

async function repairTemplateFromArgs(args: {
  workbookId: string;
  templateId: string;
  targetSheetName: string;
  repair?: Array<"styles" | "formulas" | "dataRegions" | "layout">;
}) {
  const request: {
    workbookId: WorkbookId;
    templateId: TemplateId;
    targetSheetName: string;
    repair?: Array<"styles" | "formulas" | "dataRegions" | "layout">;
  } = {
    workbookId: args.workbookId as WorkbookId,
    templateId: args.templateId as TemplateId,
    targetSheetName: args.targetSheetName
  };
  if (args.repair !== undefined) {
    request.repair = args.repair;
  }
  return runtime.repairSheetFromTemplate(request);
}

function registerRangeOperation(
  mcp: McpServer,
  name: string,
  inputSchema: Record<string, unknown>,
  createOperation: (args: any) => Record<string, unknown>
): void {
  registerMcpTool(
    mcp,
    name,
    {
      title: name.replace(/^excel\./, "").replace(/\./g, " "),
      description: `Apply ${name} through the reversible batch pipeline.`,
      inputSchema,
      annotations: { readOnlyHint: false, destructiveHint: true, openWorldHint: false }
    },
    async (args: any) => jsonResult(await applySingleOperation(args.workbookId, createOperation(args)))
  );
}

function registerSheetOperation(
  mcp: McpServer,
  name: string,
  inputSchema: Record<string, unknown>,
  createOperation: (args: any) => Record<string, unknown>
): void {
  registerRangeOperation(mcp, name, inputSchema, createOperation);
}

async function applySingleOperation(workbookId: string, operation: Record<string, unknown>) {
  return runtime.applyBatch({
    workbookId: workbookId as WorkbookId,
    mode: "apply",
    operations: [
      {
        ...operation,
        operationId: makeId<OperationId>("op")
      } as ExcelOperation
    ]
  });
}

function targetFromArgs(args: { workbookId: string; sheetName: string; address: string }): A1Range {
  return {
    workbookId: args.workbookId as WorkbookId,
    sheetName: args.sheetName,
    address: args.address
  };
}

function rangeMetadataRequest(args: { workbookId: string; sheetName: string; address: string }) {
  return {
    workbookId: args.workbookId as WorkbookId,
    sheetName: args.sheetName,
    address: args.address
  };
}

function rangeSearchRequest(args: {
  workbookId: string;
  sheetName: string;
  address: string;
  text: string;
  completeMatch?: boolean;
  matchCase?: boolean;
  searchDirection?: "Forward" | "Backwards";
}) {
  const request: {
    workbookId: WorkbookId;
    sheetName: string;
    address: string;
    text: string;
    completeMatch?: boolean;
    matchCase?: boolean;
    searchDirection?: "Forward" | "Backwards";
  } = {
    workbookId: args.workbookId as WorkbookId,
    sheetName: args.sheetName,
    address: args.address,
    text: args.text
  };
  if (args.completeMatch !== undefined) {
    request.completeMatch = args.completeMatch;
  }
  if (args.matchCase !== undefined) {
    request.matchCase = args.matchCase;
  }
  if (args.searchDirection !== undefined) {
    request.searchDirection = args.searchDirection;
  }
  return request;
}

function snapshotInputSchema() {
  return {
    workbookId: z.string(),
    reason: z.string().optional(),
    ranges: z
      .array(
        z.object({
          workbookId: z.string(),
          sheetName: z.string(),
          address: z.string()
        })
      )
      .optional()
  };
}

function snapshotRequest(workbookId: string, reason?: string, ranges?: A1Range[]): { workbookId: WorkbookId; reason?: string; ranges?: A1Range[] } {
  const request: { workbookId: WorkbookId; reason?: string; ranges?: A1Range[] } = {
    workbookId: workbookId as WorkbookId
  };
  if (reason !== undefined) {
    request.reason = reason;
  }
  if (ranges !== undefined) {
    request.ranges = ranges;
  }
  return request;
}

async function selectSheetInfo(sheetName: string): Promise<{ ok: boolean; sheet?: any; result: unknown }> {
  const result = await runtime.getWorkbookMap();
  const map = "map" in result ? result.map as { sheets?: any[] } : undefined;
  const sheet = map?.sheets?.find((candidate) => candidate.name === sheetName);
  return { ok: Boolean(result.ok && sheet), sheet, result };
}

function escapeHtml(value: string): string {
  return value.replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");
}

function registerMcpTool(
  mcp: McpServer,
  name: string,
  config: Record<string, unknown>,
  callback: (args: any, extra: any) => any
): void {
  if (!isToolExposed(name, catalogOptions)) {
    return;
  }
  (mcp.registerTool as any)(name, config, callback);
}

function jsonResult(value: unknown) {
  return {
    content: [
      {
        type: "text" as const,
        text: JSON.stringify(value, null, 2)
      }
    ]
  };
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
  if (Array.isArray(value)) {
    return value.join("/");
  }
  return value ?? "";
}

function stripResourceSheetName(address: string): string {
  const bangIndex = address.lastIndexOf("!");
  return bangIndex >= 0 ? address.slice(bangIndex + 1) : address;
}

function hasArg(name: string): boolean {
  return process.argv.includes(name);
}

function readArg(name: string): string | undefined {
  const prefix = `${name}=`;
  const inline = process.argv.find((arg) => arg.startsWith(prefix));
  if (inline) {
    return inline.slice(prefix.length);
  }
  const index = process.argv.indexOf(name);
  if (index >= 0) {
    return process.argv[index + 1];
  }
  return undefined;
}

function trimTrailingSlash(value: string): string {
  return value.endsWith("/") ? value.slice(0, -1) : value;
}
