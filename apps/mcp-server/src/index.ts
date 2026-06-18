#!/usr/bin/env node
import { createHash } from "node:crypto";
import { McpServer, ResourceTemplate } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import * as z from "zod/v4";
import { RuntimeService } from "@components-kit/open-workbook-backend/runtime";
import { startBackendServer } from "@components-kit/open-workbook-backend/server";
import type {
  A1Range,
  AgentId,
  AgentRunInput,
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
  JobId,
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
  RangeCompactReadRequest,
  TableAppendRowsRequest,
  TableApplyFiltersRequest,
  TableCompactReadRequest,
  TableCopyStructureRequest,
  TableCreateRequest,
  TableReadRequest,
  TableReorderColumnsRequest,
  TableResizeRequest,
  TableSelector,
  TableSetStyleRequest,
  TableSetTotalRowRequest,
  TableSortRequest,
  TableUpdateRowsRequest,
  LookupFindEntityRequest,
  LookupFindHeadersRequest,
  LookupFindTablesByColumnsRequest,
  LookupInspectMatchRequest,
  LookupResolveRangeRequest,
  LookupWorkbookSearchRequest,
  TemplateId,
  TransactionId,
  WorkbookScope,
  WorkbookId,
  WorkbookBackupRetentionRequest,
  WorkbookCreateFileBackupRequest,
  WorkbookRestoreFileBackupRequest,
  WorkbookLocalConfig
} from "@components-kit/open-workbook-protocol";
import { makeId } from "@components-kit/open-workbook-protocol";

type RuntimeFacade = RuntimeService & {
  compileBatch(request: BatchRequest): unknown;
};

type CompactCellOutput = "auto" | "matrix" | "sparse";

type ContextRangeCompactReadRequest = Omit<RangeCompactReadRequest, "workbookId" | "sheetName"> & {
  workbookId?: WorkbookId | string;
  workbookContextId?: string;
  sheetName?: string;
  cellOutput?: CompactCellOutput;
};

type ContextTableCompactReadRequest = Omit<TableCompactReadRequest, "workbookId" | "tableName"> & {
  workbookId?: WorkbookId | string;
  workbookContextId?: string;
  tableName?: string;
};

interface AgentWorkbookContextResource {
  ok: true;
  workbookContextId: string;
  workbook: {
    workbookId?: WorkbookId | string;
    name: string;
    activeSheet?: string;
    sheetCount: number;
  };
  selection?: unknown;
  sheets: Array<{
    name: string;
    kind?: string;
    usedRange?: string;
    rowCount?: number;
    columnCount?: number;
    headers?: unknown[];
    tableIds?: string[];
  }>;
  tables: Array<{
    name?: string;
    sheetName: string;
    range: string;
    columns?: unknown[];
  }>;
  namedRanges?: unknown[];
  summaryBlocks?: unknown[];
  formulaRegions?: unknown[];
}

type AgentContextResolution<T> =
  | { ok: true; context: AgentWorkbookContextResource; value: T }
  | { ok: false; response: Record<string, unknown> };

const host = process.env.OPEN_WORKBOOK_HOST ?? "127.0.0.1";
const port = Number(process.env.OPEN_WORKBOOK_PORT ?? 37845);
const addinPath = process.env.OPEN_WORKBOOK_ADDIN_PATH ?? "/addin";
const daemonUrl = trimTrailingSlash(readArg("--daemon-url") ?? process.env.OPEN_WORKBOOK_DAEMON_URL ?? `http://${host}:${port}`);
const agentName = readArg("--agent-name") ?? process.env.OPEN_WORKBOOK_AGENT_NAME;
const standalone = hasArg("--standalone") || process.env.OPEN_WORKBOOK_MCP_STANDALONE === "1";

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
const runtimeVersion = process.env.OPEN_WORKBOOK_VERSION ?? "0.1.14";
const COMPACT_RESOURCE_LIMIT = 100;
const COMPACT_DEFAULT_RESOURCE_THRESHOLD_BYTES = 24_000;
const COMPACT_LIMITS = {
  maxToolResultChars: 4_000,
  maxSummaryChars: 2_000,
  maxExamples: 5,
  maxWarnings: 10,
  maxValidationIssues: 20,
  maxSchemaFields: 50,
  maxSampleRows: 10,
  maxSampleCols: 20
} as const;

type CompactResourceKind = "read" | "validation" | "diff" | "snapshot" | "mutation" | "summary" | "generic";
type CompactResponseMode = "brief" | "standard" | "verbose";
type CompactNextActionRecommendation = "answer_now" | "fetch_more_context" | "validate_then_answer" | "needs_user_confirmation";
type CompactConfidence = "high" | "medium" | "low";

interface CompactStoredResource {
  resourceId: string;
  uri: string;
  kind: CompactResourceKind;
  title?: string | undefined;
  scope?: Record<string, unknown> | undefined;
  sourceHash?: string | undefined;
  pinned?: boolean | undefined;
  createdAt: string;
  lastAccessedAt: string;
  payloadBytes: number;
  estimatedTokens: number;
  accessCount: number;
  payload: unknown;
}

interface CompactCacheEntry {
  key: string;
  createdAt: string;
  value: unknown;
}

const compactResources = new Map<string, CompactStoredResource>();
const compactCache = new Map<string, CompactCacheEntry>();
const compactIdempotencyRecords = new Map<string, CompactIdempotencyRecord>();
let compactCacheInvalidationCount = 0;
let compactCacheLastInvalidatedAt: string | undefined;
let compactLastObservedEventId: string | undefined;
let compactToolResultCount = 0;
let compactToolResultBytes = 0;
let compactToolStoredBytes = 0;
let compactCacheHitCount = 0;
let compactCacheMissCount = 0;
const COMPACT_RESOURCE_TTL_MS = 60 * 60 * 1000;
const COMPACT_IDEMPOTENCY_TTL_MS = 60 * 60 * 1000;

interface CompactIdempotencyRecord {
  idempotencyKey: string;
  toolName: string;
  operationHash: string;
  createdAt: string;
  resultText: string;
}

const server = new McpServer({
  name: "open-workbook",
  version: runtimeVersion
});

registerAgentTools(server);
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
        if (property === "then") {
          return undefined;
        }
        return (...args: unknown[]) => call(property, args);
      }
    }
  );
}

function registerAgentTools(mcp: McpServer): void {
  registerMcpTool(
    mcp,
    "excel.agent.run",
    {
      title: "Run Open Workbook agent workflow",
      description:
        "Single default Open Workbook interface. Send workbook intent; the backend handles discovery, cached metadata, target resolution, preview/apply, validation, rollback, and compact proof without exposing low-level Excel tools.",
      inputSchema: {
        request: z.string(),
        mode: z.enum(["auto", "status", "prepare", "find", "answer", "preview_update", "apply_update", "validate", "rollback"]).optional(),
        workbookContextId: z.string().optional(),
        operationId: z.string().optional(),
        transactionId: z.string().optional(),
        confirmationToken: z.string().optional(),
        intent: z.object({
          action: z.enum([
            "read_values",
            "read_schema",
            "find_target",
            "write_values",
            "write_formulas",
            "format_range",
            "clear_values",
            "append_table_rows",
            "sort_table",
            "filter_range",
            "autofit",
            "copy_template_sheet",
            "calculate",
            "save"
          ]),
          confidence: z.number().min(0).max(1).optional(),
          reason: z.string().optional(),
          targetHints: z.array(z.string()).optional()
        }).optional(),
        target: z.object({
          workbookId: z.string().optional(),
          workbookName: z.string().optional(),
          candidateId: z.string().optional(),
          sheetName: z.string().optional(),
          tableName: z.string().optional(),
          range: z.string().optional(),
          row: z.number().int().optional(),
          column: z.string().optional(),
          entity: z.string().optional()
        }).optional(),
        values: z.record(z.string(), z.any()).and(z.object({
          patches: z.array(z.object({
            target: z.object({
              workbookId: z.string().optional(),
              workbookName: z.string().optional(),
              candidateId: z.string().optional(),
              sheetName: z.string().optional(),
              tableName: z.string().optional(),
              range: z.string().optional(),
              row: z.number().int().optional(),
              column: z.string().optional(),
              entity: z.string().optional()
            }),
            values: z.array(z.array(z.any())).optional(),
            rows: z.array(z.array(z.any())).optional(),
            reason: z.string().optional()
          })).optional()
        })).optional(),
        responseMode: z.enum(["brief", "standard", "verbose"]).optional(),
        budget: z.object({
          maxPayloadBytes: z.number().int().positive().optional(),
          maxEstimatedTokens: z.number().int().positive().optional(),
          maxExamples: z.number().int().positive().optional()
        }).optional()
      },
      outputSchema: agentRunOutputSchema(),
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        idempotentHint: false,
        openWorldHint: false
      }
    },
    async (args: AgentRunInput) => agentJsonResult(await runtime.runAgent(args))
  );
}

function agentRunOutputSchema() {
  return {
    status: z.enum(["SUCCESS", "PREVIEW_READY", "NEEDS_INPUT", "AMBIGUOUS_TARGET", "NOT_FOUND", "STALE_CONTEXT", "VALIDATION_FAILED", "CONFLICT", "ERROR"]),
    mode: z.string(),
    workbookContextId: z.string().optional(),
    operationId: z.string().optional(),
    transactionId: z.string().optional(),
    confirmationToken: z.string().optional(),
    summary: z.string(),
    answer: z.any().optional(),
    metrics: z.record(z.string(), z.any()).optional(),
    changes: z.array(z.any()).optional(),
    candidates: z.array(z.any()).optional(),
    proof: z.array(z.any()),
    resourceLinks: z.array(z.any()),
    nextAction: z.string(),
    warnings: z.array(z.string()),
    telemetry: z.object({
      internalCallCount: z.number(),
      payloadBytes: z.number(),
      estimatedTokens: z.number(),
      elapsedMs: z.number(),
      cacheHit: z.boolean(),
      autoApplied: z.boolean().optional(),
      safetyDecision: z.string().optional(),
      previewOperationId: z.string().optional(),
      validationStatus: z.enum(["passed", "failed", "not_run"]).optional(),
      metadataCacheStatus: z.enum(["hit", "miss", "not_applicable"]).optional(),
      internalReadCount: z.number().optional(),
      fullReadCellCount: z.number().optional(),
      candidateCount: z.number().optional(),
      resourceLinkCount: z.number().optional(),
      estimatedTokensSaved: z.number().optional(),
      routeMode: z.string().optional(),
      routeMatchedRule: z.string().optional(),
      routeConfidence: z.number().optional(),
      routeReasons: z.array(z.string()).optional(),
      operationRisk: z.string().optional(),
      actionHandlerId: z.string().optional(),
      autoApplyBlockedReason: z.string().optional(),
      targetFingerprintStatus: z.enum(["matched", "changed", "not_applicable"]).optional(),
      targetHintCount: z.number().optional(),
      targetHintUsed: z.boolean().optional(),
      intentSource: z.enum(["caller_structured", "deterministic_fallback", "mixed"]).optional(),
      intentAction: z.string().optional(),
      intentAccepted: z.boolean().optional(),
      intentRejectedReason: z.string().optional()
    })
  };
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

  registerJsonTemplateResource(
    mcp,
    "compact detail resource",
    "excel://compact/{resource_id}",
    "Stored compact-context detail payload returned by token-saving Open Workbook tools.",
    async (_uri, variables) => getCompactResource(resourceVariable(variables, "resource_id"))
  );

  registerJsonTemplateResource(
    mcp,
    "agent workbook context",
    "excel://agent/contexts/{workbook_context_id}",
    "Cached workbook metadata used by the Open Workbook agent workflow.",
    async (_uri, variables) => runtime.getAgentContextResource(resourceVariable(variables, "workbook_context_id"))
  );

  registerJsonTemplateResource(
    mcp,
    "agent pending operation",
    "excel://agent/operations/{operation_id}",
    "Pending previewed workbook operation awaiting apply confirmation.",
    async (_uri, variables) => runtime.getAgentOperationResource(resourceVariable(variables, "operation_id"))
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
      "4. Use `excel.template.create_sheet_from_template` with the confirmed template or previous-period sheet as the source.",
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
      "4. Create a repair plan using `excel.formula.repair_patterns`, `fill_down`, `fill_right`, or explicit `range.write_formulas`; never repair formulas by writing formula strings through `range.write_values`.",
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
      "4. Use table reads, formulas, PivotTables, and charts through available agent workflows and internal operations.",
      "5. Preview and apply via plan/batch; never write directly outside the target report regions.",
      "6. Validate formulas, style consistency, tables, charts, and no unintended source changes."
    ]
  );

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

function promptContext(args: Record<string, unknown>): string {
  const entries = Object.entries(args).filter(([, value]) => value !== undefined && value !== "");
  if (entries.length === 0) {
    return "";
  }
  return `Context: ${entries.map(([key, value]) => `${key}=${String(value)}`).join(", ")}`;
}



function runtimeCapabilities(includePreview?: boolean) {
  const capabilities = runtime.getCapabilities(includePreview === undefined ? {} : { includePreview });
  const typed = capabilities as Record<string, any>;
  const catalog = typed.catalog as Record<string, any> | undefined;
  const internalCapabilities = typed.internalCapabilities as Record<string, any> | undefined;
  const exposedToolNames = exposedProfileToolNames();
  return {
    ...typed,
    catalog: catalog
      ? {
          total: catalog.total,
          stable: catalog.stable,
          preview: catalog.preview,
          planned: catalog.planned,
          unsupported: catalog.unsupported,
          exposed: exposedToolNames.length,
          tools: exposedToolNames
        }
      : undefined,
    internalCapabilities: internalCapabilities
      ? {
          total: internalCapabilities.total,
          exposed: 0,
          stable: internalCapabilities.stable,
          preview: internalCapabilities.preview,
          planned: internalCapabilities.planned,
          unsupported: internalCapabilities.unsupported,
          capabilityNames: Array.isArray(internalCapabilities.capabilities)
            ? internalCapabilities.capabilities.map((capability: { name?: unknown }) => capability.name).filter((name: unknown): name is string => typeof name === "string")
            : []
        }
      : undefined,
    resources: Array.isArray(typed.resources) ? { count: typed.resources.length } : typed.resources,
    prompts: Array.isArray(typed.prompts) ? { count: typed.prompts.length } : typed.prompts
  };
}

function exposedProfileToolNames(): string[] {
  return ["excel.agent.run"];
}















async function workflowPreflight(workbookId?: WorkbookId, includePreview = true) {
  const status = runtime.getStatus();
  const activeContext = await runtime.getActiveContext();
  const activeWorkbookId = (activeContext as { activeWorkbook?: { workbookId?: WorkbookId } }).activeWorkbook?.workbookId;
  const resolvedWorkbookId = workbookId ?? activeWorkbookId;
  return {
    status,
    activeContext,
    capabilities: runtimeCapabilities(includePreview),
    workbookMap: await runtime.getWorkbookMap(),
    collaboration: runtime.getCollaborationStatus(resolvedWorkbookId),
    workbookId: resolvedWorkbookId
  };
}

async function workflowPrepareSession(workbookId?: WorkbookId, includePreview = false) {
  const preflight = await workflowPreflight(workbookId, includePreview);
  const payload = {
    ok: true,
    workflow: "excel.workflow.prepare_session",
    ...preflight
  };
  const summary = {
    ok: true,
    workflow: "excel.workflow.prepare_session",
    workbookId: preflight.workbookId,
    status: summarizeRuntimeStatus(preflight.status),
    activeContext: preflight.activeContext,
    capabilities: summarizeCapabilities(preflight.capabilities),
    workbookMap: summarizeWorkbookMap(preflight.workbookMap),
    collaboration: summarizeCollaboration(preflight.collaboration)
  };
  const sourceHash = compactSourceHash(payload);
  return withCompactTelemetry(summary, {
    detailLevel: "summary",
    responseMode: "brief",
    storeResource: true,
    resourceKind: "summary",
    resourceTitle: "Workflow prepare session detail",
    resourcePayload: payload,
    resourceScope: compactReadScope({ workbookId: preflight.workbookId, workflow: "excel.workflow.prepare_session" }),
    sourceHash,
    nextActionRecommendation: "answer_now",
    reasoningHints: ["Workflow discovery completed", "Capability catalog detail is stored behind contextId", "Use excel.runtime.get_capabilities when full tool catalog detail is required"],
    confidence: "high",
    confidenceReasons: ["Runtime status, active context, workbook map, and collaboration status were collected"],
    maxPayloadBytes: COMPACT_LIMITS.maxToolResultChars,
    budgetSummary: {
      ok: true,
      workflow: "excel.workflow.prepare_session",
      workbookId: preflight.workbookId,
      activeContext: summarizeActiveContext(preflight.activeContext),
      status: summary.status,
      workbookMap: summary.workbookMap,
      sourceHash
    }
  });
}

function summarizeActiveContext(activeContext: unknown): unknown {
  if (!activeContext || typeof activeContext !== "object") {
    return activeContext;
  }
  const typed = activeContext as { activeWorkbook?: unknown; activeSheet?: unknown; selection?: unknown; [key: string]: unknown };
  return {
    activeWorkbook: typed.activeWorkbook,
    activeSheet: typed.activeSheet,
    selection: typed.selection
  };
}

function summarizeRuntimeStatus(status: unknown): unknown {
  if (!status || typeof status !== "object") {
    return status;
  }
  const typed = status as { sessions?: any[]; [key: string]: unknown };
  return {
    ...typed,
    sessions: Array.isArray(typed.sessions)
      ? typed.sessions.slice(0, 10).map((session) => ({
          sessionId: session.sessionId,
          connectedAt: session.connectedAt,
          lastSeenAt: session.lastSeenAt,
          activeWorkbook: session.activeWorkbook
        }))
      : typed.sessions,
    sessionCount: Array.isArray(typed.sessions) ? typed.sessions.length : undefined
  };
}

function summarizeCapabilities(capabilities: unknown): unknown {
  if (!capabilities || typeof capabilities !== "object") {
    return capabilities;
  }
  const typed = capabilities as { catalog?: any; resources?: unknown; prompts?: unknown; [key: string]: unknown };
  const catalog = typed.catalog && typeof typed.catalog === "object" ? typed.catalog : undefined;
  return {
    ...typed,
    catalog: catalog
      ? {
          total: catalog.total,
          stable: catalog.stable,
          preview: catalog.preview,
          planned: catalog.planned,
          unsupported: catalog.unsupported,
          exposed: catalog.exposed,
          toolCount: Array.isArray(catalog.tools) ? catalog.tools.length : undefined
        }
      : undefined
  };
}

function summarizeWorkbookMap(workbookMap: unknown): unknown {
  if (!workbookMap || typeof workbookMap !== "object") {
    return workbookMap;
  }
  const typed = workbookMap as { ok?: boolean; map?: { workbook?: any; sheets?: any[] }; [key: string]: unknown };
  const map = typed.map;
  if (!typed.ok || !map) {
    return typed;
  }
  const sheets = Array.isArray(map.sheets) ? map.sheets : [];
  const tableCount = sheets.reduce((count, sheet) => count + (Array.isArray(sheet.tables) ? sheet.tables.length : 0), 0);
  return {
    ok: true,
    workbook: map.workbook,
    sheetCount: sheets.length,
    tableCount,
    sheets: sheets.slice(0, 20).map((sheet) => ({
      name: sheet.name,
      position: sheet.position,
      visibility: sheet.visibility,
      usedRange: sheet.usedRange,
      tableCount: Array.isArray(sheet.tables) ? sheet.tables.length : 0,
      tables: Array.isArray(sheet.tables) ? sheet.tables.slice(0, 20).map((table: { name?: string }) => table.name) : []
    }))
  };
}

function summarizeCollaboration(collaboration: unknown): unknown {
  if (!collaboration || typeof collaboration !== "object") {
    return collaboration;
  }
  const typed = collaboration as Record<string, unknown>;
  const output: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(typed)) {
    if (Array.isArray(value)) {
      output[`${key}Count`] = value.length;
      output[key] = value.slice(0, 10);
    } else {
      output[key] = value;
    }
  }
  return output;
}

async function workflowInspectAnalyze(args: {
  workbookId: string;
  tableName?: string;
  sheetName?: string;
  address?: string;
  maxRows?: number;
  responseMode?: CompactResponseMode;
  budget?: Record<string, unknown>;
}) {
  const workbookId = args.workbookId as WorkbookId;
  const budget = compactRequestedBudget(args);
  const maxRows = Math.min(args.maxRows ?? 1000, 10000);
  const source = args.tableName
    ? await workflowAnalyzeTable(workbookId, args.tableName, maxRows)
    : args.sheetName && args.address
      ? await workflowAnalyzeRange(workbookId, args.sheetName, args.address)
      : { ok: false, error: { code: "ANALYZE_TARGET_REQUIRED", message: "Provide either tableName or sheetName plus address." } };
  if (!(source as { ok?: boolean }).ok) {
    return withCompactTelemetry(
      { ok: false, workflow: "excel.workflow.inspect_analyze", workbookId, tableName: args.tableName, sheetName: args.sheetName, address: args.address, source },
      {
        detailLevel: "summary",
        responseMode: args.responseMode,
        nextActionRecommendation: "fetch_more_context",
        reasoningHints: ["Analysis target could not be read"],
        confidence: "low",
        confidenceReasons: ["Read failed before local analysis"],
        maxPayloadBytes: budget.maxChars
      }
    );
  }
  const matrix = (source as { rows?: unknown[][] }).rows ?? [];
  const headers = (source as { headers?: string[] }).headers ?? workflowDefaultHeaders(matrix);
  const rows = workflowDataRows(headers, matrix);
  const analysis = analyzeTabularRows(headers, rows);
  const fullPayload = { ok: true, workflow: "excel.workflow.inspect_analyze", source, analysis };
  const summary = {
    ok: true,
    workflow: "excel.workflow.inspect_analyze",
    workbookId,
    tableName: args.tableName,
    sheetName: args.sheetName,
    address: args.address,
    shape: analysis.shape,
    detectedColumns: analysis.columns.slice(0, budget.maxExamples),
    issueSummary: analysis.issueSummary,
    numericSummary: analysis.numericSummary.slice(0, budget.maxExamples),
    duplicateRowCount: analysis.duplicateRowCount,
    budgetSummary: budget.applied
  };
  return withCompactTelemetry(summary, {
    detailLevel: "summary",
    responseMode: args.responseMode,
    storeResource: true,
    resourceKind: "summary",
    resourceTitle: "Workflow inspect/analyze detail",
    resourcePayload: fullPayload,
    resourceScope: compactReadScope({ workbookId, tableName: args.tableName, sheetName: args.sheetName, address: args.address }),
    sourceHash: compactSourceHash(fullPayload),
    nextActionRecommendation: "answer_now",
    reasoningHints: ["Local analysis completed", "Full row/profile detail is stored behind contextId", "Agent can answer from the compact summary"],
    confidence: "high",
    confidenceReasons: ["Data was read locally", "Deterministic profiling completed inside MCP"],
    maxPayloadBytes: budget.maxChars,
    budgetSummary: summary
  });
}

async function workflowAnalyzeTable(workbookId: WorkbookId, tableName: string, maxRows: number) {
  const result = await runtime.readTable({ workbookId, tableName, includeValues: true, rowOffset: 0, rowLimit: maxRows });
  const table = (result as { table?: { headers?: string[]; values?: unknown[][] } }).table;
  if (!(result as { ok?: boolean }).ok || !table) {
    return { ok: false, source: result };
  }
  return { ok: true, sourceType: "table", workbookId, tableName, headers: table.headers ?? workflowDefaultHeaders(table.values ?? []), rows: table.values ?? [] };
}

async function workflowAnalyzeRange(workbookId: WorkbookId, sheetName: string, address: string) {
  const result = await readRangeSnapshot(workbookId, sheetName, address, ["values"]);
  const snapshot = ((result as { data?: Array<{ snapshot?: { values?: unknown[][] } }> }).data ?? [])[0]?.snapshot;
  if (!(result as { ok?: boolean }).ok || !snapshot) {
    return { ok: false, source: result };
  }
  const values = snapshot.values ?? [];
  return { ok: true, sourceType: "range", workbookId, sheetName, address, headers: workflowDefaultHeaders(values), rows: values };
}

function workflowDefaultHeaders(matrix: unknown[][]): string[] {
  const firstRow = matrix[0] ?? [];
  if (firstRow.length > 0 && firstRow.every((value) => typeof value === "string" && value.trim().length > 0)) {
    return firstRow.map((value) => String(value));
  }
  const width = matrix.reduce((max, row) => Math.max(max, Array.isArray(row) ? row.length : 0), 0);
  return Array.from({ length: width }, (_unused, index) => `Column ${index + 1}`);
}

function workflowDataRows(headers: string[], matrix: unknown[][]): unknown[][] {
  if (matrix.length === 0) {
    return [];
  }
  const firstRow = matrix[0] ?? [];
  const hasHeaderRow = headers.length > 0 && firstRow.length === headers.length && firstRow.every((value, index) => String(value) === headers[index]);
  return hasHeaderRow ? matrix.slice(1) : matrix;
}

function analyzeTabularRows(headers: string[], rows: unknown[][]) {
  const rowHashes = new Map<string, number>();
  let duplicateRowCount = 0;
  for (const row of rows) {
    const key = JSON.stringify(row);
    const count = rowHashes.get(key) ?? 0;
    if (count === 1) {
      duplicateRowCount += 1;
    }
    rowHashes.set(key, count + 1);
  }
  const columns = headers.map((header, columnIndex) => analyzeColumn(header, rows.map((row) => row[columnIndex])));
  return {
    shape: { rows: rows.length, columns: headers.length },
    columns,
    issueSummary: {
      missingValues: columns.reduce((sum, column) => sum + column.missingCount, 0),
      duplicateRows: duplicateRowCount,
      formulaErrors: columns.reduce((sum, column) => sum + column.formulaErrorCount, 0)
    },
    numericSummary: columns.filter((column) => column.inferredType === "number").map((column) => ({
      name: column.name,
      count: column.nonEmptyCount,
      min: column.min,
      max: column.max,
      average: column.average
    })),
    duplicateRowCount
  };
}

function analyzeColumn(name: string, values: unknown[]) {
  const nonEmpty = values.filter((value) => !isCompactEmptyCell(value));
  const numericValues = nonEmpty
    .map((value) => typeof value === "number" ? value : typeof value === "string" && value.trim() !== "" ? Number(value) : NaN)
    .filter((value) => Number.isFinite(value));
  const formulaErrorCount = nonEmpty.filter((value) => typeof value === "string" && /^#(VALUE!|REF!|DIV\/0!|NAME\?|N\/A|NULL!|NUM!)/.test(value)).length;
  const inferredType = numericValues.length > 0 && numericValues.length >= nonEmpty.length * 0.8
    ? "number"
    : nonEmpty.some((value) => value instanceof Date || (typeof value === "string" && !Number.isNaN(Date.parse(value))))
      ? "date"
      : "text";
  const sum = numericValues.reduce((total, value) => total + value, 0);
  return {
    name,
    inferredType,
    nonEmptyCount: nonEmpty.length,
    missingCount: values.length - nonEmpty.length,
    distinctCount: new Set(nonEmpty.map((value) => JSON.stringify(value))).size,
    formulaErrorCount,
    ...(numericValues.length > 0 ? { min: Math.min(...numericValues), max: Math.max(...numericValues), average: Number((sum / numericValues.length).toFixed(6)) } : {})
  };
}

async function workflowRollbackValidate(args: {
  workbookId: string;
  transactionId?: string;
  backupId?: string;
  confirmationToken?: string;
  responseMode?: CompactResponseMode;
  budget?: Record<string, unknown>;
}) {
  const workbookId = args.workbookId as WorkbookId;
  const budget = compactRequestedBudget(args);
  if (!args.transactionId && !args.backupId) {
    return withCompactTelemetry(
      {
        ok: false,
        workflow: "excel.workflow.rollback_validate",
        workbookId,
        error: {
          code: "ROLLBACK_TARGET_REQUIRED",
          message: "Provide transactionId or backupId."
        }
      },
      {
        detailLevel: "summary",
        responseMode: args.responseMode,
        nextActionRecommendation: "needs_user_confirmation",
        reasoningHints: ["Rollback target was missing"],
        confidence: "high",
        confidenceReasons: ["No rollback was attempted"],
        maxPayloadBytes: budget.maxChars
      }
    );
  }
  const rollbackResult = args.transactionId
    ? await runtime.rollbackTransaction(args.transactionId as TransactionId, args.confirmationToken)
    : await runtime.restoreBackup(args.backupId as BackupId, args.confirmationToken);
  const calculation = await runtime.calculateWorkbook(workbookId, "recalculate");
  const validation = await runtime.validateWorkbook({ workbookId });
  const validationIssues = Array.isArray((validation as { issues?: unknown[] }).issues) ? (validation as { issues: unknown[] }).issues : [];
  const ok = Boolean((rollbackResult as { ok?: boolean }).ok && (calculation as { ok?: boolean }).ok !== false && (validation as { ok?: boolean }).ok);
  const fullPayload = { ok, workflow: "excel.workflow.rollback_validate", rollbackResult, calculation, validation };
  const summary = {
    ok,
    workflow: "excel.workflow.rollback_validate",
    workbookId,
    transactionId: args.transactionId,
    backupId: args.backupId,
    rollback: summarizeOperationResult(rollbackResult),
    calculationOk: (calculation as { ok?: boolean }).ok,
    validationSummary: {
      ok: (validation as { ok?: boolean }).ok,
      issueCount: (validation as { issueCount?: number }).issueCount ?? validationIssues.length,
      severityCounts: compactSeverityCounts(validationIssues),
      categories: compactIssueCategories(validationIssues),
      examples: validationIssues.slice(0, budget.maxIssues).map(compactIssueExample),
      examplesTruncated: validationIssues.length > budget.maxIssues
    },
    budgetSummary: budget.applied
  };
  return withCompactTelemetry(summary, {
    detailLevel: "summary",
    responseMode: args.responseMode,
    storeResource: true,
    resourceKind: "mutation",
    resourceTitle: "Workflow rollback/validate detail",
    resourcePayload: fullPayload,
    resourceScope: compactReadScope({ workbookId, transactionId: args.transactionId, backupId: args.backupId }),
    sourceHash: compactSourceHash(fullPayload),
    nextActionRecommendation: ok ? "answer_now" : "fetch_more_context",
    reasoningHints: ok
      ? ["Rollback completed", "Workbook recalculation and validation completed", "Agent can answer now"]
      : ["Rollback workflow reported issues", "Fetch context preview before final answer"],
    confidence: ok && validationIssues.length === 0 ? "high" : ok ? "medium" : "low",
    confidenceReasons: [
      (rollbackResult as { ok?: boolean }).ok === false ? "Rollback reported failure" : "Rollback step completed",
      (calculation as { ok?: boolean }).ok === false ? "Recalculation reported failure" : "Recalculation step completed",
      (validation as { ok?: boolean }).ok === false ? "Validation reported failure" : "Validation step completed"
    ],
    maxPayloadBytes: budget.maxChars,
    budgetSummary: summary
  });
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











































type RangeReadFacet = "values" | "formulas" | "numberFormat" | "text" | "style";

function rangeReadFacets(toolName: string): RangeReadFacet[] {
  switch (toolName) {
    case "excel.range.read_values":
      return ["values"];
    case "excel.range.read_formulas":
      return ["formulas"];
    case "excel.range.read_number_formats":
      return ["numberFormat"];
    case "excel.range.read_display_text":
      return ["text"];
    case "excel.range.read_styles":
      return ["style"];
    default:
      return ["values", "formulas", "numberFormat", "text", "style"];
  }
}

async function readRangeSnapshot(workbookId: string, sheetName: string, address: string, facets?: RangeReadFacet[]) {
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
  if (facets !== undefined) {
    operation.facets = facets;
  }
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

function tableReadRequest(args: {
  workbookId: string;
  tableName: string;
  includeValues?: boolean;
  includeFormulas?: boolean;
  includeText?: boolean;
  includeNumberFormats?: boolean;
  columns?: Array<string | number>;
  rowOffset?: number;
  rowLimit?: number;
}): TableReadRequest {
  const request: TableReadRequest = tableSelector(args);
  if (args.includeValues !== undefined) {
    request.includeValues = args.includeValues;
  }
  if (args.includeFormulas !== undefined) {
    request.includeFormulas = args.includeFormulas;
  }
  if (args.includeText !== undefined) {
    request.includeText = args.includeText;
  }
  if (args.includeNumberFormats !== undefined) {
    request.includeNumberFormats = args.includeNumberFormats;
  }
  if (args.columns !== undefined) {
    request.columns = args.columns;
  }
  if (args.rowOffset !== undefined) {
    request.rowOffset = args.rowOffset;
  }
  if (args.rowLimit !== undefined) {
    request.rowLimit = args.rowLimit;
  }
  return request;
}

type CompactDetailLevel = "summary" | "compact" | "full";

interface CompactTelemetryOptions {
  detailLevel: CompactDetailLevel;
  responseMode?: CompactResponseMode | undefined;
  truncated?: boolean;
  nextPage?: { rowOffset?: number; columnOffset?: number } | undefined;
  resourceUri?: string | undefined;
  resourceKind?: CompactResourceKind | undefined;
  resourceTitle?: string | undefined;
  resourcePayload?: unknown;
  resourceScope?: Record<string, unknown> | undefined;
  sourceHash?: string | undefined;
  nextActionRecommendation?: CompactNextActionRecommendation | undefined;
  reasoningHints?: string[] | undefined;
  confidence?: CompactConfidence | undefined;
  confidenceReasons?: string[] | undefined;
  storeResource?: boolean | undefined;
  maxPayloadBytes?: number | undefined;
  maxEstimatedTokens?: number | undefined;
  budgetSummary?: Record<string, unknown> | undefined;
}

interface ParsedCompactA1Address {
  startRow: number;
  startColumn: number;
  endRow: number;
  endColumn: number;
}

function withCompactTelemetry<T extends Record<string, unknown>>(payload: T, options: CompactTelemetryOptions): T & {
  payloadBytes: number;
  estimatedTokens: number;
  truncated: boolean;
  detailLevel: CompactDetailLevel;
  nextPage?: { rowOffset?: number; columnOffset?: number } | undefined;
  resourceUri?: string | undefined;
} {
  const originalPayloadBytes = Buffer.byteLength(JSON.stringify(payload), "utf8");
  const originalEstimatedTokens = Math.ceil(originalPayloadBytes / 4);
  const responseMode = compactResponseMode(options.responseMode);
  const shouldStoreForBrief = responseMode === "brief" && options.budgetSummary !== undefined && options.resourcePayload !== undefined;
  const maxPayloadBytes = options.maxPayloadBytes ?? (options.budgetSummary !== undefined ? COMPACT_DEFAULT_RESOURCE_THRESHOLD_BYTES : undefined);
  const overBudget =
    (maxPayloadBytes !== undefined && originalPayloadBytes > maxPayloadBytes) ||
    (options.maxEstimatedTokens !== undefined && originalEstimatedTokens > options.maxEstimatedTokens);
  const resourceMetadata: { title?: string; scope?: Record<string, unknown>; sourceHash?: string } = {};
  if (options.resourceTitle !== undefined) {
    resourceMetadata.title = options.resourceTitle;
  }
  if (options.resourceScope !== undefined) {
    resourceMetadata.scope = options.resourceScope;
  }
  if (options.sourceHash !== undefined) {
    resourceMetadata.sourceHash = options.sourceHash;
  }
  const beforeResourceCount = compactResources.size;
  const stored = options.storeResource || overBudget || shouldStoreForBrief
    ? storeCompactResource(options.resourceKind ?? "generic", options.resourcePayload ?? payload, resourceMetadata)
    : undefined;
  const cacheHit = stored !== undefined && compactResources.size === beforeResourceCount;
  const output = (overBudget || shouldStoreForBrief) && options.budgetSummary !== undefined
    ? {
        ...options.budgetSummary,
        budgetExceeded: overBudget,
        warnings: [
          ...asWarningArray(options.budgetSummary.warnings),
          ...(overBudget
            ? [{
                code: "COMPACT_BUDGET_EXCEEDED",
                message: "Full compact detail exceeded the response budget and was stored behind resourceUri."
              }]
            : [])
        ],
        resourcePayloadBytes: originalPayloadBytes,
        resourceEstimatedTokens: originalEstimatedTokens
      } as unknown as T
    : payload;
  const limitedOutput = enforceCompactOutputLimits(output as Record<string, unknown>, {
    originalPayloadBytes,
    originalEstimatedTokens,
    overBudget,
    ...(stored !== undefined ? { stored } : {})
  }) as T;
  const payloadBytes = Buffer.byteLength(JSON.stringify(limitedOutput), "utf8");
  compactToolResultCount += 1;
  compactToolResultBytes += payloadBytes;
  compactToolStoredBytes += stored?.payloadBytes ?? 0;
  if (stored !== undefined) {
    if (cacheHit) {
      compactCacheHitCount += 1;
    } else {
      compactCacheMissCount += 1;
    }
  }
  return {
    ...limitedOutput,
    payloadBytes,
    estimatedTokens: Math.ceil(payloadBytes / 4),
    truncated: options.truncated ?? false,
    detailLevel: options.detailLevel,
    responseMode,
    telemetry: compactTelemetrySummary(payloadBytes, stored, cacheHit),
    ...(options.nextActionRecommendation !== undefined ? { nextActionRecommendation: options.nextActionRecommendation } : {}),
    ...(options.reasoningHints !== undefined ? { reasoningHints: limitStringList(options.reasoningHints, COMPACT_LIMITS.maxWarnings) } : {}),
    ...(options.confidence !== undefined ? { confidence: options.confidence } : {}),
    ...(options.confidenceReasons !== undefined ? { confidenceReasons: limitStringList(options.confidenceReasons, COMPACT_LIMITS.maxWarnings) } : {}),
    ...(options.nextPage !== undefined ? { nextPage: options.nextPage } : {}),
    ...(stored !== undefined ? { resourceUri: stored.uri, contextId: stored.resourceId } : options.resourceUri !== undefined ? { resourceUri: options.resourceUri, contextId: compactResourceIdFromUri(options.resourceUri) } : {})
  };
}

function enforceCompactOutputLimits(
  output: Record<string, unknown>,
  options: {
    stored?: CompactStoredResource;
    originalPayloadBytes: number;
    originalEstimatedTokens: number;
    overBudget: boolean;
  }
): Record<string, unknown> {
  const compacted = { ...output };
  const omittedCounts: Record<string, number> = {};
  if (typeof compacted.summary === "string" && compacted.summary.length > COMPACT_LIMITS.maxSummaryChars) {
    omittedCounts.summaryChars = compacted.summary.length - COMPACT_LIMITS.maxSummaryChars;
    compacted.summary = `${compacted.summary.slice(0, COMPACT_LIMITS.maxSummaryChars)}...`;
  }
  for (const key of ["examples", "issues", "warnings", "changedRanges", "affectedRanges", "sampleRows", "sampleColumns"]) {
    const value = compacted[key];
    if (!Array.isArray(value)) {
      continue;
    }
    const limit = compactArrayLimit(key);
    if (value.length > limit) {
      omittedCounts[key] = value.length - limit;
      compacted[key] = value.slice(0, limit);
      compacted[`${key}Truncated`] = true;
    }
  }
  let serialized = JSON.stringify(compacted);
  if (serialized.length > COMPACT_LIMITS.maxToolResultChars && options.stored !== undefined) {
    const keepKeys = [
      "ok",
      "toolName",
      "workbookId",
      "sheetName",
      "address",
      "tableName",
      "validator",
      "summary",
      "source",
      "window",
      "shape",
      "issueCount",
      "severityCounts",
      "categories",
      "changedRangeCount",
      "cellsChanged",
      "formulasChanged",
      "stylesChanged",
      "tablesChanged",
      "sheetsChanged",
      "destructiveLevel",
      "compactProof",
      "validationSummary",
      "diffSummary",
      "transactionId",
      "backupId",
      "rollbackAvailable"
    ];
    const reduced: Record<string, unknown> = {};
    for (const key of keepKeys) {
      if (compacted[key] !== undefined) {
        reduced[key] = compacted[key];
      }
    }
    serialized = JSON.stringify(reduced);
    Object.assign(compacted, reduced);
    for (const key of Object.keys(compacted)) {
      if (!(key in reduced)) {
        delete compacted[key];
      }
    }
    compacted.budgetExceeded = true;
  }
  if (Object.keys(omittedCounts).length > 0) {
    compacted.omittedCounts = omittedCounts;
    compacted.truncated = true;
  }
  if (options.stored !== undefined) {
    compacted.fullResult = {
      contextId: options.stored.resourceId,
      resourceUri: options.stored.uri
    };
    compacted.resourcePayloadBytes = options.originalPayloadBytes;
    compacted.resourceEstimatedTokens = options.originalEstimatedTokens;
    compacted.estimatedTokensSaved = Math.max(0, options.originalEstimatedTokens - Math.ceil(Buffer.byteLength(JSON.stringify(compacted), "utf8") / 4));
  }
  if (options.overBudget) {
    compacted.budgetExceeded = true;
  }
  return compacted;
}

function compactArrayLimit(key: string): number {
  if (key === "warnings") {
    return COMPACT_LIMITS.maxWarnings;
  }
  if (key === "issues") {
    return COMPACT_LIMITS.maxValidationIssues;
  }
  if (key === "sampleRows") {
    return COMPACT_LIMITS.maxSampleRows;
  }
  if (key === "sampleColumns") {
    return COMPACT_LIMITS.maxSampleCols;
  }
  return COMPACT_LIMITS.maxExamples;
}

function compactTelemetrySummary(payloadBytes: number, stored?: CompactStoredResource, cacheHit = false) {
  return {
    responseBytes: payloadBytes,
    estimatedResponseTokens: Math.ceil(payloadBytes / 4),
    storedPayloadBytes: stored?.payloadBytes ?? 0,
    estimatedStoredTokens: stored?.estimatedTokens ?? 0,
    estimatedTokensSaved: stored ? Math.max(0, stored.estimatedTokens - Math.ceil(payloadBytes / 4)) : 0,
    cacheHit
  };
}

function limitStringList(values: string[], limit: number): string[] {
  return values.filter((value) => typeof value === "string" && value.length > 0).slice(0, limit);
}

function compactResponseMode(requested?: CompactResponseMode): CompactResponseMode {
  if (requested !== undefined) {
    return requested;
  }
  return "brief";
}

function compactBudgetSchema() {
  return z.object({
    maxRows: z.number().int().min(0).optional(),
    maxCols: z.number().int().min(0).optional(),
    maxChars: z.number().int().min(0).optional(),
    maxIssues: z.number().int().min(0).optional(),
    maxExamples: z.number().int().min(0).optional(),
    maxWarnings: z.number().int().min(0).optional()
  });
}

function compactRequestedBudget(args: { budget?: Record<string, unknown>; maxRows?: number; maxColumns?: number; maxPayloadBytes?: number; maxIssues?: number }) {
  const budget = args.budget && typeof args.budget === "object" ? args.budget : {};
  const requestedRows = typeof budget.maxRows === "number" ? budget.maxRows : args.maxRows;
  const requestedColumns = typeof budget.maxCols === "number" ? budget.maxCols : args.maxColumns;
  const maxRows = compactClampNumber(budget.maxRows, args.maxRows, Math.max(COMPACT_LIMITS.maxSampleRows, requestedRows ?? 0));
  const maxColumns = compactClampNumber(budget.maxCols, args.maxColumns, Math.max(COMPACT_LIMITS.maxSampleCols, requestedColumns ?? 0));
  const maxChars = compactClampNumber(budget.maxChars, args.maxPayloadBytes, COMPACT_LIMITS.maxToolResultChars);
  const maxIssues = compactClampNumber(budget.maxIssues, args.maxIssues, COMPACT_LIMITS.maxValidationIssues);
  const maxExamples = compactClampNumber(budget.maxExamples, undefined, COMPACT_LIMITS.maxExamples);
  const maxWarnings = compactClampNumber(budget.maxWarnings, undefined, COMPACT_LIMITS.maxWarnings);
  return {
    maxRows,
    maxColumns,
    maxChars,
    maxIssues,
    maxExamples,
    maxWarnings,
    requested: budget,
    applied: { maxRows, maxColumns, maxChars, maxIssues, maxExamples, maxWarnings }
  };
}

function compactClampNumber(primary: unknown, fallback: number | undefined, max: number): number {
  const value = typeof primary === "number" ? primary : fallback;
  if (value === undefined) {
    return max;
  }
  return Math.max(0, Math.min(value, max));
}

function asWarningArray(value: unknown): Array<Record<string, unknown>> {
  return Array.isArray(value) ? value.filter((item): item is Record<string, unknown> => typeof item === "object" && item !== null) : [];
}

function compactSourceHash(value: unknown): string {
  return createHash("sha256").update(JSON.stringify(value)).digest("hex").slice(0, 16);
}

function compactReadScope(input: Record<string, unknown>): Record<string, unknown> {
  return Object.fromEntries(
    Object.entries(input).filter(([, value]) => value !== undefined)
  );
}

function storeCompactResource(
  kind: CompactResourceKind,
  payload: unknown,
  options: string | { title?: string; scope?: Record<string, unknown>; sourceHash?: string; pinned?: boolean } = {}
): CompactStoredResource {
  const resourceId = makeId<string>("compact");
  const payloadBytes = Buffer.byteLength(JSON.stringify(payload), "utf8");
  const metadata = typeof options === "string" ? { title: options } : options;
  const existing = findExistingCompactResource(kind, metadata.scope, metadata.sourceHash);
  if (existing) {
    existing.lastAccessedAt = new Date().toISOString();
    return existing;
  }
  const now = new Date().toISOString();
  const resource: CompactStoredResource = {
    resourceId,
    uri: `excel://compact/${resourceId}`,
    kind,
    title: metadata.title,
    scope: metadata.scope,
    sourceHash: metadata.sourceHash,
    pinned: metadata.pinned,
    createdAt: now,
    lastAccessedAt: now,
    payloadBytes,
    estimatedTokens: Math.ceil(payloadBytes / 4),
    accessCount: 0,
    payload
  };
  compactResources.set(resourceId, resource);
  gcCompactResources({ keepNewest: COMPACT_RESOURCE_LIMIT, maxAgeMs: COMPACT_RESOURCE_TTL_MS });
  return resource;
}

function findExistingCompactResource(kind: CompactResourceKind, scope?: Record<string, unknown>, sourceHash?: string): CompactStoredResource | undefined {
  if (!sourceHash) {
    return undefined;
  }
  const scopeHash = scope ? compactSourceHash(scope) : undefined;
  for (const resource of compactResources.values()) {
    if (resource.kind !== kind || resource.sourceHash !== sourceHash) {
      continue;
    }
    const resourceScopeHash = resource.scope ? compactSourceHash(resource.scope) : undefined;
    if (resourceScopeHash === scopeHash) {
      return resource;
    }
  }
  return undefined;
}

type CompactResourceReadMode = "metadata" | "preview" | "page" | "full";

function getCompactResource(
  resourceId: string,
  options: {
    mode?: CompactResourceReadMode;
    includePayload?: boolean;
    maxPayloadBytes?: number;
    maxEstimatedTokens?: number;
    offset?: number;
    limit?: number;
  } = {}
) {
  const resource = compactResources.get(compactResourceIdFromUri(resourceId));
  if (!resource) {
    return { ok: false, error: { code: "COMPACT_RESOURCE_NOT_FOUND", message: "Compact detail resource was not found or has expired." } };
  }
  resource.lastAccessedAt = new Date().toISOString();
  resource.accessCount += 1;
  const { payload, ...summary } = resource;
  const mode = options.mode ?? (options.includePayload === true ? "full" : "metadata");
  if (mode === "metadata") {
    return {
      ok: true,
      ...summary,
      contextId: summary.resourceId,
      payloadAvailable: true,
      payloadIncluded: false,
      message: "Use mode=preview or mode=page for bounded inspection, or mode=full/includePayload=true for full detail."
    };
  }
  if (mode === "preview" || mode === "page") {
    const serialized = JSON.stringify(payload, null, 2);
    const defaultLimit = mode === "preview" ? 4000 : COMPACT_DEFAULT_RESOURCE_THRESHOLD_BYTES;
    const limit = Math.max(1, Math.min(options.limit ?? defaultLimit, COMPACT_DEFAULT_RESOURCE_THRESHOLD_BYTES));
    const offset = Math.max(0, Math.min(options.offset ?? 0, serialized.length));
    const text = serialized.slice(offset, offset + limit);
    const nextOffset = offset + text.length < serialized.length ? offset + text.length : undefined;
    return {
      ok: true,
      ...summary,
      contextId: summary.resourceId,
      payloadAvailable: true,
      payloadIncluded: false,
      mode,
      format: "json-text",
      offset,
      limit,
      text,
      textBytes: Buffer.byteLength(text, "utf8"),
      totalCharacters: serialized.length,
      totalBytes: Buffer.byteLength(serialized, "utf8"),
      ...(nextOffset !== undefined ? { nextOffset } : {})
    };
  }
  const overBudget =
    (options.maxPayloadBytes !== undefined && resource.payloadBytes > options.maxPayloadBytes) ||
    (options.maxEstimatedTokens !== undefined && resource.estimatedTokens > options.maxEstimatedTokens);
  if (overBudget) {
    return {
      ok: true,
      ...summary,
      contextId: summary.resourceId,
      payloadAvailable: true,
      payloadIncluded: false,
      budgetExceeded: true,
      warnings: [
        {
          code: "COMPACT_RESOURCE_BUDGET_EXCEEDED",
          message: "Stored detail exceeded the requested response budget. Raise the budget or inspect a smaller resource."
        }
      ]
    };
  }
  return { ok: true, ...summary, contextId: summary.resourceId, payloadAvailable: true, payloadIncluded: true, payload };
}

function compactResourceReadOptions(options: {
  mode?: CompactResourceReadMode | undefined;
  includePayload?: boolean | undefined;
  maxPayloadBytes?: number | undefined;
  maxEstimatedTokens?: number | undefined;
  offset?: number | undefined;
  limit?: number | undefined;
}) {
  const normalized: { mode?: CompactResourceReadMode; includePayload?: boolean; maxPayloadBytes?: number; maxEstimatedTokens?: number; offset?: number; limit?: number } = {};
  if (options.mode !== undefined) {
    normalized.mode = options.mode;
  }
  if (options.includePayload !== undefined) {
    normalized.includePayload = options.includePayload;
  }
  if (options.maxPayloadBytes !== undefined) {
    normalized.maxPayloadBytes = options.maxPayloadBytes;
  }
  if (options.maxEstimatedTokens !== undefined) {
    normalized.maxEstimatedTokens = options.maxEstimatedTokens;
  }
  if (options.offset !== undefined) {
    normalized.offset = options.offset;
  }
  if (options.limit !== undefined) {
    normalized.limit = options.limit;
  }
  return normalized;
}

function listCompactResources(options: { kind?: string; workbook?: string; worksheet?: string; olderThanSeconds?: number; includePinned?: boolean; limit?: number } = {}) {
  const now = Date.now();
  const resources = [...compactResources.values()]
    .filter((resource) => options.kind === undefined || resource.kind === options.kind)
    .filter((resource) => options.includePinned === true || resource.pinned !== true)
    .filter((resource) => options.workbook === undefined || resource.scope?.workbookId === options.workbook || resource.scope?.workbook === options.workbook)
    .filter((resource) => options.worksheet === undefined || resource.scope?.sheetName === options.worksheet || resource.scope?.worksheet === options.worksheet)
    .filter((resource) => options.olderThanSeconds === undefined || now - Date.parse(resource.lastAccessedAt) > options.olderThanSeconds * 1000)
    .sort((left, right) => Date.parse(right.lastAccessedAt) - Date.parse(left.lastAccessedAt))
    .slice(0, options.limit ?? COMPACT_RESOURCE_LIMIT)
    .map(({ payload, ...summary }) => ({ ...summary, contextId: summary.resourceId }));
  const totalPayloadBytes = resources.reduce((sum, resource) => sum + resource.payloadBytes, 0);
  return { ok: true, count: resources.length, totalPayloadBytes, totalEstimatedTokens: Math.ceil(totalPayloadBytes / 4), resources };
}

function compactContextStats() {
  const resources = [...compactResources.values()];
  const totalPayloadBytes = resources.reduce((sum, resource) => sum + resource.payloadBytes, 0);
  const totalEstimatedTokens = resources.reduce((sum, resource) => sum + resource.estimatedTokens, 0);
  const largestResources = resources
    .sort((left, right) => right.payloadBytes - left.payloadBytes)
    .slice(0, 10)
    .map(({ payload, ...summary }) => ({ ...summary, contextId: summary.resourceId }));
  const cacheAttempts = compactCacheHitCount + compactCacheMissCount;
  return {
    ok: true,
    resourcesCount: resources.length,
    totalPayloadBytes,
    estimatedStoredTokens: totalEstimatedTokens,
    largestResources,
    toolResults: {
      count: compactToolResultCount,
      responseBytes: compactToolResultBytes,
      estimatedResponseTokens: Math.ceil(compactToolResultBytes / 4),
      storedPayloadBytes: compactToolStoredBytes,
      estimatedStoredTokens: Math.ceil(compactToolStoredBytes / 4),
      estimatedTokensSaved: Math.max(0, Math.ceil(compactToolStoredBytes / 4) - Math.ceil(compactToolResultBytes / 4))
    },
    cache: {
      cacheEntries: compactCache.size,
      cacheHits: compactCacheHitCount,
      cacheMisses: compactCacheMissCount,
      cacheHitRate: cacheAttempts === 0 ? 0 : Number((compactCacheHitCount / cacheAttempts).toFixed(4)),
      invalidationCount: compactCacheInvalidationCount,
      lastInvalidatedAt: compactCacheLastInvalidatedAt
    },
    idempotency: {
      records: compactIdempotencyRecords.size,
      ttlMs: COMPACT_IDEMPOTENCY_TTL_MS
    }
  };
}

function deleteCompactResource(resourceId: string) {
  const normalized = compactResourceIdFromUri(resourceId);
  return { ok: compactResources.delete(normalized), resourceId: normalized };
}

function clearCompactResources(kind?: string) {
  let deleted = 0;
  for (const [resourceId, resource] of compactResources) {
    if (kind === undefined || resource.kind === kind) {
      compactResources.delete(resourceId);
      deleted += 1;
    }
  }
  return { ok: true, deleted };
}

function gcCompactResources(options: { kind?: string; maxAgeMs?: number; keepNewest?: number; includePinned?: boolean } = {}) {
  const now = Date.now();
  const maxAgeMs = options.maxAgeMs ?? COMPACT_RESOURCE_TTL_MS;
  const keepNewest = options.keepNewest ?? COMPACT_RESOURCE_LIMIT;
  const eligible = [...compactResources.entries()]
    .filter(([, resource]) => options.kind === undefined || resource.kind === options.kind)
    .filter(([, resource]) => options.includePinned === true || resource.pinned !== true);
  const sortedByAccess = eligible.sort(([, left], [, right]) => Date.parse(right.lastAccessedAt) - Date.parse(left.lastAccessedAt));
  const keep = new Set(sortedByAccess.slice(0, keepNewest).map(([resourceId]) => resourceId));
  let deleted = 0;
  let expired = 0;
  let overflow = 0;
  for (const [resourceId, resource] of sortedByAccess) {
    const ageMs = now - Date.parse(resource.lastAccessedAt);
    const shouldDeleteExpired = maxAgeMs >= 0 && ageMs > maxAgeMs;
    const shouldDeleteOverflow = !keep.has(resourceId);
    if (shouldDeleteExpired || shouldDeleteOverflow) {
      compactResources.delete(resourceId);
      deleted += 1;
      if (shouldDeleteExpired) {
        expired += 1;
      } else {
        overflow += 1;
      }
    }
  }
  return { ok: true, deleted, expired, overflow, remaining: compactResources.size, maxAgeMs, keepNewest };
}

function compactResourceIdFromUri(value: string): string {
  return value.startsWith("excel://compact/") ? value.slice("excel://compact/".length) : value;
}

async function compactCacheValue<T>(key: string, producer: () => Promise<T> | T): Promise<T> {
  await invalidateCompactCacheForWorkbookEvents();
  const cached = compactCache.get(key);
  if (cached !== undefined) {
    compactCacheHitCount += 1;
    return cached.value as T;
  }
  const value = await producer();
  compactCache.set(key, { key, createdAt: new Date().toISOString(), value });
  compactCacheMissCount += 1;
  return value;
}

async function invalidateCompactCacheForWorkbookEvents(): Promise<void> {
  try {
    const recent = await runtime.getRecentEvents(1) as { events?: Array<{ eventId?: string; method?: string }> };
    const event = recent.events?.[0];
    if (!event?.eventId || event.eventId === compactLastObservedEventId) {
      return;
    }
    compactLastObservedEventId = event.eventId;
    if (event.method !== "addin.heartbeat") {
      clearCompactCache(`event:${event.method ?? "unknown"}`);
    }
  } catch {
    // Compact caching must never make read-only workbook discovery fail.
  }
}

function getCompactCacheStatus() {
  return {
    ok: true,
    size: compactCache.size,
    invalidationCount: compactCacheInvalidationCount,
    lastInvalidatedAt: compactCacheLastInvalidatedAt,
    keys: [...compactCache.keys()]
  };
}

function clearCompactCache(reason: string) {
  const cleared = compactCache.size;
  compactCache.clear();
  compactCacheInvalidationCount += 1;
  compactCacheLastInvalidatedAt = new Date().toISOString();
  return { ok: true, cleared, reason, invalidationCount: compactCacheInvalidationCount, invalidatedAt: compactCacheLastInvalidatedAt };
}

async function resolveAgentWorkbookContext(workbookContextId: string): Promise<{ ok: true; context: AgentWorkbookContextResource } | { ok: false; response: Record<string, unknown> }> {
  const current = runtime.getAgentContextResource(workbookContextId) as AgentWorkbookContextResource | { ok: false; error?: unknown };
  if (!current.ok) {
    return {
      ok: false,
      response: {
        ok: false,
        status: "NOT_FOUND",
        workbookContextId,
        summary: "Workbook context was not found or has expired.",
        nextAction: "call_prepare"
      }
    };
  }
  const refreshed = await runtime.runAgent({ request: "Refresh workbook context", mode: "prepare", workbookContextId });
  const refreshedId = refreshed.workbookContextId ? String(refreshed.workbookContextId) : workbookContextId;
  const resource = runtime.getAgentContextResource(refreshedId) as AgentWorkbookContextResource | { ok: false; error?: unknown };
  if (!resource.ok) {
    return {
      ok: false,
      response: {
        ok: false,
        status: refreshed.status === "ERROR" ? "STALE_CONTEXT" : "NOT_FOUND",
        workbookContextId,
        summary: refreshed.summary,
        nextAction: refreshed.nextAction,
        warnings: refreshed.warnings
      }
    };
  }
  return { ok: true, context: resource };
}

async function resolveContextSheet(workbookContextId: string, sheetName?: string): Promise<AgentContextResolution<AgentWorkbookContextResource["sheets"][number]>> {
  const context = await resolveAgentWorkbookContext(workbookContextId);
  if (!context.ok) {
    return context;
  }
  const sheets = context.context.sheets;
  if (sheetName) {
    const exact = sheets.find((sheet) => sheet.name === sheetName);
    if (exact) {
      return { ok: true, context: context.context, value: exact };
    }
    const normalized = normalizeContextLookup(sheetName);
    const candidates = sheets
      .map((sheet) => ({ sheet, score: contextLookupScore(normalized, sheet.name) }))
      .filter((candidate) => candidate.score > 0)
      .sort((left, right) => right.score - left.score);
    const best = candidates[0];
    if (best && (candidates.length === 1 || (best.score >= 0.86 && best.score - (candidates[1]?.score ?? 0) >= 0.12))) {
      return { ok: true, context: context.context, value: best.sheet };
    }
    return {
      ok: false,
      response: {
        ok: false,
        status: candidates.length > 0 ? "AMBIGUOUS_TARGET" : "NOT_FOUND",
        workbookContextId,
        summary: candidates.length > 0 ? "Multiple sheets match the requested sheet name." : "Requested sheet was not found in workbook context.",
        candidates: (candidates.length > 0 ? candidates.map((candidate) => candidate.sheet) : sheets).slice(0, 10).map((sheet) => ({
          name: sheet.name,
          kind: sheet.kind,
          usedRange: contextUsedRange(sheet)
        })),
        nextAction: "call_with_target"
      }
    };
  }
  const activeSheet = context.context.workbook.activeSheet ? sheets.find((sheet) => sheet.name === context.context.workbook.activeSheet) : undefined;
  if (activeSheet) {
    return { ok: true, context: context.context, value: activeSheet };
  }
  if (sheets.length === 1 && sheets[0]) {
    return { ok: true, context: context.context, value: sheets[0] };
  }
  return {
    ok: false,
    response: {
      ok: false,
      status: "NEEDS_SHEET",
      workbookContextId,
      summary: "A sheet is required because the workbook context has multiple sheets and no active sheet.",
      candidates: sheets.slice(0, 10).map((sheet) => ({ name: sheet.name, kind: sheet.kind, usedRange: contextUsedRange(sheet) })),
      nextAction: "call_with_target"
    }
  };
}

async function normalizeContextRangeReadArgs(args: ContextRangeCompactReadRequest): Promise<AgentContextResolution<RangeCompactReadRequest>> {
  if (!args.workbookContextId) {
    if (!args.workbookId || !args.sheetName) {
      return {
        ok: false,
        response: {
          ok: false,
          status: "NEEDS_INPUT",
          summary: "Range compact read requires workbookId and sheetName unless workbookContextId is supplied.",
          nextAction: "call_with_target"
        }
      };
    }
    return {
      ok: true,
      context: undefined as never,
      value: {
        ...args,
        workbookId: args.workbookId as WorkbookId,
        sheetName: args.sheetName
      } as RangeCompactReadRequest
    };
  }
  const resolved = await resolveContextSheet(args.workbookContextId, args.sheetName);
  if (!resolved.ok) {
    return resolved;
  }
  const workbookId = resolved.context.workbook.workbookId;
  if (!workbookId) {
    return {
      ok: false,
      response: {
        ok: false,
        status: "NEEDS_INPUT",
        workbookContextId: args.workbookContextId,
        summary: "Workbook context does not include a workbookId for range reads.",
        nextAction: "call_prepare"
      }
    };
  }
  return {
    ok: true,
    context: resolved.context,
    value: {
      ...args,
      workbookId: workbookId as WorkbookId,
      sheetName: resolved.value.name
    } as RangeCompactReadRequest
  };
}

async function normalizeContextTableReadArgs(args: ContextTableCompactReadRequest): Promise<AgentContextResolution<TableCompactReadRequest>> {
  if (!args.workbookContextId) {
    if (!args.workbookId || !args.tableName) {
      return {
        ok: false,
        response: {
          ok: false,
          status: "NEEDS_INPUT",
          summary: "Table compact read requires workbookId and tableName unless workbookContextId is supplied.",
          nextAction: "call_with_target"
        }
      };
    }
    return {
      ok: true,
      context: undefined as never,
      value: {
        ...args,
        workbookId: args.workbookId as WorkbookId,
        tableName: args.tableName
      } as TableCompactReadRequest
    };
  }
  const context = await resolveAgentWorkbookContext(args.workbookContextId);
  if (!context.ok) {
    return context;
  }
  const workbookId = context.context.workbook.workbookId;
  if (!workbookId) {
    return {
      ok: false,
      response: {
        ok: false,
        status: "NEEDS_INPUT",
        workbookContextId: args.workbookContextId,
        summary: "Workbook context does not include a workbookId for table reads.",
        nextAction: "call_prepare"
      }
    };
  }
  const table = resolveContextTable(context.context, args.tableName);
  if (!table.ok) {
    return { ok: false, response: table.response };
  }
  return {
    ok: true,
    context: context.context,
    value: {
      ...args,
      workbookId: workbookId as WorkbookId,
      tableName: table.table.name ?? table.table.range
    } as TableCompactReadRequest
  };
}

function resolveContextTable(context: AgentWorkbookContextResource, tableName?: string): { ok: true; table: AgentWorkbookContextResource["tables"][number] } | { ok: false; response: Record<string, unknown> } {
  if (!tableName && context.tables.length === 1 && context.tables[0]) {
    return { ok: true, table: context.tables[0] };
  }
  if (tableName) {
    const exact = context.tables.find((table) => table.name === tableName);
    if (exact) {
      return { ok: true, table: exact };
    }
    const normalized = normalizeContextLookup(tableName);
    const candidates = context.tables
      .map((table) => ({ table, score: Math.max(contextLookupScore(normalized, table.name ?? ""), contextLookupScore(normalized, table.sheetName)) }))
      .filter((candidate) => candidate.score > 0)
      .sort((left, right) => right.score - left.score);
    const best = candidates[0];
    if (best && (candidates.length === 1 || (best.score >= 0.86 && best.score - (candidates[1]?.score ?? 0) >= 0.12))) {
      return { ok: true, table: best.table };
    }
    return {
      ok: false,
      response: {
        ok: false,
        status: candidates.length > 0 ? "AMBIGUOUS_TARGET" : "NOT_FOUND",
        workbookContextId: context.workbookContextId,
        summary: candidates.length > 0 ? "Multiple tables match the requested table name." : "Requested table was not found in workbook context.",
        candidates: (candidates.length > 0 ? candidates.map((candidate) => candidate.table) : context.tables).slice(0, 10).map(contextTableCandidate),
        nextAction: "call_with_target"
      }
    };
  }
  return {
    ok: false,
    response: {
      ok: false,
      status: "NEEDS_INPUT",
      workbookContextId: context.workbookContextId,
      summary: "A table name is required because the workbook context has multiple or no tables.",
      candidates: context.tables.slice(0, 10).map(contextTableCandidate),
      nextAction: "call_with_target"
    }
  };
}

async function compactResourceResult(args: {
  resourceId?: string;
  resourceUri?: string;
  workbookContextId?: string;
  mode?: CompactResourceReadMode;
  includePayload?: boolean;
  maxPayloadBytes?: number;
  maxEstimatedTokens?: number;
  offset?: number;
  limit?: number;
}) {
  if (args.workbookContextId) {
    const context = await resolveAgentWorkbookContext(args.workbookContextId);
    if (!context.ok) {
      return context.response;
    }
  }
  return {
    ...getCompactResource(args.resourceId ?? compactResourceIdFromUri(args.resourceUri ?? ""), compactResourceReadOptions(args)),
    ...(args.workbookContextId ? { workbookContextId: args.workbookContextId } : {})
  };
}

function contextTableCandidate(table: AgentWorkbookContextResource["tables"][number]) {
  return { name: table.name, sheetName: table.sheetName, range: table.range };
}

function contextUsedRange(sheet: AgentWorkbookContextResource["sheets"][number]) {
  if (!sheet.usedRange) {
    return undefined;
  }
  return {
    address: sheet.usedRange,
    ...(sheet.rowCount !== undefined ? { rowCount: sheet.rowCount } : {}),
    ...(sheet.columnCount !== undefined ? { columnCount: sheet.columnCount } : {})
  };
}

function normalizeContextLookup(value: string): string[] {
  return value.toLowerCase().replace(/[^a-z0-9]+/g, " ").trim().split(/\s+/).filter(Boolean);
}

function contextLookupScore(queryTokens: string[], value: string): number {
  const tokens = normalizeContextLookup(value);
  if (queryTokens.length === 0 || tokens.length === 0) {
    return 0;
  }
  const query = queryTokens.join(" ");
  const candidate = tokens.join(" ");
  if (query === candidate) {
    return 1;
  }
  if (query.includes(candidate) || candidate.includes(query)) {
    return 0.86;
  }
  const overlap = tokens.filter((token) => queryTokens.includes(token)).length;
  return overlap === 0 ? 0 : 0.25 + (overlap / tokens.length) * 0.45 + (overlap / queryTokens.length) * 0.25;
}

async function workbookSummary(workbookId?: WorkbookId, workbookContextId?: string) {
  if (workbookContextId) {
    const context = await resolveAgentWorkbookContext(workbookContextId);
    if (!context.ok) {
      return context.response;
    }
    const metadata = context.context;
    const tableNames = metadata.tables.map((table) => table.name ?? `${table.sheetName}!${table.range}`);
    const usedRangeCells = metadata.sheets.reduce((sum, sheet) => sum + usedRangeCellCount(contextUsedRange(sheet)), 0);
    return withCompactTelemetry(
      {
        ok: true,
        workbookContextId: metadata.workbookContextId,
        workbook: metadata.workbook,
        workbookId: metadata.workbook.workbookId,
        sheetCount: metadata.sheets.length,
        tableCount: tableNames.length,
        usedRangeCells,
        sheets: metadata.sheets.map((sheet) => ({
          name: sheet.name,
          kind: sheet.kind,
          usedRange: contextUsedRange(sheet),
          rowCount: sheet.rowCount,
          columnCount: sheet.columnCount,
          headerCount: Array.isArray(sheet.headers) ? sheet.headers.length : 0,
          tables: metadata.tables.filter((table) => table.sheetName === sheet.name).map((table) => table.name ?? `${table.sheetName}!${table.range}`)
        }))
      },
      { detailLevel: "summary" }
    );
  }
  return compactCacheValue(`workbookSummary:${workbookId ?? "active"}`, async () => {
  const result = await runtime.getWorkbookMap();
  const map = "map" in result ? result.map as { workbook?: any; sheets?: any[] } : undefined;
  if (!result.ok || !map) {
    return withCompactTelemetry({ ok: false, workbookId, source: result }, { detailLevel: "summary" });
  }
  const sheets = map.sheets ?? [];
  const tableNames = sheets.flatMap((sheet) => (sheet.tables ?? []).map((table: { name: string }) => table.name));
  const usedRangeCells = sheets.reduce((sum, sheet) => sum + usedRangeCellCount(sheet.usedRange), 0);
  return withCompactTelemetry(
    {
      ok: true,
      workbook: map.workbook,
      workbookId: workbookId ?? map.workbook?.workbookId,
      sheetCount: sheets.length,
      tableCount: tableNames.length,
      usedRangeCells,
      sheets: sheets.map((sheet) => ({
        name: sheet.name,
        position: sheet.position,
        visibility: sheet.visibility,
        usedRange: sheet.usedRange,
        tableCount: sheet.tables?.length ?? 0,
        tables: (sheet.tables ?? []).map((table: { name: string }) => table.name)
      }))
    },
    { detailLevel: "summary" }
  );
  });
}

async function workbookUsedRangeSummary(workbookId?: WorkbookId) {
  return compactCacheValue(`workbookUsedRangeSummary:${workbookId ?? "active"}`, async () => {
  const result = await runtime.getWorkbookMap();
  const map = "map" in result ? result.map as { workbook?: any; sheets?: any[] } : undefined;
  if (!result.ok || !map) {
    return withCompactTelemetry({ ok: false, workbookId, source: result }, { detailLevel: "summary" });
  }
  const sheets = map.sheets ?? [];
  return withCompactTelemetry(
    {
      ok: true,
      workbookId: workbookId ?? map.workbook?.workbookId,
      usedRanges: sheets.map((sheet) => ({
        sheetName: sheet.name,
        usedRange: sheet.usedRange,
        cellCount: usedRangeCellCount(sheet.usedRange)
      })),
      totalCells: sheets.reduce((sum, sheet) => sum + usedRangeCellCount(sheet.usedRange), 0)
    },
    { detailLevel: "summary" }
  );
  });
}

async function sheetSummary(sheetName?: string, workbookId?: WorkbookId, workbookContextId?: string) {
  if (workbookContextId) {
    const resolved = await resolveContextSheet(workbookContextId, sheetName);
    if (!resolved.ok) {
      return resolved.response;
    }
    const { context, value: sheet } = resolved;
    const usedRange = contextUsedRange(sheet);
    return withCompactTelemetry(
      {
        ok: true,
        workbookContextId: context.workbookContextId,
        workbookId: context.workbook.workbookId,
        sheetName: sheet.name,
        kind: sheet.kind,
        usedRange,
        rowCount: sheet.rowCount,
        columnCount: sheet.columnCount,
        cellCount: usedRangeCellCount(usedRange),
        headerCount: Array.isArray(sheet.headers) ? sheet.headers.length : 0,
        tableCount: context.tables.filter((table) => table.sheetName === sheet.name).length,
        tables: context.tables.filter((table) => table.sheetName === sheet.name).map((table) => table.name ?? `${table.sheetName}!${table.range}`),
        sheet
      },
      { detailLevel: "summary" }
    );
  }
  if (!sheetName) {
    return withCompactTelemetry({ ok: false, status: "NEEDS_SHEET", summary: "sheetName is required when workbookContextId is not supplied." }, { detailLevel: "summary" });
  }
  return compactCacheValue(`sheetSummary:${workbookId ?? "active"}:${sheetName}`, async () => {
  const info = await selectSheetInfo(sheetName);
  const workbook = (info.result as { map?: { workbook?: any } }).map?.workbook;
  return withCompactTelemetry(
    {
      ok: info.ok,
      workbookId: workbookId ?? workbook?.workbookId,
      sheetName,
      usedRange: info.sheet?.usedRange,
      cellCount: usedRangeCellCount(info.sheet?.usedRange),
      tableCount: info.sheet?.tables?.length ?? 0,
      tables: (info.sheet?.tables ?? []).map((table: { name: string }) => table.name),
      sheet: info.sheet
    },
    { detailLevel: "summary" }
  );
  });
}

function rangeSummary(workbookId: WorkbookId, sheetName: string, address: string) {
  const parsed = parseCompactA1Address(address);
  const rowCount = parsed.endRow - parsed.startRow + 1;
  const columnCount = parsed.endColumn - parsed.startColumn + 1;
  const defaultRows = Math.min(rowCount, 50);
  const defaultColumns = Math.min(columnCount, 25);
  return withCompactTelemetry(
    {
      ok: true,
      workbookId,
      sheetName,
      address,
      rowCount,
      columnCount,
      cellCount: rowCount * columnCount,
      defaultCompactWindow: {
        address: compactWindowAddress(address, 0, 0, defaultRows, defaultColumns),
        rowCount: defaultRows,
        columnCount: defaultColumns,
        cellCount: defaultRows * defaultColumns
      }
    },
    { detailLevel: "summary", truncated: rowCount > defaultRows || columnCount > defaultColumns }
  );
}

async function compactRangeRead(args: ContextRangeCompactReadRequest) {
  const normalized = await normalizeContextRangeReadArgs(args);
  if (!normalized.ok) {
    return normalized.response;
  }
  const request = normalized.value as RangeCompactReadRequest & { cellOutput?: CompactCellOutput };
  const mode = request.mode ?? "window";
  const responseMode = compactResponseMode(request.responseMode);
  const budget = compactRequestedBudget(request as RangeCompactReadRequest & { budget?: Record<string, unknown> });
  const parsed = parseCompactA1Address(request.address);
  const sourceRowCount = parsed.endRow - parsed.startRow + 1;
  const sourceColumnCount = parsed.endColumn - parsed.startColumn + 1;
  const rowOffset = Math.min(request.rowOffset ?? 0, sourceRowCount);
  const columnOffset = Math.min(request.columnOffset ?? 0, sourceColumnCount);
  const maxRows = budget.maxRows;
  const maxColumns = budget.maxColumns;
  const availableRows = Math.max(0, sourceRowCount - rowOffset);
  const availableColumns = Math.max(0, sourceColumnCount - columnOffset);
  const columnLimit = Math.min(maxColumns, availableColumns);
  const maxRowsByCells = request.maxCells !== undefined && columnLimit > 0 ? Math.floor(request.maxCells / columnLimit) : maxRows;
  const rowLimit = Math.min(maxRows, maxRowsByCells, availableRows);
  const windowAddress = rowLimit > 0 && columnLimit > 0
    ? compactWindowAddress(request.address, rowOffset, columnOffset, rowLimit, columnLimit)
    : compactWindowAddress(request.address, rowOffset, columnOffset, 1, 1);
  const truncated = rowOffset + rowLimit < sourceRowCount || columnOffset + columnLimit < sourceColumnCount;
  const summary = {
    workbookId: request.workbookId,
    sheetName: request.sheetName,
    address: request.address,
    mode,
    source: {
      rowCount: sourceRowCount,
      columnCount: sourceColumnCount,
      cellCount: sourceRowCount * sourceColumnCount
    },
    window: {
      address: windowAddress,
      rowOffset,
      columnOffset,
      rowCount: rowLimit,
      columnCount: columnLimit,
      cellCount: rowLimit * columnLimit
    },
    sampled: mode === "sample"
  };
  const nextPage = truncated ? { rowOffset: rowOffset + rowLimit, columnOffset } : undefined;
  if (mode === "summary" || rowLimit === 0 || columnLimit === 0) {
    return withCompactTelemetry({ ok: true, ...summary, budgetSummary: budget.applied }, { detailLevel: "summary", responseMode, truncated, nextPage, maxPayloadBytes: budget.maxChars, maxEstimatedTokens: request.maxEstimatedTokens, resourceKind: "read", resourceTitle: "Compact range summary" });
  }

  const facets = compactRangeFacets(request);
  if (facets.length === 0) {
    return withCompactTelemetry({ ok: true, ...summary, budgetSummary: budget.applied }, { detailLevel: "compact", responseMode, truncated, nextPage, maxPayloadBytes: budget.maxChars, maxEstimatedTokens: request.maxEstimatedTokens, resourceKind: "read", resourceTitle: "Compact range read" });
  }
  if (mode === "sample" && sourceRowCount > rowLimit) {
    const samples = await Promise.all(compactSampleWindows(sourceRowCount, rowLimit).map(async (sample) => {
      const sampleAddress = compactWindowAddress(request.address, sample.rowOffset, columnOffset, sample.rowCount, columnLimit);
      const sampleResult = await readRangeSnapshot(request.workbookId, request.sheetName, sampleAddress, facets);
      const sampleSnapshot = compactReadSnapshots(sampleResult)[0]?.snapshot;
      const compactSnapshot = compactRangeSnapshotPayload(sampleSnapshot, facets, sampleAddress, request.cellOutput);
      return {
        label: sample.label,
        rowOffset: sample.rowOffset,
        rowCount: sample.rowCount,
        address: sampleAddress,
        ...compactSnapshot,
        ok: (sampleResult as { ok?: boolean }).ok,
        warnings: (sampleResult as { warnings?: unknown[] }).warnings ?? []
      };
    }));
    const payload = { ok: true, ...summary, samples };
    const sourceHash = compactSourceHash({ type: "rangeSample", payload });
    return withCompactTelemetry(
      payload,
      {
        detailLevel: "compact",
        responseMode,
        truncated,
        maxPayloadBytes: budget.maxChars,
        maxEstimatedTokens: request.maxEstimatedTokens,
        resourceKind: "read",
        resourceTitle: "Compact range sample",
        resourcePayload: payload,
        resourceScope: compactReadScope({ workbookId: request.workbookId, sheetName: request.sheetName, address: request.address, mode, rowOffset, columnOffset }),
        sourceHash,
        budgetSummary: { ok: true, ...summary, sampleCount: samples.length, sourceHash, budgetSummary: budget.applied }
      }
    );
  }
  const result = await readRangeSnapshot(request.workbookId, request.sheetName, windowAddress, facets);
  if (!(result as { ok?: boolean }).ok) {
    const payload = { ok: false, ...summary, source: result };
    return withCompactTelemetry(payload, { detailLevel: "compact", responseMode, truncated, nextPage, maxPayloadBytes: budget.maxChars, maxEstimatedTokens: request.maxEstimatedTokens, resourceKind: "read", resourceTitle: "Compact range read error", resourcePayload: payload, resourceScope: compactReadScope({ workbookId: request.workbookId, sheetName: request.sheetName, address: request.address, mode, rowOffset, columnOffset }), sourceHash: compactSourceHash(payload), budgetSummary: { ok: false, ...summary, budgetSummary: budget.applied } });
  }
  const snapshot = compactReadSnapshots(result)[0]?.snapshot;
  const compactSnapshot = compactRangeSnapshotPayload(snapshot, facets, windowAddress, request.cellOutput);
  const fullSnapshot = compactRangeSnapshotPayload(snapshot, facets, windowAddress, "matrix");
  const payload = {
    ok: true,
    ...summary,
    ...compactSnapshot,
    warnings: (result as { warnings?: unknown[] }).warnings ?? [],
    telemetry: (result as { telemetry?: unknown }).telemetry
  };
  const sourceHash = compactSourceHash({ type: "range", fingerprint: (compactSnapshot.fingerprint as unknown) ?? snapshot?.fingerprint, windowAddress, facets });
  return withCompactTelemetry(
    payload,
    {
      detailLevel: "compact",
      responseMode,
      truncated,
      nextPage,
      maxPayloadBytes: budget.maxChars,
      maxEstimatedTokens: request.maxEstimatedTokens,
      resourceKind: "read",
      resourceTitle: "Compact range read",
      resourcePayload: { ...payload, ...fullSnapshot },
      resourceScope: compactReadScope({ workbookId: request.workbookId, sheetName: request.sheetName, address: request.address, windowAddress, mode, rowOffset, columnOffset }),
      sourceHash,
      nextActionRecommendation: "answer_now",
      reasoningHints: ["Compact range proof returned", "Full detail is stored behind contextId", "No additional read is required unless the user asks for exact rows"],
      confidence: "medium",
      confidenceReasons: ["Range was read successfully", truncated ? "Result is windowed or paged" : "Requested range window is complete"],
      budgetSummary: { ok: true, ...summary, warnings: (result as { warnings?: unknown[] }).warnings ?? [], sourceHash, budgetSummary: budget.applied }
    }
  );
}

async function tableSchema(selector: TableSelector) {
  return compactCacheValue(`tableSchema:${selector.workbookId}:${selector.tableName}`, async () => {
  const result = await runtime.getTableInfo(selector);
  const info = (result as { info?: any }).info;
  return withCompactTelemetry(
    {
      ok: Boolean((result as { ok?: boolean }).ok && info),
      workbookId: selector.workbookId,
      tableName: selector.tableName,
      schema: info ? tableInfoSchema(info) : undefined,
      source: info ? undefined : result
    },
    { detailLevel: "summary" }
  );
  });
}

function compactRangeSnapshotPayload(snapshot: any, facets: RangeReadFacet[], address?: string, cellOutput: CompactCellOutput = "auto"): Record<string, unknown> {
  if (!snapshot || typeof snapshot !== "object") {
    return {};
  }
  const matrixEntries = [
    ["values", snapshot.values],
    ["formulas", snapshot.formulas],
    ["text", snapshot.text],
    ["numberFormat", snapshot.numberFormat],
    ["style", snapshot.style]
  ] as const;
  const requestedEntries = matrixEntries.filter(([name, value]) => compactFacetOutputRequested(name, facets) && Array.isArray(value));
  const bounds = compactMatrixBounds(requestedEntries.map(([, value]) => value as unknown[][]));
  const output: Record<string, unknown> = {};
  if (snapshot.fingerprint !== undefined) {
    output.fingerprint = snapshot.fingerprint;
  }
  const sparseRows = compactSparseRows(requestedEntries, address);
  const useSparse = cellOutput === "sparse" || (cellOutput !== "matrix" && shouldUseCompactSparse(bounds, sparseRows.nonEmptyCells));
  output.cellOutput = useSparse ? "sparse" : "matrix";
  output.emptySummary = {
    sourceRows: bounds.sourceRows,
    sourceColumns: bounds.sourceColumns,
    sourceCells: bounds.sourceRows * bounds.sourceColumns,
    nonEmptyCells: sparseRows.nonEmptyCells,
    emptyCells: Math.max(0, bounds.sourceRows * bounds.sourceColumns - sparseRows.nonEmptyCells),
    trailingRows: Math.max(0, bounds.sourceRows - bounds.rows),
    trailingColumns: Math.max(0, bounds.sourceColumns - bounds.columns)
  };
  if (useSparse) {
    output.sparseRows = sparseRows.rows;
  }
  for (const [name, value] of requestedEntries) {
    if (!useSparse) {
      output[name] = trimMatrixToBounds(value as unknown[][], bounds);
    }
  }
  if (!useSparse && (bounds.sourceRows > bounds.rows || bounds.sourceColumns > bounds.columns)) {
    output.omittedEmpty = {
      trailingRows: Math.max(0, bounds.sourceRows - bounds.rows),
      trailingColumns: Math.max(0, bounds.sourceColumns - bounds.columns),
      sourceRows: bounds.sourceRows,
      sourceColumns: bounds.sourceColumns,
      returnedRows: bounds.rows,
      returnedColumns: bounds.columns
    };
  }
  return output;
}

function shouldUseCompactSparse(bounds: { sourceRows: number; sourceColumns: number }, nonEmptyCells: number): boolean {
  const cellCount = bounds.sourceRows * bounds.sourceColumns;
  if (nonEmptyCells === 0) return false;
  return (cellCount >= 50 && nonEmptyCells / cellCount <= 0.4) || cellCount >= 500;
}

function compactSparseRows(
  entries: ReadonlyArray<readonly [string, unknown]>,
  address?: string
): { rows: Array<{ row: number; cells: Array<Record<string, unknown>> }>; nonEmptyCells: number } {
  const parsed = address ? parseCompactA1Address(address) : undefined;
  const startRow = parsed?.startRow ?? 1;
  const startColumn = parsed?.startColumn ?? 1;
  const matrices = entries.map(([name, value]) => [compactFacetCellKey(name), value as unknown[][]] as const);
  const rowCount = matrices.reduce((max, [, matrix]) => Math.max(max, matrix.length), 0);
  const columnCount = matrices.reduce((max, [, matrix]) => Math.max(max, ...matrix.map((row) => Array.isArray(row) ? row.length : 0)), 0);
  const rows: Array<{ row: number; cells: Array<Record<string, unknown>> }> = [];
  let nonEmptyCells = 0;
  for (let rowIndex = 0; rowIndex < rowCount; rowIndex += 1) {
    const cells: Array<Record<string, unknown>> = [];
    for (let columnIndex = 0; columnIndex < columnCount; columnIndex += 1) {
      const cell: Record<string, unknown> = {};
      for (const [key, matrix] of matrices) {
        const value = matrix[rowIndex]?.[columnIndex];
        if (!isCompactEmptyCell(value)) {
          cell[key] = value;
        }
      }
      if (Object.keys(cell).length > 0) {
        const columnNumber = startColumn + columnIndex;
        const rowNumber = startRow + rowIndex;
        cells.push({
          column: compactColumnName(columnNumber),
          address: `${compactColumnName(columnNumber)}${rowNumber}`,
          ...cell
        });
        nonEmptyCells += 1;
      }
    }
    if (cells.length > 0) {
      rows.push({ row: startRow + rowIndex, cells });
    }
  }
  return { rows, nonEmptyCells };
}

function compactFacetCellKey(name: string): string {
  if (name === "values") return "value";
  if (name === "formulas") return "formula";
  return name;
}

function compactReadSnapshots(result: unknown): Array<{ snapshot?: any }> {
  const typed = result as { data?: Array<{ snapshot?: any }>; readData?: Array<{ snapshot?: any }> };
  return typed.data ?? typed.readData ?? [];
}

function compactFacetOutputRequested(name: string, facets: RangeReadFacet[]): boolean {
  if (name === "numberFormat") {
    return facets.includes("numberFormat");
  }
  if (name === "style") {
    return facets.includes("style");
  }
  return facets.includes(name as RangeReadFacet);
}

function compactMatrixBounds(matrices: unknown[][][]): { sourceRows: number; sourceColumns: number; rows: number; columns: number } {
  const sourceRows = matrices.reduce((max, matrix) => Math.max(max, matrix.length), 0);
  const sourceColumns = matrices.reduce((max, matrix) => Math.max(max, ...matrix.map((row) => Array.isArray(row) ? row.length : 0)), 0);
  let lastRow = -1;
  let lastColumn = -1;
  for (const matrix of matrices) {
    for (let rowIndex = 0; rowIndex < matrix.length; rowIndex += 1) {
      const row = matrix[rowIndex];
      if (!Array.isArray(row)) {
        continue;
      }
      for (let columnIndex = 0; columnIndex < row.length; columnIndex += 1) {
        if (!isCompactEmptyCell(row[columnIndex])) {
          lastRow = Math.max(lastRow, rowIndex);
          lastColumn = Math.max(lastColumn, columnIndex);
        }
      }
    }
  }
  return {
    sourceRows,
    sourceColumns,
    rows: lastRow + 1,
    columns: lastColumn + 1
  };
}

function trimMatrixToBounds(matrix: unknown[][], bounds: { rows: number; columns: number }): unknown[][] {
  if (bounds.rows === 0 || bounds.columns === 0) {
    return [];
  }
  return matrix.slice(0, bounds.rows).map((row) => Array.isArray(row) ? row.slice(0, bounds.columns) : []);
}

function isCompactEmptyCell(value: unknown): boolean {
  if (value === undefined || value === null || value === "") {
    return true;
  }
  if (Array.isArray(value)) {
    return value.every(isCompactEmptyCell);
  }
  if (typeof value === "object") {
    return Object.keys(value as Record<string, unknown>).length === 0;
  }
  return false;
}

function compactTablePayload(table: any): Record<string, unknown> {
  if (!table || typeof table !== "object") {
    return {};
  }
  const matrixEntries = [
    ["values", table.values],
    ["formulas", table.formulas],
    ["text", table.text],
    ["numberFormat", table.numberFormat]
  ] as const;
  const requestedEntries = matrixEntries.filter(([, value]) => Array.isArray(value));
  const rowBounds = compactTableRowBounds(requestedEntries.map(([, value]) => value as unknown[][]));
  const output: Record<string, unknown> = {};
  if (table.headers !== undefined) {
    output.headers = table.headers;
  }
  for (const [name, value] of requestedEntries) {
    output[name] = (value as unknown[][]).slice(0, rowBounds.rows);
  }
  if (rowBounds.sourceRows > rowBounds.rows) {
    output.omittedEmpty = {
      trailingRows: rowBounds.sourceRows - rowBounds.rows,
      sourceRows: rowBounds.sourceRows,
      returnedRows: rowBounds.rows
    };
  }
  return output;
}

function compactTableFromResult(result: unknown): any | undefined {
  if (!result || typeof result !== "object") {
    return undefined;
  }
  const payload = result as { table?: any; values?: unknown; headers?: unknown; formulas?: unknown; text?: unknown; numberFormat?: unknown };
  if (payload.table && typeof payload.table === "object") {
    return {
      ...payload.table,
      ...(Array.isArray(payload.values) ? { values: payload.values } : {}),
      ...(Array.isArray(payload.headers) ? { headers: payload.headers } : {}),
      ...(Array.isArray(payload.formulas) ? { formulas: payload.formulas } : {}),
      ...(Array.isArray(payload.text) ? { text: payload.text } : {}),
      ...(Array.isArray(payload.numberFormat) ? { numberFormat: payload.numberFormat } : {})
    };
  }
  if (
    Array.isArray(payload.values) ||
    Array.isArray(payload.headers) ||
    Array.isArray(payload.formulas) ||
    Array.isArray(payload.text) ||
    Array.isArray(payload.numberFormat)
  ) {
    return payload;
  }
  return undefined;
}

function compactTableRowBounds(matrices: unknown[][][]): { sourceRows: number; rows: number } {
  const sourceRows = matrices.reduce((max, matrix) => Math.max(max, matrix.length), 0);
  let lastRow = -1;
  for (const matrix of matrices) {
    for (let rowIndex = 0; rowIndex < matrix.length; rowIndex += 1) {
      const row = matrix[rowIndex];
      if (Array.isArray(row) && row.some((cell) => !isCompactEmptyCell(cell))) {
        lastRow = Math.max(lastRow, rowIndex);
      }
    }
  }
  return { sourceRows, rows: lastRow + 1 };
}

async function compactTableRead(args: ContextTableCompactReadRequest) {
  const normalized = await normalizeContextTableReadArgs(args);
  if (!normalized.ok) {
    return normalized.response;
  }
  const request = normalized.value;
  const mode = request.mode ?? "window";
  const responseMode = compactResponseMode(request.responseMode);
  const budget = compactRequestedBudget(request as TableCompactReadRequest & { budget?: Record<string, unknown> });
  const schemaResult = await runtime.getTableInfo(tableSelector(request));
  const info = (schemaResult as { info?: any }).info;
  if (!(schemaResult as { ok?: boolean }).ok || !info) {
    return withCompactTelemetry({ ok: false, workbookId: request.workbookId, tableName: request.tableName, source: schemaResult }, { detailLevel: "summary" });
  }
  const selectedColumns = compactTableColumns(info, request.columns, budget.maxColumns);
  const rowOffset = Math.min(request.rowOffset ?? 0, info.rowCount);
  const maxRows = budget.maxRows;
  const maxRowsByCells = request.maxCells !== undefined && selectedColumns.length > 0 ? Math.floor(request.maxCells / selectedColumns.length) : maxRows;
  const rowLimit = Math.min(maxRows, maxRowsByCells, Math.max(0, info.rowCount - rowOffset));
  const truncated = rowOffset + rowLimit < info.rowCount || selectedColumns.length < info.columnCount;
  const summary = {
    workbookId: request.workbookId,
    tableName: request.tableName,
    mode,
    schema: tableInfoSchema(info),
    rowOffset,
    rowLimit,
    projectedColumns: selectedColumns,
    sampled: mode === "sample"
  };
  const nextPage = truncated && rowOffset + rowLimit < info.rowCount ? { rowOffset: rowOffset + rowLimit } : undefined;
  if (mode === "summary" || rowLimit === 0 || selectedColumns.length === 0) {
    return withCompactTelemetry({ ok: true, ...summary, budgetSummary: budget.applied }, { detailLevel: "summary", responseMode, truncated, nextPage, maxPayloadBytes: budget.maxChars, maxEstimatedTokens: request.maxEstimatedTokens, resourceKind: "read", resourceTitle: "Compact table summary" });
  }
  if (mode === "sample" && info.rowCount > rowLimit) {
    const samples = await Promise.all(compactSampleWindows(info.rowCount, rowLimit).map(async (sample) => {
      const sampleResult = await runtime.readTable({
        ...tableSelector(request),
        includeValues: request.includeValues ?? true,
        includeFormulas: request.includeFormulas === true,
        includeText: request.includeText === true,
        includeNumberFormats: request.includeNumberFormats === true,
        columns: selectedColumns.map((column: { name: string }) => column.name),
        rowOffset: sample.rowOffset,
        rowLimit: sample.rowCount
      });
      const sampleTable = compactTableFromResult(sampleResult);
      const compactTable = compactTablePayload(sampleTable);
      return {
        label: sample.label,
        rowOffset: sample.rowOffset,
        rowCount: sample.rowCount,
        ok: (sampleResult as { ok?: boolean }).ok,
        ...compactTable
      };
    }));
    const payload = { ok: true, ...summary, samples };
    const sourceHash = compactSourceHash({ type: "tableSample", payload });
    return withCompactTelemetry(
      payload,
      {
        detailLevel: "compact",
        responseMode,
        truncated,
        maxPayloadBytes: budget.maxChars,
        maxEstimatedTokens: request.maxEstimatedTokens,
        resourceKind: "read",
        resourceTitle: "Compact table sample",
        resourcePayload: payload,
        resourceScope: compactReadScope({ workbookId: request.workbookId, tableName: request.tableName, mode, rowOffset }),
        sourceHash,
        budgetSummary: { ok: true, ...summary, sampleCount: samples.length, sourceHash, budgetSummary: budget.applied }
      }
    );
  }
  const result = await runtime.readTable({
    ...tableSelector(request),
    includeValues: request.includeValues ?? true,
    includeFormulas: request.includeFormulas === true,
    includeText: request.includeText === true,
    includeNumberFormats: request.includeNumberFormats === true,
    columns: selectedColumns.map((column: { name: string }) => column.name),
    rowOffset,
    rowLimit
  });
  const table = compactTableFromResult(result);
  const compactTable = compactTablePayload(table);
  const ok = Boolean((result as { ok?: boolean }).ok || table);
  const payload = {
    ok,
    ...summary,
    ...compactTable,
    source: ok ? undefined : result
  };
  const sourceHash = compactSourceHash({ type: "table", tableName: request.tableName, rowOffset, rowLimit, selectedColumns, headers: compactTable.headers, values: compactTable.values, formulas: compactTable.formulas, text: compactTable.text, numberFormat: compactTable.numberFormat });
  return withCompactTelemetry(
    payload,
    {
      detailLevel: "compact",
      responseMode,
      truncated,
      nextPage,
      maxPayloadBytes: budget.maxChars,
      maxEstimatedTokens: request.maxEstimatedTokens,
      resourceKind: "read",
      resourceTitle: "Compact table read",
      resourcePayload: payload,
      resourceScope: compactReadScope({ workbookId: request.workbookId, tableName: request.tableName, mode, rowOffset }),
      sourceHash,
      nextActionRecommendation: "answer_now",
      budgetSummary: { ok, ...summary, sourceHash, budgetSummary: budget.applied }
    }
  );
}

type LookupMatchKind = "sheet" | "table" | "column" | "header" | "entity" | "range";

interface LookupMatch {
  matchId?: string;
  kind: LookupMatchKind;
  sheetName?: string | undefined;
  tableName?: string | undefined;
  columnName?: string | undefined;
  address?: string | undefined;
  score: number;
  reason: string;
  schema?: unknown;
  preview?: unknown;
}

async function lookupSearchWorkbook(args: LookupWorkbookSearchRequest) {
  const workbookId = args.workbookId as WorkbookId;
  const context = await lookupWorkbookContext(workbookId, args.sheetNames);
  if (!context.ok) {
    return withCompactTelemetry({ ok: false, workbookId, query: args.query, source: context.source }, { detailLevel: "summary" });
  }
  const matches: LookupMatch[] = [];
  if (args.includeSheets !== false) {
    for (const sheet of context.sheets) {
      if (lookupMatches(String(sheet.name), args.query, args)) {
        matches.push(lookupMatch({
          kind: "sheet",
          sheetName: sheet.name,
          address: sheet.usedRange?.address,
          score: lookupScore(sheet.name, args.query, args.completeMatch),
          reason: "sheet name"
        }));
      }
    }
  }
  if (args.includeTables !== false) {
    matches.push(...lookupTableMatches(context.tables, args.query, args));
  }
  matches.push(...await lookupUsedRangeMatches(workbookId, context.sheets, args.query, args, "range"));
  return compactLookupResponse({
    workbookId,
    query: args.query,
    matches,
    maxMatches: args.maxMatches ?? 25,
    maxPayloadBytes: args.maxPayloadBytes,
    maxEstimatedTokens: args.maxEstimatedTokens,
    resourceTitle: `Workbook lookup: ${args.query}`
  });
}

async function lookupFindHeaders(args: LookupFindHeadersRequest) {
  const workbookId = args.workbookId as WorkbookId;
  const terms = lookupTerms(args.headers ?? (args.query ? [args.query] : []));
  const context = await lookupWorkbookContext(workbookId, args.sheetNames);
  if (!context.ok) {
    return withCompactTelemetry({ ok: false, workbookId, query: args.query, headers: args.headers, source: context.source }, { detailLevel: "summary" });
  }
  const matches: LookupMatch[] = [];
  for (const table of context.tables) {
    for (const column of table.columns ?? []) {
      const columnName = String(column.name ?? "");
      if (terms.length === 0 || terms.some((term) => lookupMatches(columnName, term, args))) {
        matches.push(lookupMatch({
          kind: "header",
          sheetName: table.sheetName,
          tableName: table.tableName,
          columnName,
          address: table.headerAddress ?? table.address,
          score: terms.length === 0 ? 0.55 : Math.max(...terms.map((term) => lookupScore(columnName, term, false))),
          reason: "table header",
          schema: { tableName: table.tableName, column }
        }));
      }
    }
  }
  const maxRowsPerSheet = args.maxRowsPerSheet ?? 10;
  const maxColumns = args.maxColumns ?? 50;
  for (const sheet of context.sheets) {
    const usedAddress = sheet.usedRange?.address;
    if (!usedAddress) {
      continue;
    }
    try {
      const result = await compactRangeRead({
        workbookId,
        sheetName: sheet.name,
        address: usedAddress,
        mode: "window",
        maxRows: maxRowsPerSheet,
        maxColumns,
        includeValues: true,
        includeText: true
      });
      const rows = lookupRowsFromCompactRead(result);
      for (let rowIndex = 0; rowIndex < rows.length; rowIndex += 1) {
        for (let columnIndex = 0; columnIndex < (rows[rowIndex] ?? []).length; columnIndex += 1) {
          const value = rows[rowIndex]?.[columnIndex];
          const text = value === undefined || value === null ? "" : String(value);
          if (text === "" || (terms.length > 0 && !terms.some((term) => lookupMatches(text, term, args)))) {
            continue;
          }
          matches.push(lookupMatch({
            kind: "header",
            sheetName: sheet.name,
            address: compactWindowAddress(usedAddress, rowIndex, columnIndex, 1, 1),
            columnName: text,
            score: terms.length === 0 ? Math.max(0.15, 0.5 - rowIndex * 0.03) : Math.max(...terms.map((term) => lookupScore(text, term, false))),
            reason: `bounded sheet header scan row ${rowIndex + 1}`
          }));
        }
      }
    } catch {
      // Lookup should remain useful even when one sheet cannot be scanned.
    }
  }
  return compactLookupResponse({
    workbookId,
    query: args.query,
    headers: args.headers,
    matches,
    maxMatches: args.maxMatches ?? 25,
    maxPayloadBytes: args.maxPayloadBytes,
    maxEstimatedTokens: args.maxEstimatedTokens,
    resourceTitle: "Header lookup"
  });
}

async function lookupFindTablesByColumns(args: LookupFindTablesByColumnsRequest) {
  const workbookId = args.workbookId as WorkbookId;
  const context = await lookupWorkbookContext(workbookId);
  if (!context.ok) {
    return withCompactTelemetry({ ok: false, workbookId, requiredColumns: args.requiredColumns, source: context.source }, { detailLevel: "summary" });
  }
  const required = lookupTerms(args.requiredColumns);
  const optional = lookupTerms(args.optionalColumns ?? []);
  const matches = context.tables
    .map((table): LookupMatch | undefined => {
      const columns = (table.columns ?? []).map((column: { name?: string }) => String(column.name ?? ""));
      const requiredHits = required.filter((term) => columns.some((column: string) => lookupNormalized(column) === term || lookupNormalized(column).includes(term)));
      if (required.length > 0 && requiredHits.length < required.length) {
        return undefined;
      }
      const optionalHits = optional.filter((term) => columns.some((column: string) => lookupNormalized(column) === term || lookupNormalized(column).includes(term)));
      const denominator = Math.max(1, required.length + optional.length * 0.5);
      const score = Math.min(1, (requiredHits.length + optionalHits.length * 0.5) / denominator);
      if (score < (args.minScore ?? 0)) {
        return undefined;
      }
      return lookupMatch({
        kind: "table",
        sheetName: table.sheetName,
        tableName: table.tableName,
        address: table.address,
        score,
        reason: `matched ${requiredHits.length}/${required.length} required and ${optionalHits.length}/${optional.length} optional columns`,
        schema: tableInfoSchema(table)
      });
    })
    .filter((match): match is LookupMatch => match !== undefined);
  return compactLookupResponse({
    workbookId,
    requiredColumns: args.requiredColumns,
    optionalColumns: args.optionalColumns,
    matches,
    maxMatches: args.maxMatches ?? 25,
    maxPayloadBytes: args.maxPayloadBytes,
    maxEstimatedTokens: args.maxEstimatedTokens,
    resourceTitle: "Table column lookup"
  });
}

async function lookupFindEntity(args: LookupFindEntityRequest) {
  const workbookId = args.workbookId as WorkbookId;
  const context = await lookupWorkbookContext(workbookId, args.sheetNames);
  if (!context.ok) {
    return withCompactTelemetry({ ok: false, workbookId, entity: args.entity, source: context.source }, { detailLevel: "summary" });
  }
  const matches = await lookupUsedRangeMatches(workbookId, context.sheets, args.entity, args, "entity");
  return compactLookupResponse({
    workbookId,
    entity: args.entity,
    entityKind: args.kind ?? "any",
    matches,
    maxMatches: args.maxMatches ?? 25,
    maxPayloadBytes: args.maxPayloadBytes,
    maxEstimatedTokens: args.maxEstimatedTokens,
    resourceTitle: `Entity lookup: ${args.entity}`
  });
}

async function lookupResolveRange(args: LookupResolveRangeRequest) {
  const workbookId = args.workbookId as WorkbookId;
  const kind = args.kind ?? "any";
  const context = await lookupWorkbookContext(workbookId, args.preferredSheetName ? [args.preferredSheetName] : undefined);
  if (!context.ok) {
    return withCompactTelemetry({ ok: false, workbookId, target: args.target, source: context.source }, { detailLevel: "summary" });
  }
  const matches: LookupMatch[] = [];
  const parsedRange = lookupParseRangeTarget(args.target, args.preferredSheetName);
  if ((kind === "any" || kind === "range") && parsedRange) {
    matches.push(lookupMatch({
      kind: "range",
      sheetName: parsedRange.sheetName,
      address: parsedRange.address,
      score: 1,
      reason: "explicit A1 range"
    }));
  }
  if (kind === "any" || kind === "table" || kind === "column") {
    for (const table of context.tables) {
      if (args.preferredTableName && table.tableName !== args.preferredTableName) {
        continue;
      }
      if ((kind === "any" || kind === "table") && lookupMatches(String(table.tableName), args.target, args)) {
        matches.push(lookupMatch({
          kind: "table",
          sheetName: table.sheetName,
          tableName: table.tableName,
          address: table.address,
          score: lookupScore(table.tableName, args.target, false),
          reason: "table name",
          schema: tableInfoSchema(table)
        }));
      }
      if (kind === "any" || kind === "column") {
        for (const column of table.columns ?? []) {
          if (lookupMatches(String(column.name), args.target, args)) {
            matches.push(lookupMatch({
              kind: "column",
              sheetName: table.sheetName,
              tableName: table.tableName,
              columnName: column.name,
              address: table.address,
              score: lookupScore(column.name, args.target, false),
              reason: "table column",
              schema: { tableName: table.tableName, column }
            }));
          }
        }
      }
    }
  }
  if (kind === "any" || kind === "header") {
    const headerRequest: LookupFindHeadersRequest = {
      workbookId,
      query: args.target,
      maxRowsPerSheet: 10
    };
    if (args.preferredSheetName !== undefined) {
      headerRequest.sheetNames = [args.preferredSheetName];
    }
    if (args.maxMatches !== undefined) {
      headerRequest.maxMatches = args.maxMatches;
    }
    const headerResult = await lookupFindHeaders(headerRequest);
    matches.push(...((headerResult as { matches?: LookupMatch[] }).matches ?? []));
  }
  if (kind === "any" || kind === "entity") {
    matches.push(...await lookupUsedRangeMatches(workbookId, context.sheets, args.target, args, "entity"));
  }
  return compactLookupResponse({
    workbookId,
    target: args.target,
    targetKind: kind,
    matches,
    maxMatches: args.maxMatches ?? 10,
    maxPayloadBytes: args.maxPayloadBytes,
    maxEstimatedTokens: args.maxEstimatedTokens,
    resourceTitle: `Range resolution: ${args.target}`
  });
}

async function lookupInspectMatch(args: LookupInspectMatchRequest) {
  const decoded = args.matchId ? lookupDecodeMatchId(args.matchId) : undefined;
  const match = { ...decoded, ...args };
  const workbookId = args.workbookId as WorkbookId;
  if (match.tableName) {
    const request: TableCompactReadRequest = {
      workbookId,
      tableName: match.tableName,
      mode: "window",
      maxRows: args.maxRows ?? 10,
      maxColumns: args.maxColumns ?? 25,
      includeValues: args.includeValues ?? true,
      includeFormulas: args.includeFormulas === true,
      includeText: args.includeText === true
    };
    if (match.columnName !== undefined) {
      request.columns = [match.columnName];
    }
    if (args.maxPayloadBytes !== undefined) {
      request.maxPayloadBytes = args.maxPayloadBytes;
    }
    if (args.maxEstimatedTokens !== undefined) {
      request.maxEstimatedTokens = args.maxEstimatedTokens;
    }
    return compactTableRead(request);
  }
  let sheetName = match.sheetName;
  let address = match.address;
  if (sheetName && !address) {
    const context = await lookupWorkbookContext(workbookId, [sheetName]);
    address = context.sheets[0]?.usedRange?.address;
  }
  if (sheetName && address) {
    const request: RangeCompactReadRequest = {
      workbookId,
      sheetName,
      address,
      mode: "window",
      maxRows: args.maxRows ?? 10,
      maxColumns: args.maxColumns ?? 25,
      includeValues: args.includeValues ?? true,
      includeFormulas: args.includeFormulas === true,
      includeText: args.includeText === true
    };
    if (args.maxPayloadBytes !== undefined) {
      request.maxPayloadBytes = args.maxPayloadBytes;
    }
    if (args.maxEstimatedTokens !== undefined) {
      request.maxEstimatedTokens = args.maxEstimatedTokens;
    }
    return compactRangeRead(request);
  }
  return withCompactTelemetry(
    {
      ok: false,
      workbookId,
      error: {
        code: "LOOKUP_MATCH_TARGET_REQUIRED",
        message: "Provide matchId from a lookup response or direct sheetName/address/tableName fields."
      }
    },
    { detailLevel: "summary" }
  );
}

async function compactValidation(args: {
  workbookId: string;
  validator: string;
  sheetName?: string;
  targetSheetName?: string;
  address?: string;
  tableName?: string;
  templateId?: string;
  snapshotId?: string;
  leftSnapshotId?: string;
  rightSnapshotId?: string;
  maxIssues?: number;
  maxPayloadBytes?: number;
  maxEstimatedTokens?: number;
  responseMode?: CompactResponseMode;
  budget?: Record<string, unknown>;
}) {
  const workbookId = args.workbookId as WorkbookId;
  const budget = compactRequestedBudget(args);
  const report = await runCompactValidator(args, workbookId);
  const issues = Array.isArray((report as { issues?: unknown[] }).issues) ? (report as { issues: unknown[] }).issues : [];
  const maxIssues = budget.maxIssues;
  const summary = {
    ok: (report as { ok?: boolean }).ok,
    workbookId,
    validator: args.validator,
    scope: (report as { scope?: unknown }).scope,
    issueCount: (report as { issueCount?: number }).issueCount ?? issues.length,
    severityCounts: compactSeverityCounts(issues),
    categories: compactIssueCategories(issues),
    examples: issues.slice(0, maxIssues).map(compactIssueExample),
    examplesTruncated: issues.length > maxIssues
  };
  return withCompactTelemetry(
    summary,
    {
      detailLevel: "summary",
      responseMode: args.responseMode,
      truncated: issues.length > maxIssues,
      storeResource: true,
      resourceKind: "validation",
      resourceTitle: `Validation detail: ${args.validator}`,
      resourcePayload: report,
      resourceScope: compactReadScope({ workbookId, validator: args.validator, sheetName: args.sheetName, address: args.address, tableName: args.tableName, templateId: args.templateId }),
      sourceHash: compactSourceHash(report),
      nextActionRecommendation: (report as { ok?: boolean }).ok ? "answer_now" : "fetch_more_context",
      reasoningHints: compactValidationReasoningHints(Boolean((report as { ok?: boolean }).ok), issues.length),
      confidence: compactValidationConfidence(Boolean((report as { ok?: boolean }).ok), issues.length, summary.examplesTruncated),
      confidenceReasons: compactValidationConfidenceReasons(Boolean((report as { ok?: boolean }).ok), issues.length, summary.examplesTruncated),
      maxPayloadBytes: budget.maxChars,
      maxEstimatedTokens: args.maxEstimatedTokens,
      budgetSummary: { ...summary, budgetSummary: budget.applied }
    }
  );
}

function compactSnapshotOptions(snapshotId: SnapshotId, options: { maxPayloadBytes?: number | undefined; maxEstimatedTokens?: number | undefined; responseMode?: CompactResponseMode | undefined; budget?: Record<string, unknown> | undefined }) {
  const args: { snapshotId: SnapshotId; maxPayloadBytes?: number; maxEstimatedTokens?: number; responseMode?: CompactResponseMode; budget?: Record<string, unknown> } = { snapshotId };
  if (options.maxPayloadBytes !== undefined) args.maxPayloadBytes = options.maxPayloadBytes;
  if (options.maxEstimatedTokens !== undefined) args.maxEstimatedTokens = options.maxEstimatedTokens;
  if (options.responseMode !== undefined) args.responseMode = options.responseMode;
  if (options.budget !== undefined) args.budget = options.budget;
  return args;
}

function compactSnapshotDiffOptions(leftSnapshotId: SnapshotId, rightSnapshotId: SnapshotId, options: { maxPayloadBytes?: number | undefined; maxEstimatedTokens?: number | undefined; responseMode?: CompactResponseMode | undefined; budget?: Record<string, unknown> | undefined }) {
  const args: { leftSnapshotId: SnapshotId; rightSnapshotId: SnapshotId; maxPayloadBytes?: number; maxEstimatedTokens?: number; responseMode?: CompactResponseMode; budget?: Record<string, unknown> } = { leftSnapshotId, rightSnapshotId };
  if (options.maxPayloadBytes !== undefined) args.maxPayloadBytes = options.maxPayloadBytes;
  if (options.maxEstimatedTokens !== undefined) args.maxEstimatedTokens = options.maxEstimatedTokens;
  if (options.responseMode !== undefined) args.responseMode = options.responseMode;
  if (options.budget !== undefined) args.budget = options.budget;
  return args;
}

function compactSnapshot(args: { snapshotId: SnapshotId; maxPayloadBytes?: number; maxEstimatedTokens?: number; responseMode?: CompactResponseMode; budget?: Record<string, unknown> }) {
  const { snapshotId, maxEstimatedTokens, responseMode } = args;
  const budget = compactRequestedBudget(args);
  const result = runtime.getSnapshot(snapshotId);
  const snapshot = (result as { snapshot?: any }).snapshot;
  const summary = snapshot
    ? {
        ok: true,
        snapshotId,
        workbookId: snapshot.workbookId,
        createdAt: snapshot.createdAt,
        reason: snapshot.reason,
        affectedRangeCount: snapshot.affectedRanges?.length ?? 0,
        affectedRanges: snapshot.affectedRanges,
        payloadRangeCount: snapshot.payload?.rangeSnapshots?.length ?? 0
      }
    : { ok: false, snapshotId, error: (result as { error?: unknown }).error };
  return withCompactTelemetry(
    summary,
    {
      detailLevel: "summary",
      responseMode,
      truncated: false,
      storeResource: Boolean(snapshot),
      resourceKind: "snapshot",
      resourceTitle: `Snapshot detail: ${snapshotId}`,
      resourcePayload: result,
      resourceScope: compactReadScope({ snapshotId, workbookId: snapshot?.workbookId }),
      sourceHash: compactSourceHash(result),
      nextActionRecommendation: "answer_now",
      reasoningHints: ["Snapshot metadata is available", "Full snapshot payload is stored behind contextId"],
      confidence: snapshot ? "high" : "low",
      confidenceReasons: snapshot ? ["Snapshot was found", "Snapshot payload was stored locally"] : ["Snapshot was not found"],
      maxPayloadBytes: budget.maxChars,
      maxEstimatedTokens,
      budgetSummary: { ...summary, budgetSummary: budget.applied }
    }
  );
}

function compactSnapshotCreationResult(result: unknown): unknown {
  const snapshot = (result as { snapshot?: any })?.snapshot;
  if (!snapshot) {
    return result;
  }
  const summary = snapshotSummary(snapshot);
  const stored = storeCompactResource("snapshot", result, {
    title: `Snapshot detail: ${summary.snapshotId}`,
    scope: compactReadScope({ snapshotId: summary.snapshotId, workbookId: summary.workbookId }),
    sourceHash: compactSourceHash(result)
  });
  return withCompactTelemetry(
    {
      ok: (result as { ok?: boolean }).ok,
      snapshot: summary
    },
    {
      detailLevel: "summary",
      resourceUri: stored.uri,
      resourceKind: "snapshot",
      resourceTitle: `Snapshot detail: ${summary.snapshotId}`,
      nextActionRecommendation: "validate_then_answer"
    }
  );
}

function snapshotSummary(snapshot: Record<string, any>) {
  return {
    snapshotId: snapshot.snapshotId,
    workbookId: snapshot.workbookId,
    createdAt: snapshot.createdAt,
    reason: snapshot.reason,
    affectedRangeCount: snapshot.affectedRanges?.length ?? 0,
    affectedRanges: snapshot.affectedRanges,
    payloadRangeCount: snapshot.payload?.rangeSnapshots?.length ?? snapshot.rangeSnapshots?.length ?? 0
  };
}

function compactSnapshotDiff(args: { leftSnapshotId: SnapshotId; rightSnapshotId: SnapshotId; maxPayloadBytes?: number; maxEstimatedTokens?: number; responseMode?: CompactResponseMode; budget?: Record<string, unknown> }) {
  const { leftSnapshotId, rightSnapshotId, maxEstimatedTokens, responseMode } = args;
  const budget = compactRequestedBudget(args);
  const result = runtime.compareSnapshots(leftSnapshotId, rightSnapshotId);
  const diff = (result as { diff?: any }).diff;
  const summary = diff
    ? {
        ok: true,
        leftSnapshotId,
        rightSnapshotId,
        title: diff.summary?.title,
        changedRangeCount: diff.summary?.changedRanges?.length ?? diff.changedRanges?.length ?? 0,
        cellsChanged: diff.summary?.cellsChanged ?? diff.cellsChanged,
        formulasChanged: diff.summary?.formulasChanged ?? diff.formulasChanged,
        stylesChanged: diff.summary?.stylesChanged ?? diff.stylesChanged,
        tablesChanged: diff.summary?.tablesChanged ?? diff.tablesChanged,
        sheetsChanged: diff.summary?.sheetsChanged ?? diff.sheetsChanged,
        destructiveLevel: diff.summary?.destructiveLevel ?? diff.destructiveLevel,
        changedRanges: (diff.summary?.changedRanges ?? diff.changedRanges ?? []).slice(0, budget.maxExamples),
        changedRangesTruncated: (diff.summary?.changedRanges ?? diff.changedRanges ?? []).length > budget.maxExamples
      }
    : { ok: false, leftSnapshotId, rightSnapshotId, error: (result as { error?: unknown }).error };
  return withCompactTelemetry(
    summary,
    {
      detailLevel: "summary",
      responseMode,
      truncated: Boolean((summary as { changedRangesTruncated?: boolean }).changedRangesTruncated),
      storeResource: Boolean(diff),
      resourceKind: "diff",
      resourceTitle: `Snapshot diff: ${leftSnapshotId}..${rightSnapshotId}`,
      resourcePayload: result,
      resourceScope: compactReadScope({ leftSnapshotId, rightSnapshotId }),
      sourceHash: compactSourceHash(result),
      nextActionRecommendation: "answer_now",
      reasoningHints: diff ? ["Compact diff summary returned", "Full diff detail is stored behind contextId", "No further diff read is required unless audit detail is requested"] : ["Diff was not available"],
      confidence: compactDiffConfidence(Boolean(diff), summary),
      confidenceReasons: compactDiffConfidenceReasons(Boolean(diff), summary),
      maxPayloadBytes: budget.maxChars,
      maxEstimatedTokens,
      budgetSummary: { ...summary, budgetSummary: budget.applied }
    }
  );
}

function compactValidationReasoningHints(ok: boolean, issueCount: number): string[] {
  if (ok && issueCount === 0) {
    return ["Workbook validation passed", "No validation issues were returned", "Agent can answer now"];
  }
  if (ok) {
    return ["Validation completed with non-blocking issues", "Inspect contextId only if the user asked for audit detail"];
  }
  return ["Validation failed", "Fetch preview/page context before claiming the workbook is fixed"];
}

function compactValidationConfidence(ok: boolean, issueCount: number, truncated: boolean): CompactConfidence {
  if (!ok) {
    return "low";
  }
  if (issueCount === 0 && !truncated) {
    return "high";
  }
  return "medium";
}

function compactValidationConfidenceReasons(ok: boolean, issueCount: number, truncated: boolean): string[] {
  const reasons: string[] = [];
  reasons.push(ok ? "Validation completed successfully" : "Validation reported failure");
  reasons.push(issueCount === 0 ? "No issues were found" : `${issueCount} issue(s) were found`);
  if (truncated) {
    reasons.push("Issue examples were truncated");
  }
  return reasons;
}

function compactDiffConfidence(diffAvailable: boolean, summary: Record<string, unknown>): CompactConfidence {
  if (!diffAvailable) {
    return "low";
  }
  if (summary.changedRangesTruncated === true) {
    return "medium";
  }
  return "high";
}

function compactDiffConfidenceReasons(diffAvailable: boolean, summary: Record<string, unknown>): string[] {
  if (!diffAvailable) {
    return ["Diff was not available"];
  }
  const reasons = ["Snapshot diff completed", "Full diff detail was stored locally"];
  if (summary.changedRangesTruncated === true) {
    reasons.push("Changed range examples were truncated");
  }
  return reasons;
}

async function runCompactValidator(args: {
  workbookId: string;
  validator: string;
  sheetName?: string;
  targetSheetName?: string;
  address?: string;
  tableName?: string;
  templateId?: string;
  snapshotId?: string;
  leftSnapshotId?: string;
  rightSnapshotId?: string;
}, workbookId: WorkbookId) {
  switch (args.validator) {
    case "workbook":
      return runtime.validateWorkbook({ workbookId });
    case "sheet":
      return runtime.validateSheet({ workbookId, sheetName: requiredCompactArg(args.sheetName, "sheetName") });
    case "formulas":
      return runtime.validateFormulas({ workbookId, ...compactOptional({ sheetName: args.sheetName, address: args.address }) });
    case "styles":
      return runtime.validateStyles(compactStylesValidationRequest(workbookId, args));
    case "tables":
      return runtime.validateTables(compactTablesValidationRequest(workbookId, args));
    case "filters":
      return runtime.validateFilters({ workbookId, ...compactOptional({ tableName: args.tableName }) });
    case "print_layout":
      return runtime.validatePrintLayout(compactPrintLayoutValidationRequest(workbookId, args));
    case "no_broken_references":
      return runtime.validateNoBrokenReferences({ workbookId, ...compactOptional({ sheetName: args.sheetName, address: args.address }) });
    case "no_formula_errors":
      return runtime.validateNoFormulaErrors({ workbookId, ...compactOptional({ sheetName: args.sheetName, address: args.address }) });
    case "no_unintended_changes":
      return runtime.validateNoUnintendedChanges(compactUnintendedChangesValidationRequest(workbookId, args));
    default:
      return {
        ok: false,
        workbookId,
        scope: args.validator,
        issueCount: 1,
        issues: [
          {
            code: "VALIDATOR_UNSUPPORTED",
            severity: "error",
            category: "validation",
            message: `Unsupported compact validator: ${args.validator}`
          }
        ]
      };
  }
}

function compactSeverityCounts(issues: unknown[]) {
  const counts: Record<string, number> = {};
  for (const issue of issues) {
    const severity = typeof issue === "object" && issue !== null && "severity" in issue ? String((issue as { severity?: unknown }).severity) : "unknown";
    counts[severity] = (counts[severity] ?? 0) + 1;
  }
  return counts;
}

function compactStylesValidationRequest(
  workbookId: WorkbookId,
  args: { sheetName?: string; templateId?: string; targetSheetName?: string }
): { workbookId: WorkbookId; templateId?: TemplateId; targetSheetName?: string; sheetName?: string } {
  const request: { workbookId: WorkbookId; templateId?: TemplateId; targetSheetName?: string; sheetName?: string } = { workbookId };
  if (args.sheetName !== undefined) {
    request.sheetName = args.sheetName;
  }
  if (args.templateId !== undefined) {
    request.templateId = args.templateId as TemplateId;
  }
  if (args.targetSheetName !== undefined) {
    request.targetSheetName = args.targetSheetName;
  }
  return request;
}

function compactTablesValidationRequest(
  workbookId: WorkbookId,
  args: { tableName?: string; templateId?: string }
): { workbookId: WorkbookId; tableName?: string; templateId?: TemplateId } {
  const request: { workbookId: WorkbookId; tableName?: string; templateId?: TemplateId } = { workbookId };
  if (args.tableName !== undefined) {
    request.tableName = args.tableName;
  }
  if (args.templateId !== undefined) {
    request.templateId = args.templateId as TemplateId;
  }
  return request;
}

function compactPrintLayoutValidationRequest(
  workbookId: WorkbookId,
  args: { templateId?: string; targetSheetName?: string }
): { workbookId: WorkbookId; templateId?: TemplateId; targetSheetName?: string } {
  const request: { workbookId: WorkbookId; templateId?: TemplateId; targetSheetName?: string } = { workbookId };
  if (args.templateId !== undefined) {
    request.templateId = args.templateId as TemplateId;
  }
  if (args.targetSheetName !== undefined) {
    request.targetSheetName = args.targetSheetName;
  }
  return request;
}

function compactUnintendedChangesValidationRequest(
  workbookId: WorkbookId,
  args: { snapshotId?: string; leftSnapshotId?: string; rightSnapshotId?: string }
): { workbookId: WorkbookId; snapshotId?: SnapshotId; leftSnapshotId?: SnapshotId; rightSnapshotId?: SnapshotId } {
  const request: { workbookId: WorkbookId; snapshotId?: SnapshotId; leftSnapshotId?: SnapshotId; rightSnapshotId?: SnapshotId } = { workbookId };
  if (args.snapshotId !== undefined) {
    request.snapshotId = args.snapshotId as SnapshotId;
  }
  if (args.leftSnapshotId !== undefined) {
    request.leftSnapshotId = args.leftSnapshotId as SnapshotId;
  }
  if (args.rightSnapshotId !== undefined) {
    request.rightSnapshotId = args.rightSnapshotId as SnapshotId;
  }
  return request;
}

function compactIssueCategories(issues: unknown[]) {
  return [...new Set(issues.map((issue) => typeof issue === "object" && issue !== null && "category" in issue ? String((issue as { category?: unknown }).category) : "unknown"))];
}

function compactIssueExample(issue: unknown) {
  if (typeof issue !== "object" || issue === null) {
    return issue;
  }
  const typed = issue as { code?: unknown; severity?: unknown; category?: unknown; message?: unknown; target?: unknown };
  return {
    code: typed.code,
    severity: typed.severity,
    category: typed.category,
    message: typed.message,
    target: typed.target
  };
}

function requiredCompactArg<T>(value: T | undefined, name: string): T {
  if (value === undefined || value === "") {
    throw new Error(`${name} is required for this compact validator.`);
  }
  return value;
}

function usedRangeCellCount(usedRange: { rowCount?: number; columnCount?: number } | undefined): number {
  return (usedRange?.rowCount ?? 0) * (usedRange?.columnCount ?? 0);
}

function tableInfoSchema(info: any) {
  return {
    workbookId: info.workbookId,
    tableName: info.tableName,
    sheetName: info.sheetName,
    address: info.address,
    headerAddress: info.headerAddress,
    rowCount: info.rowCount,
    columnCount: info.columnCount,
    columns: info.columns ?? [],
    style: info.style,
    showHeaders: info.showHeaders,
    showTotals: info.showTotals,
    showFilterButton: info.showFilterButton,
    showBandedRows: info.showBandedRows,
    showBandedColumns: info.showBandedColumns,
    hasFilters: info.filters !== undefined,
    hasSort: info.sort !== undefined
  };
}

async function lookupWorkbookContext(workbookId: WorkbookId, sheetNames?: string[]) {
  const result = await runtime.getWorkbookMap();
  const map = "map" in result ? result.map as { workbook?: any; sheets?: any[] } : undefined;
  if (!(result as { ok?: boolean }).ok || !map) {
    return { ok: false as const, workbookId, source: result, sheets: [], tables: [] };
  }
  const requestedSheets = new Set((sheetNames ?? []).map(lookupNormalized));
  const sheets = (map.sheets ?? []).filter((sheet) => requestedSheets.size === 0 || requestedSheets.has(lookupNormalized(sheet.name)));
  const tableRefs = sheets.flatMap((sheet) =>
    (sheet.tables ?? []).map((table: { name?: string; tableName?: string }) => ({
      sheetName: sheet.name,
      tableName: table.name ?? table.tableName
    }))
  ).filter((table) => table.tableName);
  const tables = await Promise.all(tableRefs.map(async (table) => {
    try {
      const infoResult = await runtime.getTableInfo({ workbookId, tableName: table.tableName as string });
      const info = (infoResult as { info?: any }).info;
      return info ? { ...info, sheetName: info.sheetName ?? table.sheetName } : undefined;
    } catch {
      return undefined;
    }
  }));
  return {
    ok: true as const,
    workbookId,
    workbook: map.workbook,
    sheets,
    tables: tables.filter((table): table is any => table !== undefined)
  };
}

function lookupTableMatches(tables: any[], query: string, options: { completeMatch?: boolean; matchCase?: boolean }): LookupMatch[] {
  const matches: LookupMatch[] = [];
  for (const table of tables) {
    const schema = tableInfoSchema(table);
    if (lookupMatches(String(table.tableName), query, options)) {
      matches.push(lookupMatch({
        kind: "table",
        sheetName: table.sheetName,
        tableName: table.tableName,
        address: table.address,
        score: lookupScore(table.tableName, query, options.completeMatch),
        reason: "table name",
        schema
      }));
    }
    for (const column of table.columns ?? []) {
      if (lookupMatches(String(column.name), query, options)) {
        matches.push(lookupMatch({
          kind: "column",
          sheetName: table.sheetName,
          tableName: table.tableName,
          columnName: column.name,
          address: table.address,
          score: lookupScore(column.name, query, options.completeMatch),
          reason: "table column",
          schema: { tableName: table.tableName, column }
        }));
      }
    }
  }
  return matches;
}

async function lookupUsedRangeMatches(
  workbookId: WorkbookId,
  sheets: any[],
  query: string,
  options: unknown,
  kind: "entity" | "range"
): Promise<LookupMatch[]> {
  const matches: LookupMatch[] = [];
  const lookupOptions = options as { completeMatch?: boolean; matchCase?: boolean; maxPreviewRows?: number };
  for (const sheet of sheets) {
    const address = sheet.usedRange?.address;
    if (!address) {
      continue;
    }
    try {
      const result = await runtime.readRangeMetadata("range.search", {
        workbookId,
        sheetName: sheet.name,
        address,
        text: query,
        completeMatch: lookupOptions.completeMatch,
        matchCase: lookupOptions.matchCase
      });
      const found = (result as { matches?: { address?: string; areaCount?: number; cellCount?: number; isNullObject?: boolean } }).matches;
      if (!found?.address || found.isNullObject) {
        continue;
      }
      const match = lookupMatch({
        kind,
        sheetName: sheet.name,
        address: found.address,
        score: lookupOptions.completeMatch ? 0.92 : 0.82,
        reason: `${kind === "entity" ? "entity" : "range"} text match`,
        preview: {
          areaCount: found.areaCount,
          cellCount: found.cellCount
        }
      });
      if ((lookupOptions.maxPreviewRows ?? 0) > 0) {
        match.preview = await lookupPreviewRange(workbookId, sheet.name, found.address, lookupOptions.maxPreviewRows ?? 3);
      }
      matches.push(match);
    } catch {
      // Ignore sheet-local search failures; other sheets and metadata can still guide the model.
    }
  }
  return matches;
}

async function lookupPreviewRange(workbookId: WorkbookId, sheetName: string, address: string, maxRows: number) {
  try {
    const result = await compactRangeRead({
      workbookId,
      sheetName,
      address,
      mode: "window",
      maxRows,
      maxColumns: 10,
      includeValues: true,
      includeText: true
    });
    return {
      address: (result as { window?: { address?: string } }).window?.address ?? address,
      values: (result as { values?: unknown }).values,
      text: (result as { text?: unknown }).text
    };
  } catch {
    return { address };
  }
}

function compactLookupResponse(args: {
  workbookId: WorkbookId;
  matches: LookupMatch[];
  maxMatches: number;
  maxPayloadBytes?: number | undefined;
  maxEstimatedTokens?: number | undefined;
  resourceTitle: string;
  [key: string]: unknown;
}) {
  const allMatches = lookupSortAndDedupe(args.matches);
  const shownMatches = allMatches.slice(0, args.maxMatches);
  const summary = {
    ok: true,
    workbookId: args.workbookId,
    ...Object.fromEntries(Object.entries(args).filter(([key]) =>
      !["workbookId", "matches", "maxMatches", "maxPayloadBytes", "maxEstimatedTokens", "resourceTitle"].includes(key)
    )),
    matchCount: allMatches.length,
    matches: shownMatches,
    matchesTruncated: shownMatches.length < allMatches.length
  };
  return withCompactTelemetry(
    summary,
    {
      detailLevel: "summary",
      truncated: shownMatches.length < allMatches.length,
      maxPayloadBytes: args.maxPayloadBytes,
      maxEstimatedTokens: args.maxEstimatedTokens,
      resourceKind: "summary",
      resourceTitle: args.resourceTitle,
      resourcePayload: { ...summary, matches: allMatches, matchesTruncated: false },
      budgetSummary: { ...summary, matches: shownMatches.map(({ preview, schema, ...match }) => match) }
    }
  );
}

function lookupMatch(match: LookupMatch): LookupMatch {
  return {
    ...match,
    score: Number(match.score.toFixed(3)),
    matchId: lookupEncodeMatchId(match)
  };
}

function lookupSortAndDedupe(matches: LookupMatch[]): LookupMatch[] {
  const best = new Map<string, LookupMatch>();
  for (const match of matches) {
    const key = [match.kind, match.sheetName, match.tableName, match.columnName, match.address].map((value) => value ?? "").join("|");
    const existing = best.get(key);
    if (!existing || match.score > existing.score) {
      best.set(key, match);
    }
  }
  return [...best.values()].sort((left, right) =>
    right.score - left.score ||
    String(left.sheetName ?? "").localeCompare(String(right.sheetName ?? "")) ||
    String(left.tableName ?? "").localeCompare(String(right.tableName ?? "")) ||
    String(left.address ?? "").localeCompare(String(right.address ?? ""))
  );
}

function lookupTerms(values: string[]): string[] {
  return values.map(lookupNormalized).filter((value) => value !== "");
}

function lookupMatches(candidate: string, query: string, options: unknown): boolean {
  const lookupOptions = options as { completeMatch?: boolean; matchCase?: boolean };
  const left = lookupOptions.matchCase ? candidate.trim() : lookupNormalized(candidate);
  const right = lookupOptions.matchCase ? query.trim() : lookupNormalized(query);
  return lookupOptions.completeMatch ? left === right : left.includes(right);
}

function lookupScore(candidate: string, query: string, completeMatch?: boolean): number {
  const left = lookupNormalized(candidate);
  const right = lookupNormalized(query);
  if (left === right) {
    return 1;
  }
  if (completeMatch) {
    return 0;
  }
  if (left.startsWith(right)) {
    return 0.9;
  }
  return left.includes(right) ? 0.75 : 0;
}

function lookupNormalized(value: unknown): string {
  return String(value ?? "").trim().toLowerCase();
}

function lookupRowsFromCompactRead(result: unknown): unknown[][] {
  const typed = result as { text?: unknown[][]; values?: unknown[][] };
  return Array.isArray(typed.text) ? typed.text : Array.isArray(typed.values) ? typed.values : [];
}

function lookupParseRangeTarget(target: string, preferredSheetName?: string) {
  const parts = target.split("!");
  const sheetName = parts.length > 1 ? parts.slice(0, -1).join("!").replace(/^'|'$/g, "") : preferredSheetName;
  const address = parts.length > 1 ? parts[parts.length - 1] : target;
  if (address === undefined) {
    return undefined;
  }
  try {
    parseCompactA1Address(address);
    return sheetName ? { sheetName, address } : undefined;
  } catch {
    return undefined;
  }
}

function lookupEncodeMatchId(match: LookupMatch): string {
  const payload = JSON.stringify({
    kind: match.kind,
    sheetName: match.sheetName,
    tableName: match.tableName,
    columnName: match.columnName,
    address: match.address
  });
  return `lookup:${Buffer.from(payload, "utf8").toString("base64url")}`;
}

function lookupDecodeMatchId(matchId: string): Partial<LookupMatch> | undefined {
  if (!matchId.startsWith("lookup:")) {
    return undefined;
  }
  try {
    return JSON.parse(Buffer.from(matchId.slice("lookup:".length), "base64url").toString("utf8")) as Partial<LookupMatch>;
  } catch {
    return undefined;
  }
}

function compactTableColumns(info: any, requested: Array<string | number> | undefined, maxColumns: number | undefined) {
  const columns = info.columns ?? [];
  if (requested?.length) {
    return columns.filter((column: { name: string; index: number; id?: number }) =>
      requested.some((item) => item === column.name || item === column.index || item === column.id)
    );
  }
  return columns.slice(0, maxColumns ?? 25);
}

function compactRangeFacets(args: RangeCompactReadRequest): RangeReadFacet[] {
  const facets: RangeReadFacet[] = [];
  if (args.includeValues !== false) {
    facets.push("values");
  }
  if (args.includeFormulas === true) {
    facets.push("formulas");
  }
  if (args.includeText === true) {
    facets.push("text");
  }
  if (args.includeNumberFormats === true) {
    facets.push("numberFormat");
  }
  if (args.includeStyles === true) {
    facets.push("style");
  }
  return facets;
}

function compactSampleWindows(totalRows: number, requestedRows: number): Array<{ label: "head" | "middle" | "tail"; rowOffset: number; rowCount: number }> {
  const sampleRows = Math.max(1, Math.min(totalRows, requestedRows));
  if (totalRows <= sampleRows) {
    return [{ label: "head", rowOffset: 0, rowCount: totalRows }];
  }
  const headRows = Math.max(1, Math.floor(sampleRows / 3));
  const middleRows = Math.max(1, Math.floor(sampleRows / 3));
  const tailRows = Math.max(1, sampleRows - headRows - middleRows);
  const middleOffset = Math.max(headRows, Math.floor((totalRows - middleRows) / 2));
  const tailOffset = Math.max(middleOffset + middleRows, totalRows - tailRows);
  const windows = [
    { label: "head" as const, rowOffset: 0, rowCount: headRows },
    { label: "middle" as const, rowOffset: middleOffset, rowCount: Math.min(middleRows, totalRows - middleOffset) },
    { label: "tail" as const, rowOffset: tailOffset, rowCount: Math.min(tailRows, totalRows - tailOffset) }
  ];
  return windows.filter((window, index, all) =>
    window.rowCount > 0 &&
    all.findIndex((candidate) => rangesOverlap(candidate.rowOffset, candidate.rowCount, window.rowOffset, window.rowCount)) === index
  );
}

function rangesOverlap(leftOffset: number, leftCount: number, rightOffset: number, rightCount: number): boolean {
  const leftEnd = leftOffset + leftCount;
  const rightEnd = rightOffset + rightCount;
  return leftOffset < rightEnd && rightOffset < leftEnd;
}

function compactWindowAddress(address: string, rowOffset: number, columnOffset: number, rowCount: number, columnCount: number): string {
  const parsed = parseCompactA1Address(address);
  const startRow = parsed.startRow + rowOffset;
  const startColumn = parsed.startColumn + columnOffset;
  const endRow = startRow + Math.max(rowCount, 1) - 1;
  const endColumn = startColumn + Math.max(columnCount, 1) - 1;
  const start = `${compactColumnName(startColumn)}${startRow}`;
  const end = `${compactColumnName(endColumn)}${endRow}`;
  return start === end ? start : `${start}:${end}`;
}

function parseCompactA1Address(address: string): ParsedCompactA1Address {
  const range = stripResourceSheetName(address).trim();
  const match = /^(?<startCol>[A-Z]+)(?<startRow>\d+)(?::(?<endCol>[A-Z]+)(?<endRow>\d+))?$/i.exec(range);
  if (!match?.groups?.startCol || !match.groups.startRow) {
    throw new Error(`Invalid A1 address: ${address}`);
  }
  const startColumn = compactColumnNumber(match.groups.startCol);
  const startRow = Number(match.groups.startRow);
  const endColumn = compactColumnNumber(match.groups.endCol ?? match.groups.startCol);
  const endRow = Number(match.groups.endRow ?? match.groups.startRow);
  if (startRow < 1 || endRow < startRow || endColumn < startColumn) {
    throw new Error(`Invalid A1 range bounds: ${address}`);
  }
  return { startRow, startColumn, endRow, endColumn };
}

function compactColumnNumber(columnName: string): number {
  let value = 0;
  for (const char of columnName.toUpperCase()) {
    const code = char.charCodeAt(0);
    if (code < 65 || code > 90) {
      throw new Error(`Invalid column name: ${columnName}`);
    }
    value = value * 26 + (code - 64);
  }
  return value;
}

function compactColumnName(columnNumber: number): string {
  let value = "";
  let n = columnNumber;
  while (n > 0) {
    n -= 1;
    value = String.fromCharCode(65 + (n % 26)) + value;
    n = Math.floor(n / 26);
  }
  return value;
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

function workflowOperations(workbookId: WorkbookId, operations: unknown[], reason?: string): ExcelOperation[] {
  return operations.map((operation, index) => workflowOperation(workbookId, operation, reason, index));
}

function workflowOperation(workbookId: WorkbookId, operation: unknown, reason: string | undefined, index: number): ExcelOperation {
  if (!operation || typeof operation !== "object") {
    throw new Error(`Workflow operation ${index + 1} must be an object.`);
  }
  const record = operation as Record<string, any>;
  if (typeof record.kind === "string" && !record.kind.startsWith("excel.")) {
    return {
      ...record,
      workbookId: (record.workbookId ?? workbookId) as WorkbookId,
      operationId: (record.operationId ?? makeId<OperationId>("op")) as OperationId,
      reason: record.reason ?? reason ?? "MCP workflow risky edit",
      destructiveLevel: record.destructiveLevel ?? workflowDestructiveLevel(record.kind)
    } as ExcelOperation;
  }

  const toolName = String(record.tool ?? record.name ?? record.operation ?? record.kind ?? "");
  const args = (record.args ?? record.input ?? record.parameters ?? record) as Record<string, any>;
  switch (toolName) {
    case "excel.range.write_values":
    case "range.write_values":
      assertWorkflowRangeArgs(toolName, args);
      return {
        kind: "range.write_values",
        operationId: makeId<OperationId>("op"),
        workbookId,
        target: workflowTarget(workbookId, args),
        values: args.values,
        preserveFormats: true,
        destructiveLevel: "values",
        reason: reason ?? "MCP workflow write values"
      } as ExcelOperation;
    case "excel.range.write_formulas":
    case "range.write_formulas":
      assertWorkflowRangeArgs(toolName, args);
      return {
        kind: "range.write_formulas",
        operationId: makeId<OperationId>("op"),
        workbookId,
        target: workflowTarget(workbookId, args),
        formulas: args.formulas,
        preserveFormats: true,
        destructiveLevel: "values",
        reason: reason ?? "MCP workflow write formulas"
      } as ExcelOperation;
    case "excel.range.write_number_formats":
    case "range.write_number_formats":
      assertWorkflowRangeArgs(toolName, args);
      return {
        kind: "range.write_number_formats",
        operationId: makeId<OperationId>("op"),
        workbookId,
        target: workflowTarget(workbookId, args),
        numberFormat: args.numberFormat,
        preserveValues: true,
        destructiveLevel: "format",
        reason: reason ?? "MCP workflow write number formats"
      } as ExcelOperation;
    case "excel.range.clear_values_keep_format":
    case "range.clear_values_keep_format":
      assertWorkflowRangeArgs(toolName, args);
      return {
        kind: "range.clear_values_keep_format",
        operationId: makeId<OperationId>("op"),
        workbookId,
        target: workflowTarget(workbookId, args),
        destructiveLevel: "values",
        reason: reason ?? "MCP workflow clear values"
      } as ExcelOperation;
    case "excel.sheet.create":
    case "sheet.create":
      if (typeof args.sheetName !== "string") {
        throw new Error(`${toolName} requires sheetName.`);
      }
      return {
        kind: "sheet.create",
        operationId: makeId<OperationId>("op"),
        workbookId,
        sheetName: args.sheetName,
        activate: args.activate,
        destructiveLevel: "structure",
        reason: reason ?? "MCP workflow sheet create"
      } as ExcelOperation;
    default:
      throw new Error(`Unsupported workflow operation: ${toolName || "unknown"}. Use scoped range writes, clear-values-keep-format, sheet.create, or canonical ExcelOperation objects.`);
  }
}

function workflowRanges(workbookId: WorkbookId, ranges: unknown): A1Range[] | undefined {
  if (!Array.isArray(ranges)) {
    return undefined;
  }
  return ranges.map((range, index) => {
    if (!range || typeof range !== "object") {
      throw new Error(`Workflow range ${index + 1} must be an object.`);
    }
    const record = range as Record<string, any>;
    if (typeof record.sheetName !== "string" || typeof record.address !== "string") {
      throw new Error(`Workflow range ${index + 1} requires sheetName and address.`);
    }
    return {
      workbookId: (record.workbookId ?? workbookId) as WorkbookId,
      sheetName: record.sheetName,
      address: record.address
    };
  });
}

function workflowTarget(workbookId: WorkbookId, args: Record<string, any>): A1Range {
  return {
    workbookId: (args.workbookId ?? workbookId) as WorkbookId,
    sheetName: args.sheetName,
    address: args.address
  };
}

function assertWorkflowRangeArgs(toolName: string, args: Record<string, any>): void {
  if (typeof args.sheetName !== "string" || typeof args.address !== "string") {
    throw new Error(`${toolName} requires sheetName and address so the workflow can snapshot a scoped range.`);
  }
}

function workflowDestructiveLevel(kind: string) {
  if (kind.startsWith("sheet.") || kind.startsWith("template.")) {
    return "structure";
  }
  if (kind.includes("style") || kind.includes("format")) {
    return "format";
  }
  if (kind.startsWith("range.")) {
    return "values";
  }
  return "workbook";
}

function styleEntriesToOperations(
  workbookId: WorkbookId,
  entries: Array<{ sheetName: string; address: string; style: Record<string, unknown>; reason?: string }>,
  reason?: string
): ExcelOperation[] {
  return entries.map((entry) => ({
    kind: "range.write_styles",
    operationId: makeId<OperationId>("op"),
    workbookId,
    target: {
      workbookId,
      sheetName: entry.sheetName,
      address: entry.address
    },
    style: entry.style,
    preserveValues: true,
    destructiveLevel: "format",
    reason: entry.reason ?? reason ?? "MCP bulk range write styles"
  }));
}

function styleBatchChunkSize(): number {
  const value = Number(process.env.OPEN_WORKBOOK_STYLE_BATCH_CHUNK_SIZE ?? 25);
  return Number.isFinite(value) && value > 0 ? Math.round(value) : 25;
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
  if (!shouldExposeMcpTool(name)) {
    return;
  }
  const registeredConfig = withIdempotencySchema(name, config);
  (mcp.registerTool as any)(name, registeredConfig, async (args: any, extra: any) => {
    const idempotencyReplay = getIdempotencyReplay(name, args);
    if (idempotencyReplay !== undefined) {
      return idempotencyReplay;
    }
    const result = await callback(args, extra);
    const shouldDecorateMutation = name !== "excel.agent.run" && isWorkbookMutatingMcpTool(name);
    const decorated = shouldDecorateMutation ? addCompactMutationProof(result) : result;
    if (shouldDecorateMutation) {
      clearCompactCache(`mutation:${name}`);
    }
    const finalResult = enforceCompactResultBudget(name, decorated);
    storeIdempotencyResult(name, args, finalResult);
    return finalResult;
  });
}

function withIdempotencySchema(toolName: string, config: Record<string, unknown>): Record<string, unknown> {
  if (!isWorkbookMutatingMcpTool(toolName)) {
    return config;
  }
  const inputSchema = config.inputSchema;
  if (!inputSchema || typeof inputSchema !== "object" || Array.isArray(inputSchema)) {
    return config;
  }
  if ("idempotencyKey" in inputSchema) {
    return config;
  }
  return {
    ...config,
    inputSchema: {
      ...(inputSchema as Record<string, unknown>),
      idempotencyKey: z.string().optional()
    }
  };
}

function getIdempotencyReplay(toolName: string, args: Record<string, unknown> | undefined) {
  const idempotencyKey = typeof args?.idempotencyKey === "string" ? args.idempotencyKey : undefined;
  if (!idempotencyKey || !isWorkbookMutatingMcpTool(toolName)) {
    return undefined;
  }
  const requestArgs = args as Record<string, unknown>;
  pruneIdempotencyRecords();
  const recordKey = compactIdempotencyRecordKey(toolName, idempotencyKey);
  const record = compactIdempotencyRecords.get(recordKey);
  if (!record) {
    return undefined;
  }
  const operationHash = compactIdempotencyOperationHash(toolName, requestArgs);
  if (record.operationHash !== operationHash) {
    return jsonResult({
      ok: false,
      toolName,
      idempotencyKey,
      idempotencyConflict: true,
      error: {
        code: "IDEMPOTENCY_KEY_CONFLICT",
        message: "This idempotencyKey was already used with a different mutation payload. Use a new key for a different edit.",
        retryable: false
      },
      nextActionRecommendation: "needs_user_confirmation",
      reasoningHints: ["Idempotency key conflict detected", "Do not retry this mutation with the same key and different payload"],
      confidence: "high",
      confidenceReasons: ["Stored operation hash differs from retry operation hash"]
    });
  }
  try {
    const replayPayload = JSON.parse(record.resultText) as Record<string, unknown>;
    return jsonResult({
      ...replayPayload,
      idempotentReplay: true,
      idempotencyKey,
      originalTransactionId: replayPayload.transactionId ?? (replayPayload.compactProof as { transactionId?: unknown } | undefined)?.transactionId,
      nextActionRecommendation: replayPayload.nextActionRecommendation ?? "answer_now",
      reasoningHints: [
        "Returned cached idempotent mutation result",
        "The mutation was not executed again",
        ...limitStringList(Array.isArray(replayPayload.reasoningHints) ? replayPayload.reasoningHints as string[] : [], COMPACT_LIMITS.maxWarnings - 2)
      ]
    });
  } catch {
    return jsonResult({
      ok: true,
      toolName,
      idempotentReplay: true,
      idempotencyKey,
      text: record.resultText,
      nextActionRecommendation: "answer_now"
    });
  }
}

function storeIdempotencyResult(toolName: string, args: Record<string, unknown> | undefined, result: unknown): void {
  const idempotencyKey = typeof args?.idempotencyKey === "string" ? args.idempotencyKey : undefined;
  if (!idempotencyKey || !isWorkbookMutatingMcpTool(toolName) || !isJsonTextResult(result)) {
    return;
  }
  const requestArgs = args as Record<string, unknown>;
  pruneIdempotencyRecords();
  compactIdempotencyRecords.set(compactIdempotencyRecordKey(toolName, idempotencyKey), {
    idempotencyKey,
    toolName,
    operationHash: compactIdempotencyOperationHash(toolName, requestArgs),
    createdAt: new Date().toISOString(),
    resultText: result.content[0]!.text
  });
}

function compactIdempotencyRecordKey(toolName: string, idempotencyKey: string): string {
  return `${toolName}:${idempotencyKey}`;
}

function compactIdempotencyOperationHash(toolName: string, args: Record<string, unknown>): string {
  const { idempotencyKey: _idempotencyKey, ...operationArgs } = args;
  return compactSourceHash({ toolName, args: operationArgs });
}

function pruneIdempotencyRecords(): void {
  const now = Date.now();
  for (const [key, record] of compactIdempotencyRecords) {
    if (now - Date.parse(record.createdAt) > COMPACT_IDEMPOTENCY_TTL_MS) {
      compactIdempotencyRecords.delete(key);
    }
  }
  const overflow = compactIdempotencyRecords.size - COMPACT_RESOURCE_LIMIT;
  if (overflow <= 0) {
    return;
  }
  for (const [key] of [...compactIdempotencyRecords.entries()].sort(([, left], [, right]) => Date.parse(left.createdAt) - Date.parse(right.createdAt)).slice(0, overflow)) {
    compactIdempotencyRecords.delete(key);
  }
}

function enforceCompactResultBudget(toolName: string, result: unknown) {
  if (!isJsonTextResult(result)) {
    return result;
  }
  if (
    toolName === "excel.agent.run" ||
    toolName === "excel.compact.get_resource" ||
    toolName === "excel.runtime.get_capabilities" ||
    toolName === "excel.range.read_compact" ||
    toolName === "excel.table.read_compact" ||
    toolName === "excel.validate.compact" ||
    toolName === "excel.snapshot.get_compact" ||
    toolName === "excel.snapshot.compare_compact" ||
    toolName === "excel.diff.get_compact"
  ) {
    return result;
  }
  const payloadBytes = Buffer.byteLength(result.content[0]!.text, "utf8");
  if (payloadBytes <= COMPACT_LIMITS.maxToolResultChars) {
    return result;
  }
  try {
    const payload = JSON.parse(result.content[0]!.text) as Record<string, unknown>;
    if (payload.resourceUri !== undefined && payload.budgetExceeded === true) {
      return result;
    }
    const stored = storeCompactResource(compactResourceKindForTool(toolName), payload, `Full result: ${toolName}`);
    const summary = compactResultSummary(toolName, payload, payloadBytes, stored);
    return jsonResult(summary);
  } catch {
    const stored = storeCompactResource("generic", result.content[0]!.text, `Full text result: ${toolName}`);
    return jsonResult({
      ok: true,
      toolName,
      detailLevel: "summary",
      responseMode: "brief",
      budgetExceeded: true,
      payloadBytes: COMPACT_LIMITS.maxToolResultChars,
      estimatedTokens: Math.ceil(COMPACT_LIMITS.maxToolResultChars / 4),
      resourcePayloadBytes: stored.payloadBytes,
      resourceEstimatedTokens: stored.estimatedTokens,
      estimatedTokensSaved: Math.max(0, stored.estimatedTokens - Math.ceil(COMPACT_LIMITS.maxToolResultChars / 4)),
      contextId: stored.resourceId,
      resourceUri: stored.uri,
      telemetry: compactTelemetrySummary(COMPACT_LIMITS.maxToolResultChars, stored, false),
      warnings: [
        {
          code: "COMPACT_RESULT_BUDGET_EXCEEDED",
          message: "Tool result exceeded compact response budget and was stored behind resourceUri."
        }
      ]
    });
  }
}

function shouldExposeMcpTool(name: string): boolean {
  return name === "excel.agent.run";
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

function agentJsonResult(value: unknown) {
  const jsonSafeValue = JSON.parse(JSON.stringify(value)) as unknown;
  const resourceLinks = Array.isArray((jsonSafeValue as { resourceLinks?: unknown[] })?.resourceLinks)
    ? ((jsonSafeValue as { resourceLinks: Array<{ uri?: unknown; name?: unknown; description?: unknown; mimeType?: unknown }> }).resourceLinks)
    : [];
  return {
    structuredContent: jsonSafeValue,
    content: [
      {
        type: "text" as const,
        text: JSON.stringify(jsonSafeValue, null, 2)
      },
      ...resourceLinks
        .filter((link) => typeof link.uri === "string")
        .map((link) => ({
          type: "resource_link" as const,
          uri: link.uri as string,
          name: typeof link.name === "string" ? link.name : link.uri as string,
          description: typeof link.description === "string" ? link.description : undefined,
          mimeType: typeof link.mimeType === "string" ? link.mimeType : "application/json"
        }))
    ]
  };
}

function isWorkbookMutatingMcpTool(name: string): boolean {
  const namespace = name.split(".")[1];
  if (!namespace || name.startsWith("excel.compact.")) {
    return false;
  }
  if (namespace === "workbook") {
    return /\.(calculate|save|save_as|restore_backup|close|embed|import)/.test(name);
  }
  if (namespace === "plan") {
    return /\.(apply|rollback)/.test(name);
  }
  if (namespace === "transaction") {
    return /\.(rollback|rollback_chain)/.test(name);
  }
  if (!new Set(["sheet", "range", "batch", "workflow", "template", "style", "formula", "table", "filter", "sort", "pivot", "chart", "names", "region", "repair", "clean"]).has(namespace)) {
    return false;
  }
  if (name === "excel.workflow.preview_risky_edit") {
    return true;
  }
  return /\.(set_|write_|create|copy|rename|delete|move|hide|unhide|protect|unprotect|clear|apply|repair|fill|append|update|resize|sort|save|restore|close|insert|merge|unmerge|lock|unlock|convert|calculate|recalculate|register|unregister|commit|rollback|cancel|refresh|invalidate|parse|normalize|trim|remove|standardize|split|import|embed)/.test(
    name
  );
}

function addCompactMutationProof(result: unknown) {
  if (!isJsonTextResult(result)) {
    return result;
  }
  try {
    const payload = JSON.parse(result.content[0]!.text) as Record<string, unknown>;
    if (payload.compactProof !== undefined) {
      return result;
    }
    const compactProof = summarizeMutationProof(payload);
    const invalidatedContextIds = invalidateCompactResourcesForMutation(payload);
    const stored = storeCompactResource("mutation", payload, "Full mutation result");
    const nextActionRecommendation = mutationNextActionRecommendation(payload, compactProof);
    const confidence = mutationConfidence(payload, compactProof, nextActionRecommendation);
    return jsonResult({
      ...compactMutationPayload(payload),
      compactProof,
      invalidatedContextIds,
      contextId: stored.resourceId,
      nextActionRecommendation,
      reasoningHints: mutationReasoningHints(nextActionRecommendation),
      confidence,
      confidenceReasons: mutationConfidenceReasons(payload, compactProof, confidence),
      telemetry: compactTelemetrySummary(Buffer.byteLength(JSON.stringify(compactMutationPayload(payload)), "utf8"), stored, false),
      resourceUri: stored.uri
    });
  } catch {
    return result;
  }
}

function isJsonTextResult(result: unknown): result is { content: Array<{ type: "text"; text: string }> } {
  return Boolean(
    result &&
    typeof result === "object" &&
    Array.isArray((result as { content?: unknown }).content) &&
    (result as { content: Array<{ text?: unknown }> }).content[0]?.text !== undefined
  );
}

function summarizeMutationProof(payload: Record<string, unknown>) {
  const diffSummary = payload.diffSummary as Record<string, unknown> | undefined;
  const telemetry = payload.telemetry as Record<string, unknown> | undefined;
  const warnings = Array.isArray(payload.warnings) ? payload.warnings : [];
  const validation = validationSummary(payload);
  return {
    ok: payload.ok,
    transactionId: payload.transactionId,
    transactionStatus: payload.transactionStatus,
    taskId: payload.taskId,
    rollbackAvailable: payload.rollbackAvailable,
    backups: payload.backups,
    changedRanges: diffSummary?.changedRanges,
    cellsChanged: diffSummary?.cellsChanged ?? telemetry?.cellsWritten,
    formulasChanged: diffSummary?.formulasChanged,
    stylesChanged: diffSummary?.stylesChanged,
    tablesChanged: diffSummary?.tablesChanged,
    sheetsChanged: diffSummary?.sheetsChanged,
    warnings: warnings.length,
    validation,
    payloadBytes: Buffer.byteLength(JSON.stringify(payload), "utf8"),
    estimatedTokens: Math.ceil(Buffer.byteLength(JSON.stringify(payload), "utf8") / 4)
  };
}

function mutationNextActionRecommendation(payload: Record<string, unknown>, compactProof: { ok?: unknown; validation?: unknown }): CompactNextActionRecommendation {
  if (payload.ok === false || compactProof.ok === false) {
    return "fetch_more_context";
  }
  if (payload.confirmationRequired === true || payload.confirmationToken !== undefined) {
    return "needs_user_confirmation";
  }
  const validation = compactProof.validation as { ok?: boolean; issueCount?: number } | undefined;
  if (validation && validation.ok === false && (validation.issueCount ?? 0) > 0) {
    return "fetch_more_context";
  }
  return "answer_now";
}

function invalidateCompactResourcesForMutation(payload: Record<string, unknown>): string[] {
  if (payload.ok === false) {
    return [];
  }
  const workbookId = findStringFieldDeep(payload, "workbookId");
  if (!workbookId) {
    return [];
  }
  const invalidated: string[] = [];
  for (const [resourceId, resource] of compactResources) {
    if (resource.kind === "mutation") {
      continue;
    }
    if (!compactResourceMatchesWorkbook(resource, workbookId)) {
      continue;
    }
    compactResources.delete(resourceId);
    invalidated.push(resourceId);
  }
  return invalidated;
}

function compactResourceMatchesWorkbook(resource: CompactStoredResource, workbookId: string): boolean {
  const scope = resource.scope;
  if (scope?.workbookId === workbookId || scope?.workbook === workbookId) {
    return true;
  }
  return findStringFieldDeep(scope, "workbookId") === workbookId;
}

function findStringFieldDeep(value: unknown, fieldName: string, seen = new Set<unknown>()): string | undefined {
  if (!value || typeof value !== "object" || seen.has(value)) {
    return undefined;
  }
  seen.add(value);
  const record = value as Record<string, unknown>;
  if (typeof record[fieldName] === "string") {
    return record[fieldName];
  }
  for (const nested of Object.values(record)) {
    if (!nested || typeof nested !== "object") {
      continue;
    }
    const found = findStringFieldDeep(nested, fieldName, seen);
    if (found !== undefined) {
      return found;
    }
  }
  return undefined;
}

function mutationReasoningHints(nextActionRecommendation: CompactNextActionRecommendation): string[] {
  if (nextActionRecommendation === "answer_now") {
    return ["Mutation proof returned", "Validation/rollback proof is included when available", "Agent can answer now"];
  }
  if (nextActionRecommendation === "needs_user_confirmation") {
    return ["Mutation requires user confirmation", "Do not apply additional changes before the user confirms"];
  }
  if (nextActionRecommendation === "validate_then_answer") {
    return ["Mutation completed without final validation", "Run compact validation before final answer"];
  }
  return ["Mutation proof indicates issues or failure", "Fetch compact context preview before final answer"];
}

function mutationConfidence(payload: Record<string, unknown>, compactProof: { ok?: unknown; validation?: unknown }, nextActionRecommendation: CompactNextActionRecommendation): CompactConfidence {
  if (nextActionRecommendation === "fetch_more_context" || payload.ok === false || compactProof.ok === false) {
    return "low";
  }
  const validation = compactProof.validation as { ok?: boolean; issueCount?: number } | undefined;
  if (validation?.ok === true && (validation.issueCount ?? 0) === 0) {
    return "high";
  }
  return nextActionRecommendation === "answer_now" ? "medium" : "low";
}

function mutationConfidenceReasons(payload: Record<string, unknown>, compactProof: { ok?: unknown; validation?: unknown }, confidence: CompactConfidence): string[] {
  const validation = compactProof.validation as { ok?: boolean; issueCount?: number } | undefined;
  const reasons = [
    payload.ok === false || compactProof.ok === false ? "Mutation result indicates failure" : "Mutation result indicates success",
    validation === undefined ? "No validation summary was present" : validation.ok === false ? "Validation summary reported failure" : "Validation summary reported success"
  ];
  if (validation?.issueCount !== undefined) {
    reasons.push(`${validation.issueCount} validation issue(s) reported`);
  }
  reasons.push(`Confidence is ${confidence}`);
  return reasons;
}

function compactMutationPayload(payload: Record<string, unknown>): Record<string, unknown> {
  const output: Record<string, unknown> = {};
  for (const [key, value] of Object.entries(payload)) {
    if (key === "beforeSnapshot" || key === "afterSnapshot") {
      output[key] = value && typeof value === "object" ? snapshotSummary(value as Record<string, any>) : value;
      continue;
    }
    if (key === "backup") {
      output[key] = summarizeBackup(value);
      continue;
    }
    if (key === "diff") {
      output[key] = summarizeDiffResult(value);
      continue;
    }
    if (key === "planPreview") {
      output[key] = summarizePlanPreview(value);
      continue;
    }
    if (key === "applyResult") {
      output[key] = summarizeOperationResult(value);
      continue;
    }
    if (key === "beforeSnapshotResult") {
      output[key] = summarizeOperationResult(value);
      continue;
    }
    output[key] = value;
  }
  return output;
}

function summarizeBackup(value: unknown): unknown {
  if (!value || typeof value !== "object") {
    return value;
  }
  const record = value as Record<string, any>;
  return {
    ok: record.ok ?? (record.backupId !== undefined || record.affectedRanges !== undefined ? true : undefined),
    backupId: record.backupId,
    workbookId: record.workbookId,
    kind: record.kind,
    reason: record.reason,
    createdAt: record.createdAt,
    affectedRangeCount: record.affectedRanges?.length ?? 0,
    affectedRanges: record.affectedRanges,
    payloadRef: record.payloadRef
  };
}

function summarizeDiffResult(value: unknown): unknown {
  if (!value || typeof value !== "object") {
    return value;
  }
  const record = value as Record<string, any>;
  const diff = record.diff ?? record;
  return {
    ok: record.ok,
    diffId: record.diffId,
    summary: diff.summary ?? {
      title: diff.title,
      changedRanges: Array.isArray(diff.changedRanges) ? diff.changedRanges.slice(0, 20) : undefined,
      changedRangesTruncated: Array.isArray(diff.changedRanges) && diff.changedRanges.length > 20,
      cellsChanged: diff.cellsChanged,
      formulasChanged: diff.formulasChanged,
      stylesChanged: diff.stylesChanged,
      tablesChanged: diff.tablesChanged,
      sheetsChanged: diff.sheetsChanged,
      destructiveLevel: diff.destructiveLevel
    }
  };
}

function summarizePlanPreview(value: unknown): unknown {
  if (!value || typeof value !== "object") {
    return value;
  }
  const record = value as Record<string, any>;
  return {
    ok: record.ok,
    planId: record.planId,
    operationCount: record.operationCount ?? record.operations?.length,
    requiredBackupCount: record.requiredBackups?.length ?? 0,
    requiredBackups: record.requiredBackups,
    diffSummary: record.diffSummary ?? summarizeDiffResult(record.diff),
    confirmationRequired: record.confirmationRequired,
    confirmationToken: record.confirmationToken,
    warnings: Array.isArray(record.warnings) ? record.warnings.length : record.warnings
  };
}

function summarizeOperationResult(value: unknown): unknown {
  if (!value || typeof value !== "object") {
    return value;
  }
  const record = value as Record<string, any>;
  return {
    ok: record.ok,
    transactionId: record.transactionId,
    transactionStatus: record.transactionStatus,
    rollbackAvailable: record.rollbackAvailable,
    backups: record.backups,
    backup: summarizeBackup(record.backup),
    warnings: Array.isArray(record.warnings) ? record.warnings.length : record.warnings,
    telemetry: record.telemetry,
    error: record.error
  };
}

function compactResourceKindForTool(toolName: string): CompactResourceKind {
  if (toolName.includes(".snapshot.")) {
    return "snapshot";
  }
  if (toolName.includes(".diff.")) {
    return "diff";
  }
  if (toolName.includes(".validate.")) {
    return "validation";
  }
  if (isWorkbookMutatingMcpTool(toolName)) {
    return "mutation";
  }
  if (toolName.includes(".read") || toolName.includes(".lookup.") || toolName.includes(".summary") || toolName.includes(".schema")) {
    return "read";
  }
  return "generic";
}

function compactResultSummary(toolName: string, payload: Record<string, unknown>, payloadBytes: number, stored: CompactStoredResource) {
  const responseBytes = Math.min(payloadBytes, COMPACT_LIMITS.maxToolResultChars);
  const transactions = summarizeTransactionList(payload.transactions);
  return {
    ok: payload.ok ?? true,
    toolName,
    detailLevel: "summary",
    budgetExceeded: true,
    payloadBytes: responseBytes,
    estimatedTokens: Math.ceil(responseBytes / 4),
    resourcePayloadBytes: stored.payloadBytes,
    resourceEstimatedTokens: stored.estimatedTokens,
    estimatedTokensSaved: Math.max(0, stored.estimatedTokens - Math.ceil(responseBytes / 4)),
    contextId: stored.resourceId,
    resourceUri: stored.uri,
    responseMode: "brief",
    nextActionRecommendation: payload.ok === false ? "fetch_more_context" : "answer_now",
    reasoningHints: payload.ok === false ? ["Tool result was compacted after failure", "Fetch context preview before final answer"] : ["Tool result was compacted", "Full detail is stored behind contextId", "Agent can answer from this summary unless audit detail is required"],
    confidence: payload.ok === false ? "low" : "medium",
    confidenceReasons: payload.ok === false ? ["Tool result indicates failure"] : ["Tool result succeeded", "Large detail was stored locally"],
    telemetry: compactTelemetrySummary(responseBytes, stored, false),
    transactionId: payload.transactionId,
    transactionStatus: payload.transactionStatus,
    taskId: payload.taskId,
    applied: payload.applied,
    previewed: payload.previewed,
    rollbackAvailable: payload.rollbackAvailable,
    backups: payload.backups,
    backup: summarizeBackup(payload.backup),
    errorStep: payload.errorStep,
    error: payload.error,
    compactProof: payload.compactProof,
    diff: payload.diff !== undefined ? summarizeDiffResult(payload.diff) : undefined,
    rollbackPreview: payload.rollbackPreview !== undefined ? summarizeRollbackPreview(payload.rollbackPreview) : undefined,
    ...(transactions !== undefined ? { transactions, transactionCount: Array.isArray(payload.transactions) ? payload.transactions.length : transactions.length } : {}),
    summary: summarizePayloadForBudget(payload),
    warnings: [
      ...asWarningArray(payload.warnings),
      {
        code: "COMPACT_RESULT_BUDGET_EXCEEDED",
        message: "Tool result exceeded compact response budget and was stored behind resourceUri."
      }
    ]
  };
}

function summarizePayloadForBudget(payload: Record<string, unknown>) {
  return {
    transactionId: payload.transactionId,
    transactionStatus: payload.transactionStatus,
    transactionCount: Array.isArray(payload.transactions) ? payload.transactions.length : undefined,
    transactions: summarizeTransactionList(payload.transactions),
    taskId: payload.taskId,
    rollbackAvailable: payload.rollbackAvailable,
    backups: payload.backups,
    backup: summarizeBackup(payload.backup),
    snapshot: payload.snapshot && typeof payload.snapshot === "object" ? snapshotSummary(payload.snapshot as Record<string, any>) : undefined,
    applied: payload.applied,
    previewed: payload.previewed,
    errorStep: payload.errorStep,
    error: payload.error,
    diffSummary: payload.diffSummary ?? summarizeDiffResult(payload.diff),
    rollbackPreview: summarizeRollbackPreview(payload.rollbackPreview),
    compactProof: payload.compactProof,
    telemetry: payload.telemetry,
    validation: validationSummary(payload),
    table: summarizeTablePayload(payload.table),
    source: summarizeSourcePayload(payload)
  };
}

function summarizeRollbackPreview(value: unknown): unknown {
  if (!value || typeof value !== "object") {
    return value;
  }
  const record = value as Record<string, any>;
  return {
    ok: record.ok,
    previewed: record.previewed,
    rollbackAvailable: record.rollbackAvailable,
    transactionId: record.transactionId,
    conflictCount: record.conflictCount ?? (Array.isArray(record.conflicts) ? record.conflicts.length : undefined),
    conflicts: Array.isArray(record.conflicts) ? record.conflicts.slice(0, COMPACT_LIMITS.maxValidationIssues) : record.conflicts,
    warnings: Array.isArray(record.warnings) ? record.warnings.length : record.warnings,
    error: record.error
  };
}

function summarizeTransactionList(value: unknown): unknown[] | undefined {
  if (!Array.isArray(value)) {
    return undefined;
  }
  return value.slice(-50).map((transaction) => {
    if (!transaction || typeof transaction !== "object") {
      return transaction;
    }
    const item = transaction as Record<string, unknown>;
    return {
      transactionId: item.transactionId,
      workbookId: item.workbookId,
      status: item.status ?? item.transactionStatus,
      operation: item.operation ?? item.operationType ?? item.toolName,
      createdAt: item.createdAt,
      updatedAt: item.updatedAt,
      committedAt: item.committedAt,
      completedAt: item.completedAt,
      backupId: item.backupId,
      parentJobId: item.parentJobId,
      dependsOn: item.dependsOn
    };
  });
}

function summarizeTablePayload(value: unknown): unknown {
  if (!value || typeof value !== "object") {
    return undefined;
  }
  const table = value as Record<string, any>;
  return {
    info: table.info
      ? {
          workbookId: table.info.workbookId,
          tableName: table.info.tableName,
          rowCount: table.info.rowCount,
          columnCount: table.info.columnCount
        }
      : undefined,
    headerCount: Array.isArray(table.headers?.[0]) ? table.headers[0].length : undefined,
    rowCount: Array.isArray(table.values) ? table.values.length : undefined,
    hasValues: Array.isArray(table.values),
    hasFormulas: Array.isArray(table.formulas),
    hasText: Array.isArray(table.text),
    hasNumberFormat: Array.isArray(table.numberFormat)
  };
}

function summarizeSourcePayload(payload: Record<string, unknown>) {
  if (Array.isArray(payload.data)) {
    return {
      dataCount: payload.data.length,
      snapshotCount: payload.data.filter((item) => Boolean((item as { snapshot?: unknown }).snapshot)).length
    };
  }
  return undefined;
}

function validationSummary(payload: Record<string, unknown>) {
  const issueCount = typeof payload.issueCount === "number" ? payload.issueCount : undefined;
  const issues = Array.isArray(payload.issues) ? payload.issues : undefined;
  const severityCounts = (payload.severityCounts ?? payload.summary) as unknown;
  if (issueCount === undefined && issues === undefined && severityCounts === undefined) {
    return undefined;
  }
  return {
    issueCount: issueCount ?? issues?.length ?? 0,
    severityCounts
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
