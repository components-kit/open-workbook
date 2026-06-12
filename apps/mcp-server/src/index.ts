#!/usr/bin/env node
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import * as z from "zod/v4";
import { RuntimeService } from "@open-workbook/backend/runtime";
import { startBackendServer } from "@open-workbook/backend/server";
import type { BatchRequest, ExcelOperation, OperationId, PlanId, WorkbookId } from "@open-workbook/protocol";
import { makeId } from "@open-workbook/protocol";

const runtime = new RuntimeService();

const host = process.env.OPEN_WORKBOOK_HOST ?? "127.0.0.1";
const port = Number(process.env.OPEN_WORKBOOK_PORT ?? 37845);
const addinPath = process.env.OPEN_WORKBOOK_ADDIN_PATH ?? "/addin";

await startBackendServer(runtime, { host, port, addinPath });
console.error(`open-workbook add-in backend listening on ws://${host}:${port}${addinPath}`);

const server = new McpServer({
  name: "open-workbook",
  version: "0.1.0"
});

registerRuntimeTools(server);
registerWorkbookTools(server);
registerRangeTools(server);
registerBatchTools(server);
registerPlanTools(server);

await server.connect(new StdioServerTransport());

function registerRuntimeTools(mcp: McpServer): void {
  mcp.registerTool(
    "excel.runtime.get_status",
    {
      title: "Get Excel runtime status",
      description: "Return backend and Excel add-in connection status.",
      inputSchema: {},
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async () => jsonResult(runtime.getStatus())
  );

  mcp.registerTool(
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
}

function registerWorkbookTools(mcp: McpServer): void {
  mcp.registerTool(
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
}

function registerRangeTools(mcp: McpServer): void {
  mcp.registerTool(
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
    async ({ workbookId, sheetName, address, includeStyles, includeFormulas }) => {
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
}

function registerBatchTools(mcp: McpServer): void {
  mcp.registerTool(
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
    async ({ workbookId, operations }) => {
      const request: BatchRequest = {
        workbookId: workbookId as WorkbookId,
        mode: "validate",
        operations: operations as ExcelOperation[]
      };
      return jsonResult(runtime.compiler.compile(request));
    }
  );

  mcp.registerTool(
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
    async ({ workbookId, operations }) => {
      const request: BatchRequest = {
        workbookId: workbookId as WorkbookId,
        mode: "dry_run",
        operations: operations as ExcelOperation[]
      };
      return jsonResult(runtime.compiler.compile(request));
    }
  );

  mcp.registerTool(
    "excel.batch.apply",
    {
      title: "Apply Excel batch",
      description: "Apply a batch through snapshots, backups, target conflict checks, and Office.js execution.",
      inputSchema: {
        workbookId: z.string(),
        operations: z.array(z.any()),
        confirmationToken: z.string().optional(),
        expectedTargetFingerprints: z.array(z.any()).optional()
      },
      annotations: {
        readOnlyHint: false,
        destructiveHint: true,
        openWorldHint: false
      }
    },
    async ({ workbookId, operations, confirmationToken, expectedTargetFingerprints }) => {
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
      return jsonResult(await runtime.applyBatch(request));
    }
  );
}

function registerPlanTools(mcp: McpServer): void {
  mcp.registerTool(
    "excel.plan.create",
    {
      title: "Create Excel plan",
      description: "Create a reversible plan from proposed Excel operations.",
      inputSchema: {
        workbookId: z.string(),
        goal: z.string(),
        operations: z.array(z.any())
      },
      annotations: {
        readOnlyHint: true,
        destructiveHint: false,
        openWorldHint: false
      }
    },
    async ({ workbookId, goal, operations }) =>
      jsonResult(
        runtime.createPlan({
          workbookId: workbookId as WorkbookId,
          goal,
          operations: operations as ExcelOperation[]
        })
      )
  );

  mcp.registerTool(
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
    async ({ planId }) => jsonResult(await runtime.previewPlan(planId as PlanId))
  );

  mcp.registerTool(
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
    async ({ planId, confirmationToken }) => jsonResult(await runtime.applyPlan(planId as PlanId, confirmationToken))
  );
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
