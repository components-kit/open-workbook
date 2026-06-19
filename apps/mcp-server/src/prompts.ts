import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import * as z from "zod/v4";

export function registerPrompts(mcp: McpServer): void {
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
      "Use `excel.agent.run` with `mode: \"prepare\"` to collect workbook context, then preview a template or sheet-copy workflow with `mode: \"preview_update\"`.",
      "Prefer a registered template. If no template is clear, ask the user to confirm the source sheet before previewing.",
      "Apply only with a follow-up `excel.agent.run` `mode: \"apply_update\"` using the returned operationId and confirmationToken.",
      "Validate formulas, styles, tables, filters, and template consistency before reporting success."
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
      "Use `excel.agent.run` `mode: \"find\"` or `mode: \"answer\"` to identify the data-entry region.",
      "Preview cleaning with `mode: \"preview_update\"`; prefer table, region, or scoped range targets.",
      "Apply once with the returned operationId and confirmationToken, then validate the affected area."
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
      "Use `excel.agent.run` to find formula errors, inspect patterns/dependencies, and preview a formula repair.",
      "Never repair formulas by writing formula strings as plain values.",
      "Apply only after preview, recalculate, and validate formula errors again."
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
      "Use `excel.agent.run` to compare style/template consistency and preview a scoped style repair.",
      "Ask before changing structure-level layout such as hidden rows/columns, freeze panes, print settings, or page layout.",
      "Validate styles, formulas, tables, filters, and print layout after applying."
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
      "Use `excel.agent.run` `mode: \"validate\"` for workbook, sheet, formula, style, table, filter, and unintended-change validation.",
      "Repair only with explicit scoped previews and backups.",
      "Save through `excel.agent.run` only when errors are clean or the user confirms known warnings."
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
      "Use `excel.agent.run` `mode: \"prepare\"` to map workbook context and ask the user for missing metrics/groupings/date ranges.",
      "Preview report creation through `mode: \"preview_update\"`; keep writes scoped to the target report sheet or regions.",
      "Apply once, then validate formulas, style consistency, tables, charts, and no unintended source changes."
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
    { title, description, argsSchema },
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
  return entries.length === 0 ? "" : `Context: ${entries.map(([key, value]) => `${key}=${String(value)}`).join(", ")}`;
}
