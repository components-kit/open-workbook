import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";
import { AGENT_DETAIL_LEVELS, AGENT_INTENT_ACTIONS } from "@components-kit/open-workbook-protocol";
import { agentRunInputSchema, normalizeAgentRunArgs } from "./agent-run.js";

describe("excel.agent.run MCP schema", () => {
  it("accepts operation lifecycle statuses returned by the backend", () => {
    const source = readFileSync(new URL("./agent-run.ts", import.meta.url), "utf8");

    expect(source).toContain("AGENT_RUN_STATUSES");
    expect(source).toContain("AGENT_RUN_MODES");
    expect(source).toContain("AGENT_DETAIL_LEVELS");
  });

  it("mentions every public agent intent action through the protocol import", () => {
    expect(AGENT_INTENT_ACTIONS).toContain("replace_range_with_styled_table");
    expect(AGENT_INTENT_ACTIONS).toContain("read_style_summary");
    expect(AGENT_INTENT_ACTIONS).toContain("format_diagnostics");
    expect(AGENT_INTENT_ACTIONS).toContain("find_similar_rows");
    expect(AGENT_INTENT_ACTIONS).toContain("read_formulas");
    expect(AGENT_INTENT_ACTIONS).toContain("improve_visual_readability");
    expect(AGENT_INTENT_ACTIONS).toContain("workbook_design_overview");
    expect(AGENT_INTENT_ACTIONS).toContain("get_permissions");
    expect(AGENT_INTENT_ACTIONS).toContain("set_permissions");
    expect(AGENT_INTENT_ACTIONS).toContain("allow_destructive_actions");
    expect(AGENT_DETAIL_LEVELS).toContain("full_table");
    expect(AGENT_DETAIL_LEVELS).toContain("semantic_index");
    expect(AGENT_DETAIL_LEVELS).toContain("workbook_design_overview");
  });

  it("exposes semantic and workflow telemetry fields in the output schema", () => {
    const source = readFileSync(new URL("./agent-run.ts", import.meta.url), "utf8");

    expect(source).toContain("workflowRoute");
    expect(source).toContain("semanticIndexStatus");
    expect(source).toContain("semanticCandidateUsed");
    expect(source).toContain("metadataPolicy");
    expect(source).toContain("readPolicy");
  });

  it("exposes task outcome fields that tell agents when to stop looping", () => {
    const source = readFileSync(new URL("./agent-run.ts", import.meta.url), "utf8");

    expect(source).toContain("taskOutcome");
    expect(source).toContain("finalAnswer");
    expect(source).toContain("agentInstruction");
    expect(source).toContain("maxRecommendedFollowupCalls");
    expect(source).toContain("requiredFollowup");
  });

  it("advertises safe default auto-apply and preview opt-out", () => {
    const source = readFileSync(new URL("./agent-run.ts", import.meta.url), "utf8");

    expect(source).toContain("safe exact small edits may auto-apply");
    expect(source).toContain("session-scoped write permission");
    expect(source).toContain("Do not ask the user to confirm every small exact edit");
    expect(source).toContain("apply_complete");
    expect(source).toContain("maxRecommendedFollowupCalls 0");
    expect(source).toContain("autoApply false");
  });

  it("advertises one-call grouped patches for multiple explicit edits", () => {
    const source = readFileSync(new URL("./agent-run.ts", import.meta.url), "utf8");

    expect(source).toContain("multiple explicit value edits from the same user request");
    expect(source).toContain("one mode:auto call with values.patches");
    expect(source).toContain("independent row/range edits should still be grouped");
    expect(source).toContain("Do not issue parallel or sequential excel.agent.run update calls");
  });

  it("advertises backend-compiled broad transforms and row-aware derivations", () => {
    const source = readFileSync(new URL("./agent-run.ts", import.meta.url), "utf8");

    expect(source).toContain("intent.action transform_values");
    expect(source).toContain("do not read full columns into model context");
    expect(source).toContain("intent.action derive_values");
    expect(source).toContain("compiles changed cells server-side");
    expect(source).toContain("formula_like calculations");
    expect(source).toContain("Payment Variance = Actual Amount - Cash Amount");
    expect(source).toContain("intent.action settle_reconciliation");
    expect(source).toContain("Payment Variance, Reconciliation Note, and Detail Notes");
    expect(source).toContain("one grouped preview");
    expect(source).toContain("intent.action transform_sheets");
    expect(source).toContain("one bounded rename plan");
  });

  it("advertises formula-first inspection and preview-safe formula workflows", () => {
    const source = readFileSync(new URL("./agent-run.ts", import.meta.url), "utf8");

    expect(source).toContain("intent.action read_formulas");
    expect(source).toContain("is this a formula");
    expect(source).toContain("raw formula");
    expect(source).toContain("never infer formula existence from displayed values or numbers alone");
    expect(source).toContain("Formula mutations, formula repairs, and formula-like broad derivations are preview/apply workflows");
    expect(source).toContain("validate_formula_against_template");
    expect(source).toContain("preview grouped formula/note repairs");
  });

  it("advertises dropdown source-list proof before value corrections", () => {
    const source = readFileSync(new URL("./agent-run.ts", import.meta.url), "utf8");

    expect(source).toContain("call intent.action read_data_validation once");
    expect(source).toContain("answer from data_validation_summary");
    expect(source).toContain("do not fetch fullResultUri");
    expect(source).toContain("update the returned source-list cell/range with mode auto");
    expect(source).toContain("inline comma-list");
    expect(source).toContain("one preview_update with intent.action write_data_validation");
  });

  it("advertises section-anchor semantic patches for row-label and column-header edits", () => {
    const source = readFileSync(new URL("./agent-run.ts", import.meta.url), "utf8");

    expect(source).toContain("sheet_summary/semantic_index anchors");
    expect(source).toContain("row label and column header");
    expect(source).toContain("values.semanticPatches");
    expect(source).toContain("instead of reading whole sections or guessing coordinates");
  });

  it("advertises cross-sheet reference search intents before broad reads", () => {
    const source = readFileSync(new URL("./agent-run.ts", import.meta.url), "utf8");

    expect(source).toContain("cross-sheet labels");
    expect(source).toContain("intent.action find_similar_rows");
    expect(source).toContain("named reference sheet");
    expect(source).toContain("instead of broad-reading sheets");
    expect(source).toContain("intent.action find_style_references");
  });

  it("advertises live Excel selection handling in the public tool description", () => {
    const source = readFileSync(new URL("./agent-run.ts", import.meta.url), "utf8");

    expect(source).toContain("current live Excel selection");
    expect(source).toContain("current cell/range/row/column");
    expect(source).toContain("selected area");
    expect(source).toContain("before asking for row or column numbers");
    expect(source).toContain("incidental for broad workbook/worksheet overview requests");
  });

  it("tells agents to read stored excel result handles through excel.agent.run, not web fetch", () => {
    const source = readFileSync(new URL("./agent-run.ts", import.meta.url), "utf8");

    expect(source).toContain("resultUri/fullResultUri values are internal Open Workbook handles");
    expect(source).toContain("not web URLs");
    expect(source).toContain("never use Webfetch/browser");
    expect(source).toContain("continuation.fullResultUri");
  });

  it("forbids Python/openpyxl fallback when the live add-in is connected", () => {
    const source = readFileSync(new URL("./agent-run.ts", import.meta.url), "utf8");

    expect(source).toContain("report that Open Workbook failure");
    expect(source).toContain("do not fall back to Python/openpyxl");
    expect(source).toContain("offline file analysis");
  });

  it("keeps packaged Open Workbook skill guidance aligned with auto small edits", () => {
    const root = new URL("../../../../", import.meta.url);
    const skill = readFileSync(new URL("skills/open-workbook-skills/SKILL.md", root), "utf8");
    const agentRun = readFileSync(new URL("skills/open-workbook-skills/references/agent-run.md", root), "utf8");
    const workflows = readFileSync(new URL("skills/open-workbook-skills/references/workflows.md", root), "utf8");
    const performance = readFileSync(new URL("skills/open-workbook-skills/references/performance.md", root), "utf8");
    const toolSelection = readFileSync(new URL("skills/open-workbook-skills/references/tool-selection.md", root), "utf8");
    const combined = [skill, agentRun, workflows, performance, toolSelection].join("\n");

    expect(combined).toContain("write access is allowed for the session");
    expect(combined).toContain("do not ask the user to confirm every small exact edit");
    expect(combined).toContain("dropdown options are wrong");
    expect(combined).toContain("source-list proof");
    expect(combined).toContain("values.semanticPatches");
    expect(combined).toContain("row label and column header");
    expect(combined).toContain("Multiple Updates");
    expect(combined).toContain("different topic does not mean different tool call");
    expect(combined).toContain("same user instruction plus explicit ranges means one grouped patch");
    expect(combined).toContain("Do not split independent exact edits into separate calls");
    expect(combined).toContain("transform_values");
    expect(combined).toContain("derive_values");
    expect(combined).toContain("read_formulas");
    expect(combined).toContain("never infer formula existence from displayed values");
    expect(combined).toContain("formula_like");
    expect(combined).toContain("settle_reconciliation");
    expect(combined).toContain("Payment Variance");
    expect(combined).toContain("Reconciliation Note");
    expect(combined).toContain("Detail Notes");
    expect(combined).toContain("transform_sheets");
    expect(combined).toContain("improve_visual_readability");
    expect(combined).toContain("values.visualReadability");
    expect(combined).toContain("comprehensive validation/formula suggestions remain preview-only");
    expect(combined).toContain("do not apply dropdowns, formulas, inserted rows/columns, or summary blocks through the visual styling apply path");
    expect(combined).toContain("do not fetch full source/target columns");
    expect(combined).not.toContain("use `excel.agent.run` with `mode: \"preview_update\"` and then `mode: \"apply_update\"` for scoped value edits");
    expect(combined).not.toContain("group related range value edits with `values.patches` in one `preview_update`");
  });

  it("accepts continuation-only stored result fetches", () => {
    const schema = agentRunInputSchema();

    expect((schema.request as any).safeParse(undefined).success).toBe(true);
    expect((schema.continuation as any).parse({
      fullResultUri: "excel://agent/results/agentres_1?view=full"
    })).toEqual({ fullResultUri: "excel://agent/results/agentres_1?view=full" });
  });

  it("normalizes confirmation-token calls that put apply_update in responseMode", () => {
    const schema = agentRunInputSchema();
    const parsed = {
      operationId: "agentop_1",
      confirmationToken: "confirm_1",
      responseMode: "apply_update",
      request: "Apply the update"
    };

    expect((schema.responseMode as any).parse("apply_update")).toBe("apply_update");
    expect(normalizeAgentRunArgs(parsed as any)).toEqual({
      operationId: "agentop_1",
      confirmationToken: "confirm_1",
      mode: "apply_update",
      request: "Apply the update"
    });
  });

  it("allows cache invalidation fields returned after successful applies", () => {
    const source = readFileSync(new URL("./agent-run.ts", import.meta.url), "utf8");

    expect(source).toContain("invalidatedContextIds");
    expect(source).toContain("invalidatedResourceUris");
  });

  it("normalizes JSON-string structured fields from lenient MCP clients", () => {
    const schema = agentRunInputSchema();

    expect((schema.target as any).parse("{\"sheetName\":\"Booking\",\"range\":\"A1:X7\"}")).toEqual({ sheetName: "Booking", range: "A1:X7" });
    expect((schema.intent as any).parse("{\"action\":\"read_values\",\"targetHints\":[\"Booking\"]}")).toEqual({ action: "read_values", targetHints: ["Booking"] });
    expect((schema.continuation as any).parse("{\"workbookContextId\":\"wbctx_1\",\"fullResultUri\":\"excel://agent/results/agentres_1?view=full\",\"freshness\":{\"workbookId\":\"workbook_1\",\"workbookContentVersion\":4,\"workbookStructureHash\":\"abc\"}}")).toEqual({
      workbookContextId: "wbctx_1",
      fullResultUri: "excel://agent/results/agentres_1?view=full",
      freshness: {
        workbookId: "workbook_1",
        workbookContentVersion: 4,
        workbookStructureHash: "abc"
      }
    });
    expect((schema.values as any).parse({
      patches: [
        {
          target: "{\"sheetName\":\"Booking\",\"range\":\"A1:B2\"}",
          values: [[1, 2]]
        }
      ]
    })).toEqual({
      patches: [
        {
          target: { sheetName: "Booking", range: "A1:B2" },
          values: [[1, 2]]
        }
      ]
    });
  });

  it("accepts common structured update payloads without requiring agents to send schemas", () => {
    const schema = agentRunInputSchema();

    expect((schema.values as any).parse({
      values: [["Reviewed"]],
      style: { fillColor: "#1F4E78", fontColor: "#FFFFFF", fontBold: true },
      options: ["Open", "Reviewed", "Closed"],
      validation: { type: "list", source: ["Open", "Reviewed", "Closed"], inCellDropDown: true },
      rule: { type: "custom", formula: "=$E2=\"Open\"", style: { fillColor: "#FFFF00" } },
      columnOrder: [2, 1],
      numberFormat: "dd/mm/yyyy"
    })).toMatchObject({
      values: [["Reviewed"]],
      style: { fillColor: "#1F4E78", fontColor: "#FFFFFF", fontBold: true },
      options: ["Open", "Reviewed", "Closed"],
      validation: { type: "list", source: ["Open", "Reviewed", "Closed"], inCellDropDown: true },
      rule: { type: "custom", formula: "=$E2=\"Open\"", style: { fillColor: "#FFFF00" } },
      columnOrder: [2, 1],
      numberFormat: "dd/mm/yyyy"
    });
    expect((schema.values as any).parse({ numberFormats: [["dd/mm/yyyy"]] })).toMatchObject({ numberFormats: [["dd/mm/yyyy"]] });
    expect((schema.values as any).parse({
      rows: [
        { index: 1, values: ["2026-01-04", "Northwind", "Support", 525, "Closed"] }
      ]
    })).toMatchObject({
      rows: [
        { index: 1, values: ["2026-01-04", "Northwind", "Support", 525, "Closed"] }
      ]
    });
  });

  it("rejects malformed common update payloads with field-level schema errors", () => {
    const schema = agentRunInputSchema();

    expect(() => (schema.values as any).parse({ style: { fillColor: 42 } })).toThrow(/fillColor/i);
    expect(() => (schema.values as any).parse({ options: ["Open", 7] })).toThrow(/options/i);
    expect(() => (schema.values as any).parse({ validation: { source: ["Open", 7] } })).toThrow(/source/i);
    expect(() => (schema.values as any).parse({ rule: { formula: 42, style: { fillColor: "#FFFF00" } } })).toThrow(/formula/i);
    expect(() => (schema.values as any).parse({ columnOrder: [2, { bad: true }] })).toThrow(/columnOrder/i);
    expect(() => (schema.values as any).parse({ numberFormat: [["dd/mm/yyyy", 42]] })).toThrow(/numberFormat/i);
  });
});
