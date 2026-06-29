import { describe, expect, it } from "vitest";
import { AgentOrchestrator } from "./agent-orchestrator.js";
import { FakeAgentRuntime, createCachedMetadata, selectionInfo, sheets, workbookId } from "./agent-orchestrator.test-support.js";

describe("AgentOrchestrator Preview Apply Safety", () => {
  const valuePatch = (sheetName: string, range: string, values: unknown[][]) => ({
    patches: [{ target: { sheetName, range }, values }]
  });

  it("applies visual readability to Thai invoice body after grouped headers without skipping all rules", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_thai_invoices_grouped");
      const headers = [
        "Billing Status",
        "Subcontractor Paid?",
        "Booking Number",
        "Customer",
        "Job Status",
        "Load Date",
        "Job Price",
        "Lift On Fee",
        "Lift Off Fee",
        "Total Lift Fee",
        "Other Expenses",
        "Total Billing Amount",
        "Withholding Tax",
        "Net Receipt",
        "Subcontracted?"
      ];
      const invoiceColumns = headers.map((header, index) => ({
        name: header,
        normalizedName: header,
        inferredType: "unknown" as const,
        role: "unknown" as const,
        importance: 0.55,
        index,
        letter: String.fromCharCode("A".charCodeAt(0) + index)
      }));
      metadata.workbook = { ...metadata.workbook, activeSheet: "Invoices", name: "June.xlsx" };
      metadata.sheets = [{
        id: "sheet:Invoices",
        name: "Invoices",
        index: 0,
        usedRange: "A1:O1002",
        rowCount: 1002,
        columnCount: 15,
        kind: "transaction",
        headers: [{
          id: "header:Invoices:wide",
          sheetName: "Invoices",
          row: 2,
          range: "A2:O1002",
          confidence: 0.9,
          columns: invoiceColumns
        }],
        tableIds: ["table:InvoicesTable"],
        sectionIds: ["section:Invoices:grouped-header"],
        summaryBlockIds: ["summary:Invoices:grouped-header"],
        formulaRegionIds: []
      }];
      metadata.tables = [{
        id: "table:InvoicesTable",
        sheetName: "Invoices",
        name: "InvoicesTable",
        range: "A2:O1002",
        columns: invoiceColumns
      }];
      metadata.sections = [{
        id: "section:Invoices:grouped-header",
        sheetName: "Invoices",
        label: "Grouped header",
        kind: "summary",
        range: "A1:O2",
        columns: [],
        labels: ["Status", "Booking Details", "Amount", "Subcontracted Work"],
        rowCount: 2,
        columnCount: 15,
        nonEmptyCellCount: 19,
        confidence: 0.95
      }];
      metadata.summaryBlocks = [{
        id: "summary:Invoices:grouped-header",
        sheetName: "Invoices",
        range: "A1:O2",
        labels: ["Status", "Booking Details", "Amount", "Subcontracted Work"],
        confidence: 0.95
      }];
      agent.metadataCache.set(metadata);

      const preview = await agent.run({
        request: "Make Invoices easier to read. Header is already good, focus on each column and data cells.",
        mode: "preview_update",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "improve_visual_readability" },
        target: { sheetName: "Invoices", tableName: "InvoicesTable" }
      });

      const visualPlan = (preview.answer as any).visualPlan;
      expect(preview.status).toBe("PREVIEW_READY");
      expect(preview.nextAction).toBe("call_apply_update");
      expect((preview.metrics as any).operationCount).toBeGreaterThan(0);
      expect((preview.metrics as any).skippedRuleCount).toBeLessThan((preview.metrics as any).groupedOperationCount);
      expect((preview.answer as any).detected).toMatchObject({
        headerRow: 2,
        headerRange: "A2:O2",
        dataRange: "A3:O1002"
      });
      expect((preview.answer as any).columnRoles.map((column: any) => [column.column, column.role])).toEqual(expect.arrayContaining([
        ["A", "status"],
        ["C", "id"],
        ["D", "entity"],
        ["E", "status"],
        ["F", "date"],
        ["G", "money"]
      ]));
      expect(visualPlan.skipped.map((skip: any) => skip.reason).join(" ")).not.toMatch(/column\.G\.number_format.*protected/);

      const applied = await agent.run({
        request: "Apply visual readability preview",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations.map((operation) => operation.kind)).toEqual(expect.arrayContaining([
        "range.write_styles_many",
        "range.write_number_formats_many"
      ]));
      expect(runtime.lastBatchOperations.some((operation) => operation.kind === "range.apply_autofilter")).toBe(false);
      expect(visualPlan.skipped).toEqual(expect.arrayContaining([
        expect.objectContaining({ ruleId: "layout.filter", reason: expect.stringContaining("already provided by the detected Excel table") })
      ]));
      const numberFormatOperation = runtime.lastBatchOperations.find((operation) => operation.kind === "range.write_number_formats_many") as any;
      expect(numberFormatOperation.entries).toEqual(expect.arrayContaining([
        expect.objectContaining({ target: expect.objectContaining({ sheetName: "Invoices", address: "F3:F1002" }) }),
        expect.objectContaining({ target: expect.objectContaining({ sheetName: "Invoices", address: "G3:G1002" }) }),
        expect.objectContaining({ target: expect.objectContaining({ sheetName: "Invoices", address: "N3:N1002" }) })
      ]));
      expect(numberFormatOperation.entries.find((entry: any) => entry.target.address === "G3:G1002").numberFormat[0][0]).toBe("#,##0.00");
      const styleOperation = runtime.lastBatchOperations.find((operation) => operation.kind === "range.write_styles_many") as any;
      expect(styleOperation.entries).toEqual(expect.arrayContaining([
        expect.objectContaining({ target: expect.objectContaining({ sheetName: "Invoices", address: "A3:A1002" }), style: expect.objectContaining({ horizontalAlignment: "Center" }) }),
        expect.objectContaining({ target: expect.objectContaining({ sheetName: "Invoices", address: "G3:G1002" }), style: expect.objectContaining({ horizontalAlignment: "Right" }) })
      ]));
    });

  it("previews and applies visual readability safe operations through the update lifecycle", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Make this sheet easier to read",
        mode: "preview_update",
        intent: { action: "improve_visual_readability" },
        target: { sheetName: "Data" },
        values: {
          visualReadability: {
            styleDepth: "standard",
            profile: "auto",
            density: "comfortable",
            preserveFormulas: true,
            preserveExistingStyle: true
          }
        }
      });
      const applied = await agent.run({
        request: "Apply visual readability preview",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });
      expect(preview.status).toBe("PREVIEW_READY");
      expect(preview.operationId).toBeTruthy();
      expect(preview.confirmationToken).toBeTruthy();
      expect((preview.answer as any).kind).toBe("visual_readability_preview");
      expect((preview.answer as any).defaults).toEqual({
        styleDepth: "standard",
        profile: "record_tracker",
        density: "comfortable",
        preserveFormulas: true,
        preserveExistingStyle: true,
        stylePreservationMode: "protected_regions",
        allowValidationSuggestions: false,
        allowFormulaSuggestions: false,
        allowReplaceConditionalFormatting: false,
        allowReplaceDataValidation: false,
        allowInsertRowsOrColumns: false,
        applySuggestionBuckets: []
      });
      expect((preview.answer as any).detected).toMatchObject({
        sheetName: "Data",
        usedRange: "A1:D4",
        headerRow: 1,
        dataRange: "A2:D4",
        tableRanges: ["A1:D4"],
        hasFilter: true,
        detectionSource: "metadata"
      });
      expect((preview.answer as any).columnRoles.map((column: any) => [column.column, column.header, column.role])).toEqual([
        ["A", "Date", "date"],
        ["B", "Account", "entity"],
        ["C", "Amount", "money"],
        ["D", "Status", "status"]
      ]);
      expect((preview.answer as any).sheetType).toBe("record_tracker");
      const visualPlan = (preview.answer as any).visualPlan;
      expect(visualPlan.compilerStatus).toBe("preview_compiled_apply_pending");
      expect(visualPlan.counts.totalRules).toBeGreaterThan(0);
      expect(visualPlan.counts.columnRules).toBeGreaterThan(0);
      expect(visualPlan.counts.groupRules).toBeGreaterThan(0);
      expect(visualPlan.counts.conditionalRules).toBeGreaterThan(0);
      expect(visualPlan.ruleScopes.column).toBeGreaterThan(0);
      expect(visualPlan.operationCount).toBeGreaterThan(0);
      expect(visualPlan.skipped.map((skip: any) => skip.ruleId)).toContain("layout.freeze_header");
      expect(visualPlan.ruleIds).toEqual(expect.arrayContaining([
        "layout.header_style",
        "layout.filter",
        "column.A.width",
        "column.C.number_format",
        "conditional.D.missing_required"
      ]));
      expect(preview.changes.length).toBeGreaterThan(0);
      expect(preview.resourceLinks.map((link) => link.uri)).toContain(`excel://agent/operations/${preview.operationId}`);
      expect(preview.warnings.join(" ")).toContain("safe visual operation");
      expect(applied.status).toBe("SUCCESS");
      expect((applied.answer as any).kind).toBe("apply_update_result");
      expect((applied.answer as any).ok).toBe(true);
      expect(runtime.lastBatchOperations.map((operation) => operation.kind)).toEqual([
        "range.write_styles_many",
        "range.write_conditional_formatting",
        "range.write_conditional_formatting",
        "range.write_conditional_formatting",
        "range.write_number_formats_many"
      ]);
      expect(runtime.lastBatchOperations.some((operation) => operation.kind === "range.apply_autofilter")).toBe(false);
      expect(visualPlan.skipped).toEqual(expect.arrayContaining([
        expect.objectContaining({ ruleId: "layout.filter", reason: expect.stringContaining("already provided by the detected Excel table") })
      ]));
      expect(runtime.lastBatchOperations.some((operation) => operation.kind === "range.write_values" || operation.kind === "range.write_formulas")).toBe(false);
      const styleOperation = runtime.lastBatchOperations.find((operation) => operation.kind === "range.write_styles_many") as any;
      expect(styleOperation.entries.every((entry: any) => entry.preserveValues === true)).toBe(true);
      const numberFormatOperation = runtime.lastBatchOperations.find((operation) => operation.kind === "range.write_number_formats_many") as any;
      expect(numberFormatOperation.entries.every((entry: any) => entry.preserveValues === true)).toBe(true);
      expect(runtime.lastBatchOperations.filter((operation) => operation.kind === "range.write_conditional_formatting").map((operation: any) => operation.rule.formula)).toEqual(expect.arrayContaining([
        '=AND(COUNTA($A2:$D2)>0,$A2="")',
        '=AND(COUNTA($A2:$D2)>0,$D2="")'
      ]));
      expect(runtime.writeBatchCount).toBe(1);
    });

  it("keeps basic visual readability previews to layout and column rules", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Make this sheet easier to read",
        mode: "preview_update",
        intent: { action: "improve_visual_readability" },
        target: { sheetName: "Data" },
        values: {
          visualReadability: {
            styleDepth: "basic"
          }
        }
      });

      const visualPlan = (preview.answer as any).visualPlan;
      expect(preview.status).toBe("PREVIEW_READY");
      expect(visualPlan.compilerStatus).toBe("preview_compiled_apply_pending");
      expect(visualPlan.counts.totalRules).toBeGreaterThan(0);
      expect(visualPlan.counts.columnRules).toBeGreaterThan(0);
      expect(visualPlan.counts.groupRules).toBe(0);
      expect(visualPlan.counts.conditionalRules).toBe(0);
      expect(visualPlan.operationCount).toBeGreaterThan(0);
      expect(visualPlan.ruleIds).toEqual(expect.arrayContaining([
        "layout.header_style",
        "column.A.width",
        "column.C.number_format"
      ]));
      expect(visualPlan.ruleIds.some((ruleId: string) => ruleId.startsWith("group."))).toBe(false);
      expect(visualPlan.ruleIds.some((ruleId: string) => ruleId.startsWith("conditional."))).toBe(false);
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("keeps comprehensive visual readability validation and formula suggestions preview-only", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Run a comprehensive visual readability update and include suggestions",
        mode: "preview_update",
        intent: { action: "improve_visual_readability" },
        target: { sheetName: "Data" },
        values: {
          visualReadability: {
            styleDepth: "comprehensive"
          }
        }
      });
      const applied = await agent.run({
        request: "Apply visual readability preview",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      const visualPlan = (preview.answer as any).visualPlan;
      expect(preview.status).toBe("PREVIEW_READY");
      expect(visualPlan.counts.validationSuggestions).toBeGreaterThan(0);
      expect(visualPlan.counts.formulaSuggestions).toBeGreaterThan(0);
      expect(visualPlan.validationSuggestions).toEqual(expect.arrayContaining([
        expect.objectContaining({ id: "validation.D.dropdown", risk: "medium", existingValidation: "not_detected" })
      ]));
      expect(visualPlan.formulaSuggestions).toEqual(expect.arrayContaining([
        expect.objectContaining({ id: "formula.overdue_flag", risk: "medium" })
      ]));
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations.some((operation) => operation.kind === "range.write_data_validation" || operation.kind === "range.write_formulas")).toBe(false);
    });

  it("applies visual readability validation suggestions only when the validation bucket is requested", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Make this sheet easier to read and add dropdowns",
        mode: "preview_update",
        intent: { action: "improve_visual_readability" },
        target: { sheetName: "Data" },
        values: {
          visualReadability: {
            applySuggestionBuckets: ["validation"]
          }
        }
      });
      const applied = await agent.run({
        request: "Apply visual readability preview",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      const visualPlan = (preview.answer as any).visualPlan;
      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).defaults.applySuggestionBuckets).toEqual(["validation"]);
      expect(visualPlan.validationSuggestions).toEqual(expect.arrayContaining([
        expect.objectContaining({ id: "validation.D.dropdown", target: "D2:D4" })
      ]));
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations.some((operation) => operation.kind === "range.write_data_validation")).toBe(true);
      expect(runtime.lastBatchOperations.find((operation) => operation.kind === "range.write_data_validation")).toMatchObject({
        target: { sheetName: "Data", address: "D2:D4" },
        validation: { type: "list", source: ["Open", "In Progress", "Blocked", "Done"] }
      });
      expect(runtime.lastBatchOperations.some((operation) => operation.kind === "range.write_formulas")).toBe(false);
    });

  it("applies freeze column suggestions when the freeze_panes bucket is requested", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Make this easier to read and freeze first column",
        mode: "preview_update",
        intent: { action: "improve_visual_readability" },
        target: { sheetName: "Data" },
        values: {
          visualReadability: {
            applySuggestionBuckets: ["freeze_panes"]
          }
        }
      });
      const applied = await agent.run({
        request: "Apply visual readability preview",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      const visualPlan = (preview.answer as any).visualPlan;
      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).defaults.freezePanes).toEqual({ columns: 1 });
      expect(visualPlan.ruleIds).toEqual(expect.arrayContaining(["layout.freeze_header", "layout.freeze_columns"]));
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations.filter((operation) => operation.kind === "sheet.freeze_panes")).toEqual([
        expect.objectContaining({ kind: "sheet.freeze_panes", sheetName: "Data", rows: 1, columns: 1 })
      ]);
    });

  it("previews and applies direct unfreeze pane requests without visual readability", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Unfreeze all panes in Data",
        mode: "preview_update",
        intent: { action: "freeze_panes" },
        target: { sheetName: "Data" }
      });
      const applied = await agent.run({
        request: "Apply unfreeze panes",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe("freeze_panes_preview");
      expect((preview.answer as any).freezePanes).toEqual({ rows: 0, columns: 0 });
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations.filter((operation) => operation.kind === "sheet.freeze_panes")).toEqual([
        expect.objectContaining({ kind: "sheet.freeze_panes", sheetName: "Data", rows: 0, columns: 0 })
      ]);
    });

  it("keeps visual readability reference-style adaptation preview-only", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Make Apr 2026 look like May 2026 without changing formulas",
        mode: "preview_update",
        intent: { action: "improve_visual_readability" },
        target: { sheetName: "Apr 2026" },
        values: {
          visualReadability: {
            referenceStyle: {
              sheet: "May 2026",
              adaptToTargetStructure: true,
              preserveTargetValues: true,
              preserveFormulas: true
            }
          }
        }
      });
      const applied = await agent.run({
        request: "Apply visual readability preview",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      const visualPlan = (preview.answer as any).visualPlan;
      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).defaults.referenceStyle).toMatchObject({
        sheetName: "May 2026",
        adaptToTargetStructure: true,
        preserveTargetValues: true,
        preserveFormulas: true
      });
      expect(visualPlan.counts.referenceStyleSuggestions).toBeGreaterThan(0);
      expect(visualPlan.referenceStyleSuggestions).toEqual(expect.arrayContaining([
        expect.objectContaining({
          id: "reference_style.header",
          referenceSheetName: "May 2026",
          preserveTargetValues: true,
          preserveFormulas: true
        }),
        expect.objectContaining({ id: "reference_style.columns_by_role" })
      ]));
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations.some((operation) => operation.kind === "range.write_values" || operation.kind === "range.write_formulas")).toBe(false);
    });

  it("keeps visual readability print and presentation suggestions preview-only", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Make Apr 2026 print ready",
        mode: "preview_update",
        intent: { action: "improve_visual_readability" },
        target: { sheetName: "Apr 2026" },
        values: {
          visualReadability: {
            presentationMode: "print_ready"
          }
        }
      });
      const applied = await agent.run({
        request: "Apply visual readability preview",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      const visualPlan = (preview.answer as any).visualPlan;
      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).defaults.presentationMode).toBe("print_ready");
      expect(visualPlan.counts.printSuggestions).toBeGreaterThan(0);
      expect(visualPlan.printSuggestions).toEqual(expect.arrayContaining([
        expect.objectContaining({ id: "print.orientation", value: "landscape" }),
        expect.objectContaining({ id: "print.fit_to_width" }),
        expect.objectContaining({ id: "print.repeat_header", target: "A1:AG1" })
      ]));
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations.every((operation) => !String(operation.kind).includes("print"))).toBe(true);
    });

  it("verifies formula preservation when applying visual readability to formula ranges", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_visual_formula_preservation");
      metadata.formulaRegions = [{ id: "formula:apr-payment-variance", sheetName: "Apr 2026", range: "I2:I244", formulaCount: 243 }];
      agent.metadataCache.set(metadata);

      const preview = await agent.run({
        request: "Make Apr 2026 easier to read",
        mode: "preview_update",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "improve_visual_readability" },
        target: { sheetName: "Apr 2026" }
      });
      const applied = await agent.run({
        request: "Apply visual readability preview",
        mode: "apply_update",
        workbookContextId: metadata.workbookContextId,
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).visualPlan.preservation.formulaRanges).toEqual(["I2:I244"]);
      expect(applied.status).toBe("SUCCESS");
      expect((applied.answer as any).formulaPreservation).toMatchObject({
        checkedRanges: ["I2:I244"],
        formulasChanged: 0,
        unchanged: true
      });
      expect((applied.answer as any).formulaPreservation.formulasChecked).toBeGreaterThan(0);
      expect((applied.answer as any).telemetry).toMatchObject({ visualReadabilityApply: true, formulasChanged: 0 });
      expect(runtime.readBatchCount).toBe(2);
      expect(runtime.writeBatchCount).toBe(1);
      expect(runtime.lastWriteOperations.some((operation) => operation.kind === "range.write_values" || operation.kind === "range.write_formulas")).toBe(false);
    });

  it("fails visual readability apply when formula preservation changes are detected", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.mutateFormulaReadsAfterWrite = true;
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_visual_formula_preservation_failure");
      metadata.formulaRegions = [{ id: "formula:apr-payment-variance", sheetName: "Apr 2026", range: "I2:I244", formulaCount: 243 }];
      agent.metadataCache.set(metadata);

      const preview = await agent.run({
        request: "Make Apr 2026 easier to read",
        mode: "preview_update",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "improve_visual_readability" },
        target: { sheetName: "Apr 2026" }
      });
      const applied = await agent.run({
        request: "Apply visual readability preview",
        mode: "apply_update",
        workbookContextId: metadata.workbookContextId,
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect(applied.status).toBe("VALIDATION_FAILED");
      expect((applied.answer as any).formulaPreservation).toMatchObject({
        checkedRanges: ["I2:I244"],
        unchanged: false
      });
      expect((applied.answer as any).formulaPreservation.formulasChanged).toBeGreaterThan(0);
      expect(applied.warnings.join(" ")).toContain("Formula preservation failed");
    });

  it("asks for a header range when visual readability structure is ambiguous", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Make this report easier to read",
        mode: "preview_update",
        intent: { action: "improve_visual_readability" },
        target: { sheetName: "Report" }
      });

      expect(preview.status).toBe("NEEDS_INPUT");
      expect(preview.summary).toContain("Could not confidently detect a header row");
      expect(preview.nextAction).toBe("ask_user");
      expect(preview.warnings.join(" ")).toContain("No visual styling operations were prepared");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("blocks oversized visual readability targets before compiling operations", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Make this huge sheet easier to read",
        mode: "preview_update",
        intent: { action: "improve_visual_readability" },
        target: { sheetName: "Data", range: "A1:XFD200" }
      });

      expect(preview.status).toBe("VALIDATION_FAILED");
      expect(preview.summary).toContain("target is too large");
      expect(preview.nextAction).toBe("ask_user");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("blocks visual readability on hidden sheets by default", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_visual_hidden_sheet");
      metadata.sheets.push({
        id: "sheet:hidden",
        name: "Hidden Data",
        index: 99,
        usedRange: "A1:B4",
        rowCount: 4,
        columnCount: 2,
        isHidden: true,
        kind: "transaction",
        headers: [{
          id: "header:hidden",
          sheetName: "Hidden Data",
          row: 1,
          range: "A1:B1",
          confidence: 0.95,
          columns: [
            { name: "Date", normalizedName: "date", inferredType: "date", role: "date", importance: 0.9, index: 0, letter: "A" },
            { name: "Status", normalizedName: "status", inferredType: "status", role: "status", importance: 0.9, index: 1, letter: "B" }
          ]
        }],
        tableIds: [],
        sectionIds: [],
        summaryBlockIds: [],
        formulaRegionIds: []
      });
      agent.metadataCache.set(metadata);

      const preview = await agent.run({
        request: "Make hidden data easier to read",
        mode: "preview_update",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "improve_visual_readability" },
        target: { sheetName: "Hidden Data" }
      });

      expect(preview.status).toBe("VALIDATION_FAILED");
      expect(preview.summary).toContain("sheet is hidden");
      expect(preview.nextAction).toBe("ask_user");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("skips visual readability rules that overlap existing styled summary areas", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_visual_existing_style");
      metadata.summaryBlocks.push({
        id: "summary:data-title",
        sheetName: "Data",
        range: "A1:D1",
        labels: ["Transactions"],
        confidence: 0.95
      });
      agent.metadataCache.set(metadata);

      const preview = await agent.run({
        request: "Make Data easier to read",
        mode: "preview_update",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "improve_visual_readability" },
        target: { sheetName: "Data" }
      });
      const visualPlan = (preview.answer as any).visualPlan;

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).detected.existingStyleRanges).toEqual(["A1:D1"]);
      expect(visualPlan.skipped).toEqual(expect.arrayContaining([
        expect.objectContaining({ ruleId: "layout.header_style", reason: expect.stringContaining("protected summary/template style area") }),
        expect.objectContaining({ ruleId: "layout.header_alignment", reason: expect.stringContaining("protected summary/template style area") })
      ]));
      expect(visualPlan.operationCount).toBeGreaterThan(0);
      expect(preview.changes.some((change) => change.range === "A1:D1")).toBe(true);
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("supports strict style preservation when callers want all existing style protected", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_visual_strict_style");
      metadata.summaryBlocks.push({
        id: "summary:data-title",
        sheetName: "Data",
        range: "A1:D1",
        labels: ["Transactions"],
        confidence: 0.95
      });
      agent.metadataCache.set(metadata);

      const preview = await agent.run({
        request: "Make Data easier to read but preserve every existing style",
        mode: "preview_update",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "improve_visual_readability" },
        target: { sheetName: "Data" },
        values: {
          visualReadability: {
            stylePreservationMode: "strict"
          }
        }
      });
      const visualPlan = (preview.answer as any).visualPlan;

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).defaults.stylePreservationMode).toBe("strict");
      expect(visualPlan.skipped).toEqual(expect.arrayContaining([
        expect.objectContaining({ ruleId: "layout.header_style", reason: expect.stringContaining("protected summary/template style area") })
      ]));
      expect(visualPlan.skipped.some((skip: any) => skip.ruleId === "column.A.width")).toBe(false);
      expect(visualPlan.operationCount).toBeGreaterThan(0);
    });

  it("does not invite apply when a visual readability preview compiles zero operations", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_visual_zero_operations");
      const dataSheet = metadata.sheets.find((sheet) => sheet.name === "Data")!;
      dataSheet.tableIds = [];
      dataSheet.headers = [{
        id: "header:Data:1",
        sheetName: "Data",
        row: 1,
        range: "A1:D1",
        confidence: 0.95,
        columns: [
          { name: "Date", normalizedName: "date", inferredType: "date", role: "date", importance: 0.97, index: 0, letter: "A" },
          { name: "Account", normalizedName: "account", inferredType: "text", role: "account", importance: 0.82, index: 1, letter: "B" },
          { name: "Amount", normalizedName: "amount", inferredType: "currency", role: "amount", importance: 0.99, index: 2, letter: "C" },
          { name: "Status", normalizedName: "status", inferredType: "text", role: "status", importance: 0.9, index: 3, letter: "D" }
        ]
      }];
      metadata.tables = [];
      metadata.summaryBlocks.push({
        id: "summary:data-all",
        sheetName: "Data",
        range: "A1:D4",
        labels: ["Protected Data"],
        confidence: 0.95
      });
      agent.metadataCache.set(metadata);

      const preview = await agent.run({
        request: "Make Data easier to read but preserve protected layout",
        mode: "preview_update",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "improve_visual_readability" },
        target: { sheetName: "Data" },
        values: {
          visualReadability: {
            styleDepth: "basic"
          }
        }
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect(preview.nextAction).toBe("answer_now");
      expect(preview.agentInstruction).toContain("Do not call apply_update");
      expect((preview.metrics as any).operationCount).toBe(0);
      expect(preview.warnings.join(" ")).toContain("No apply-ready visual operations");
    });

  it("suggests grouped headers for wide visual readability previews without applying structural edits by default", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_visual_grouped_header_suggestion");
      agent.metadataCache.set(metadata);

      const preview = await agent.run({
        request: "Make Apr 2026 easier to read with modern grouped headers if useful",
        mode: "preview_update",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "improve_visual_readability" },
        target: { sheetName: "Apr 2026" }
      });
      const applied = await agent.run({
        request: "Apply safe visual readability preview",
        mode: "apply_update",
        workbookContextId: metadata.workbookContextId,
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).groupedHeaderSuggestion).toMatchObject({
        kind: "grouped_header_suggestion",
        requiresStructuralPreview: true,
        defaultApplyBehavior: "suggest_only"
      });
      expect((preview.answer as any).groupedHeaderSuggestion.operationsNeeded).toEqual(expect.arrayContaining([
        "insert_rows",
        "merge_range",
        "write_styles_many"
      ]));
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations.some((operation) => operation.kind === "range.insert_rows" || operation.kind === "range.merge" || operation.kind === "range.reorder_columns")).toBe(false);
    });

  it("compiles grouped header styling into one structural preview instead of named-range or value-only patches", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_grouped_header_regression");
      agent.metadataCache.set(metadata);

      const preview = await agent.run({
        request: "Create merged group header cells for Apr 2026 with a 2-layer grouped header and color bands.",
        mode: "preview_update",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "grouped_header" },
        target: { sheetName: "Apr 2026" },
        values: {
          groupedHeader: {
            groups: [
              { label: "Transactions", startColumn: "A", endColumn: "N", fillColor: "#1F4E78", headerFillColor: "#D9EAF7" },
              { label: "Invoices", startColumn: "O", endColumn: "AE", fillColor: "#548235", headerFillColor: "#E2EFDA" },
              { label: "Summary", startColumn: "AF", endColumn: "AJ", fillColor: "#8064A2", headerFillColor: "#EDE7F6" }
            ]
          }
        }
      });
      const applied = await agent.run({
        request: "Apply grouped header preview",
        mode: "apply_update",
        workbookContextId: metadata.workbookContextId,
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any)).toMatchObject({
        kind: "grouped_header_preview",
        sheetName: "Apr 2026",
        headerRow: 1,
        groupRow: 1,
        shiftedHeaderRow: 2,
        preservesExistingHeaderLabels: true
      });
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations.map((operation) => operation.kind)).toEqual([
        "range.insert_rows",
        "range.write_values_many",
        "range.merge",
        "range.merge",
        "range.merge",
        "range.write_styles_many"
      ]);
      expect(runtime.lastBatchOperations[0]).toMatchObject({
        kind: "range.insert_rows",
        target: { sheetName: "Apr 2026", address: "A1:AJ1" }
      });
      expect(runtime.lastBatchOperations[1]).toMatchObject({
        kind: "range.write_values_many",
        entries: [
          expect.objectContaining({ target: expect.objectContaining({ sheetName: "Apr 2026", address: "A1:A1" }), values: [["Transactions"]] }),
          expect.objectContaining({ target: expect.objectContaining({ sheetName: "Apr 2026", address: "O1:O1" }), values: [["Invoices"]] }),
          expect.objectContaining({ target: expect.objectContaining({ sheetName: "Apr 2026", address: "AF1:AF1" }), values: [["Summary"]] })
        ]
      });
      const styleOperation = runtime.lastBatchOperations.at(-1) as any;
      expect(styleOperation.entries.some((entry: any) => entry.target.address === "A1:N1" && entry.style.fillColor === "#1F4E78")).toBe(true);
      expect(styleOperation.entries.some((entry: any) => entry.target.address === "A2:N2" && entry.style.fillColor === "#D9EAF7")).toBe(true);
    });

  it("keeps grouped header row darker than row 2 when matching header styling is requested", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_grouped_header_color_hierarchy");
      agent.metadataCache.set(metadata);

      const preview = await agent.run({
        request: "Apply matching header fill color and font styling to row 1 grouped header to match the style already applied to row 2.",
        mode: "preview_update",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "format_range" },
        target: { sheetName: "Data", range: "A1:D1" }
      });
      const applied = await agent.run({
        request: "Apply grouped header color hierarchy.",
        mode: "apply_update",
        workbookContextId: metadata.workbookContextId,
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any)).toMatchObject({
        kind: "style_preview",
        sheetName: "Data",
        range: "A1:D1",
        style: {
          fillColor: "#1A3C6E",
          fontColor: "#FFFFFF",
          fontBold: true,
          horizontalAlignment: "center"
        }
      });
      expect(preview.warnings.join(" ")).toContain("visually distinct from row 2");
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations[0]).toMatchObject({
        kind: "range.write_styles",
        target: { sheetName: "Data", address: "A1:D1" },
        style: {
          fillColor: "#1A3C6E",
          fontColor: "#FFFFFF",
          fontBold: true,
          horizontalAlignment: "center"
        }
      });
    });

  it("accepts target.address as an exact format range without expanding to the used range", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_target_address_alias");
      agent.metadataCache.set(metadata);

      const preview = await agent.run({
        request: "Apply dark blue fill to row 1 only A1:D1",
        mode: "preview_update",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "format_range" },
        values: {
          patches: [{
            target: { sheetName: "Data", range: "A1:D1" },
            style: { fillColor: "#1A3C6E", fontColor: "#FFFFFF", fontBold: true, horizontalAlignment: "center" }
          }]
        }
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect(preview.changes?.[0]).toMatchObject({ sheetName: "Data", range: "A1:D1" });
      expect(preview.changes?.[0]?.range).not.toBe("A1:D4");
    });

  it("accepts OpenCode grouped header column arrays and does not trigger broad scope guard", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_grouped_header_columns_shape");
      agent.metadataCache.set(metadata);

      const preview = await agent.run({
        request: "Preview a grouped_header workflow for Transactions table. Add a higher-level grouped header row above the existing column headers and apply it to all rows of the table without changing data rows.",
        mode: "preview_update",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "grouped_header" },
        target: { sheetName: "Data", tableName: "Transactions" },
        values: {
          stylePreservationMode: "none",
          groupedHeader: {
            groups: [
              { label: "Timeline", columns: ["A"] },
              { label: "Account", columns: ["B"] },
              { label: "Financial", columns: ["C"] },
              { label: "Workflow", columns: ["D"] }
            ]
          }
        }
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe("grouped_header_preview");
      expect((preview.answer as any).groups.map((group: any) => [group.label, group.startColumn, group.endColumn])).toEqual([
        ["Timeline", "A", "A"],
        ["Account", "B", "B"],
        ["Financial", "C", "C"],
        ["Workflow", "D", "D"]
      ]);
      expect((preview.answer as any).groups[0].fillColor).toBe("#1A3C6E");
      expect((preview.answer as any).groups[0].headerFillColor).toBe("#D9EAF7");
      expect((preview.answer as any).operationCount).toBeGreaterThan(0);
      expect((preview.answer as any).kind).not.toBe("broad_mutation_scope_guard");
      expect(preview.summary).not.toContain("15,015");
      expect((preview.metrics as any).workflowKind).toBe("grouped_header_preview");
    });

  it("accepts grouped header ranges as group spans", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_grouped_header_range_shape");
      agent.metadataCache.set(metadata);

      const preview = await agent.run({
        request: "Preview a two-level grouped header for Apr 2026 using range-shaped group spans.",
        mode: "preview_update",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "grouped_header" },
        target: { sheetName: "Apr 2026" },
        values: {
          groupedHeader: {
            groups: [
              { label: "Transactions", range: "A:N" },
              { label: "Invoices", range: "O:AE" },
              { label: "Summary", range: "AF:AJ" }
            ]
          }
        }
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).groups.map((group: any) => [group.label, group.startColumn, group.endColumn])).toEqual([
        ["Transactions", "A", "N"],
        ["Invoices", "O", "AE"],
        ["Summary", "AF", "AJ"]
      ]);
      expect((preview.answer as any).operationCount).toBe(6);
    });

  it("rejects preview_update calls that reuse a stale operationId from another workflow", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const visualPreview = await agent.run({
        request: "Make this sheet easier to read",
        mode: "preview_update",
        intent: { action: "improve_visual_readability" },
        target: { sheetName: "Data" }
      });
      const reused = await agent.run({
        request: "Apply the grouped_header preview to all rows of Transactions.",
        mode: "preview_update",
        operationId: visualPreview.operationId,
        confirmationToken: visualPreview.confirmationToken,
        intent: { action: "grouped_header" }
      });
      const status = await agent.run({
        request: "Check grouped_header preview status.",
        mode: "operation_status",
        operationId: visualPreview.operationId,
        intent: { action: "grouped_header" }
      });

      expect(visualPreview.status).toBe("PREVIEW_READY");
      expect(reused.status).toBe("VALIDATION_FAILED");
      expect((reused.answer as any).kind).toBe("invalid_preview_operation_reuse");
      expect(reused.warnings.join(" ")).toContain("operationId from a different preview");
      expect((status.answer as any).workflowKind).toBe("visual_readability_preview");
      expect(status.warnings.join(" ")).toContain("belongs to visual_readability_preview");
      expect(status.warnings.join(" ")).toContain("grouped_header_preview");
    });

  it("requires structured values for write previews even when request text includes rows", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: `Add 2 mockup rows to the Report sheet at range Report!A3:B4.

Data rows:
["A", "B"]
["C", "D"]`,
        mode: "preview_update",
        intent: { action: "write_values" },
        target: { sheetName: "Report", range: "Report!A3:B4" }
      });

      expect(preview.status).toBe("NEEDS_INPUT");
      expect(preview.summary).toContain("Preview needs structured values");
      expect(preview.summary).toContain("rows embedded in request text");
      expect(preview.nextAction).toBe("ask_user");
      expect(preview.warnings).toContain("Direct cell/range updates must be supplied as values.patches. Specialized workflows such as table appends may use their dedicated structured values.");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("reports missing structured values before target ambiguity for write previews", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Write rows into the report",
        mode: "preview_update",
        intent: { action: "write_values" }
      });

      expect(preview.status).toBe("NEEDS_INPUT");
      expect(preview.summary).toContain("Preview needs structured values");
      expect(preview.nextAction).toBe("ask_user");
      expect(preview.candidates?.length).toBeGreaterThan(0);
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("requires preview confirmation token before apply", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Update Data B2",
        mode: "preview_update",
        target: { sheetName: "Data", range: "B2" },
        values: valuePatch("Data", "B2", [[999]])
      });
      const missingToken = await agent.run({
        request: "Apply update",
        mode: "apply_update",
        operationId: preview.operationId
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect(preview.confirmationToken).toBeTruthy();
      expect(preview.taskOutcome).toBe("preview_ready");
      expect(preview.maxRecommendedFollowupCalls).toBe(1);
      expect(preview.requiredFollowup).toMatchObject({
        mode: "apply_update",
        nextAction: "call_apply_update",
        operationId: preview.operationId
      });
      expect(preview.agentInstruction).toContain("Do not rediscover workbook context");
      expect(missingToken.status).toBe("NEEDS_INPUT");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("rejects wrong confirmation tokens without applying", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Update Data B2",
        mode: "preview_update",
        target: { sheetName: "Data", range: "B2" },
        values: valuePatch("Data", "B2", [[999]])
      });
      const wrongToken = await agent.run({
        request: "Apply update",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: "confirm_wrong"
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect(wrongToken.status).toBe("NEEDS_INPUT");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("returns terminal apply output on repeated apply without rerunning the batch", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Update Data B2",
        mode: "preview_update",
        target: { sheetName: "Data", range: "B2" },
        values: valuePatch("Data", "B2", [[999]])
      });
      const first = await agent.run({
        request: "Apply update",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });
      const second = await agent.run({
        request: "Apply update again",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });
      const status = await agent.run({
        request: "Check applied operation",
        mode: "operation_status",
        operationId: preview.operationId
      });
      const cancelApplied = await agent.run({
        request: "Cancel applied operation",
        mode: "cancel_operation",
        operationId: preview.operationId
      });

      expect(first.status).toBe("SUCCESS");
      expect(second).toMatchObject({ status: "SUCCESS", operationId: preview.operationId });
      expect(runtime.writeBatchCount).toBe(1);
      expect((status.answer as any).applyStatus).toBe("applied");
      expect(cancelApplied.status).toBe("VALIDATION_FAILED");
    });

  it("invalidates workbook context and stored result handles after successful apply", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_apply_invalidate");
      agent.metadataCache.set(metadata);

      const read = await agent.run({
        request: "Read first transaction rows",
        mode: "answer",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "read_values" },
        target: { tableName: "Transactions" },
        values: { rowLimit: 2 }
      });
      const resultUri = (read.answer as any).resultUri as string;
      const resultId = resultUri.split("/").pop()!;
      const preview = await agent.run({
        request: "Update Data B2",
        mode: "preview_update",
        workbookContextId: metadata.workbookContextId,
        target: { sheetName: "Data", range: "B2" },
        values: valuePatch("Data", "B2", [[999]])
      });
      const applied = await agent.run({
        request: "Apply update",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(applied.status).toBe("SUCCESS");
      expect((applied.answer as any).cacheImpact).toMatchObject({
        cacheAction: "updated_from_patch",
        updatedFacets: ["values"],
        freshness: {
          status: "mostly_fresh",
          staleFacets: expect.arrayContaining(["aggregates", "formulaResults"]),
          staleRanges: expect.arrayContaining(["Data!B2"])
        },
        journalEntry: {
          operationId: preview.operationId,
          affectedRanges: ["Data!B2"],
          affectedFacets: expect.arrayContaining(["values", "aggregates", "formulaResults"]),
          invalidatedFacets: expect.arrayContaining(["aggregates", "formulaResults"]),
          cacheAction: "updated_from_patch"
        }
      });
      expect((applied.answer as any).cacheImpact.freshness.staleFacets).not.toContain("values");
      expect(applied.operationJournalRef).toMatchObject({
        workbookContextId: metadata.workbookContextId,
        operationId: preview.operationId,
        contextVersion: expect.any(Number),
        appliedAt: expect.any(Number)
      });
      expect(applied.invalidatedContextIds).not.toContain(metadata.workbookContextId);
      expect(applied.invalidatedResourceUris).toContain(resultUri);
      expect(agent.metadataCache.getByContextId(metadata.workbookContextId)).toBeDefined();
      expect(agent.metadataCache.getContextState(metadata.workbookContextId)?.freshness.staleFacets).toEqual(expect.arrayContaining(["aggregates", "formulaResults"]));
      expect(agent.metadataCache.getContextState(metadata.workbookContextId)?.freshness.staleFacets).not.toContain("values");
      expect(agent.metadataCache.getOptimisticValue(metadata.workbookContextId, "Data", "B2")).toMatchObject({
        range: "B2",
        value: 999,
        operationId: preview.operationId
      });
      expect(applied.contextFreshness).toMatchObject({
        status: "mostly_fresh",
        staleFacets: expect.arrayContaining(["aggregates", "formulaResults"])
      });
      expect(applied.contextFreshness?.staleFacets).not.toContain("values");
      expect(agent.getResultResource(resultId) as any).toMatchObject({ ok: false });
      expect(agent.getContextResource(metadata.workbookContextId) as any).toMatchObject({ ok: true });
    });

  it("fully invalidates workbook context after rebuild-context mutations", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_rebuild_context_apply");
      agent.metadataCache.set(metadata);

      const preview = await agent.run({
        request: "Delete column B",
        mode: "preview_update",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "delete_columns" },
        target: { sheetName: "Data", range: "B:B" }
      });
      const applied = await agent.run({
        request: "Apply delete column",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect(applied.status).toBe("SUCCESS");
      expect((applied.answer as any).updateRisk).toMatchObject({ cacheAction: "rebuild_context" });
      expect(applied.invalidatedContextIds).toContain(metadata.workbookContextId);
      expect(agent.getContextResource(metadata.workbookContextId) as any).toMatchObject({ ok: false });
    });

  it("reports and cancels previewed operations without workbook readiness", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Update Data B2",
        mode: "preview_update",
        target: { sheetName: "Data", range: "B2" },
        values: valuePatch("Data", "B2", [[999]])
      });
      const status = await agent.run({
        request: "Check operation",
        mode: "operation_status",
        operationId: preview.operationId
      });
      const cancelled = await agent.run({
        request: "Cancel operation",
        mode: "cancel_operation",
        operationId: preview.operationId
      });
      const applyAfterCancel = await agent.run({
        request: "Apply cancelled operation",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(status.status).toBe("SUCCESS");
      expect((status.answer as any).applyStatus).toBe("previewed");
      expect(cancelled.status).toBe("SUCCESS");
      expect(applyAfterCancel.status).toBe("NOT_FOUND");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("previews and applies grouped range patches as one batch operation", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Update related pricing zones",
        mode: "preview_update",
        values: {
          patches: [
            { target: { sheetName: "Data", range: "B2:C2" }, values: [[9000, 9000]], reason: "Zone A price" },
            { target: { sheetName: "Data", range: "B3:C3" }, values: [[6000, 6000]], reason: "Zone B price" }
          ]
        }
      });
      const applied = await agent.run({
        request: "Apply grouped pricing update",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe("multi_range_preview");
      expect((preview.answer as any).patchCount).toBe(2);
      expect(preview.summary).toContain("Apply this grouped preview once");
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.writeBatchCount).toBe(1);
      expect(runtime.lastBatchOperations).toHaveLength(1);
      expect(runtime.lastBatchOperations[0]?.kind).toBe("range.write_values_many");
      expect((runtime.lastBatchOperations[0] as any).entries.map((entry: any) => entry.target.address)).toEqual(["B2:C2", "B3:C3"]);
    });

  it("validates grouped patches against dropdown options before previewing", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_valid_dropdown_patch");
      agent.metadataCache.set(metadata);

      const preview = await agent.run({
        request: "Set the selected status to Closed.",
        mode: "preview_update",
        workbookContextId: metadata.workbookContextId,
        values: valuePatch("Data", "D2", [["Closed"]])
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).updateRisk).toMatchObject({
        cacheRisk: "low",
        safetyRisk: "low",
        cacheAction: "update_cache",
        invalidatedFacets: expect.arrayContaining(["values", "aggregates", "formulaResults"]),
        preservedFacets: expect.arrayContaining(["schema", "headers", "fieldContext", "validation"]),
        requiresRefreshBeforeNextMutation: false
      });
      expect((preview.answer as any).validationChecks[0]).toMatchObject({
        field: "Status",
        allowedValues: ["Open", "Closed", "Pending"],
        proposedValues: ["Closed"],
        invalidValues: []
      });
      expect(preview.contextUsed).toMatchObject({
        strategy: "mutation",
        levelUsed: 3,
        stagesPlanned: expect.arrayContaining(["metadata", "schema", "field_context", "validation"]),
        stagesUsed: expect.arrayContaining(["field_context", "audit_facets", "target_resolution", "preview_proof"]),
        stopReason: expect.stringContaining("safe preview"),
        included: expect.arrayContaining(["validation", "field_context", "preview"])
      });
      expect(preview.warnings.join(" ")).toContain("validated Status");
    });

  it("blocks grouped patches with values outside dropdown options", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_invalid_dropdown_patch");
      agent.metadataCache.set(metadata);

      const preview = await agent.run({
        request: "Set the selected status to Done.",
        mode: "preview_update",
        workbookContextId: metadata.workbookContextId,
        values: valuePatch("Data", "D2", [["Done"]])
      });

      expect(preview.status).toBe("VALIDATION_FAILED");
      expect((preview.answer as any)).toMatchObject({
        kind: "patch_validation_failed",
        code: "VALUE_NOT_IN_DROPDOWN_OPTIONS",
        field: "Status",
        proposedValues: ["Done"],
        invalidValues: ["Done"],
        allowedValues: ["Open", "Closed", "Pending"]
      });
      expect(preview.operationId).toBeUndefined();
      expect(preview.warnings.join(" ")).toContain("VALUE_NOT_IN_DROPDOWN_OPTIONS");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("auto-applies small grouped range patches with explicit patch targets", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Update related pricing zones",
        values: {
          patches: [
            { target: { sheetName: "Data", range: "B2:C2" }, values: [[9000, 9000]], reason: "Zone A price" },
            { target: { sheetName: "Data", range: "B3:C3" }, values: [[6000, 6000]], reason: "Zone B price" }
          ]
        }
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.taskOutcome).toBe("apply_complete");
      expect(result.telemetry.autoApplied).toBe(true);
      expect(result.telemetry.safetyDecision).toBe("auto_apply:scoped_value_edit");
      expect(runtime.writeBatchCount).toBe(1);
      expect(runtime.lastBatchOperations[0]?.kind).toBe("range.write_values_many");
      expect((runtime.lastBatchOperations[0] as any).entries.map((entry: any) => entry.target.address)).toEqual(["B2:C2", "B3:C3"]);
    });

  it("auto-applies non-adjacent row updates from one prompt as one grouped patch batch", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Update row 19 maintenance fields and row 20 owner funding fields",
        mode: "auto",
        values: {
          patches: [
            { target: { sheetName: "Data", range: "C19:E19" }, values: [["700-5229", "maintenance note", "maintenance"]], reason: "row 19 maintenance update" },
            { target: { sheetName: "Data", range: "D20:E20" }, values: [["Owner cash top-up", "owner_cash_topup"]], reason: "row 20 owner funding label" },
            { target: { sheetName: "Data", range: "H20:J20" }, values: [[10000, 0, "Owner fund top-up"]], reason: "row 20 owner funding amounts" }
          ]
        }
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.taskOutcome).toBe("apply_complete");
      expect(result.telemetry.autoApplied).toBe(true);
      expect(runtime.writeBatchCount).toBe(1);
      expect(runtime.lastBatchOperations).toHaveLength(1);
      expect(runtime.lastBatchOperations[0]?.kind).toBe("range.write_values_many");
      expect((runtime.lastBatchOperations[0] as any).entries.map((entry: any) => entry.target.address)).toEqual(["C19:E19", "D20:E20", "H20:J20"]);
    });

  it("auto-applies a bounded missing dropdown source-list option as one value write", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Add owner_cash_topup to the transaction type source list.",
        mode: "auto",
        intent: { action: "write_values" },
        target: { sheetName: "Data", range: "A4" },
        values: valuePatch("Data", "A4", [["owner_cash_topup"]])
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.taskOutcome).toBe("apply_complete");
      expect(result.maxRecommendedFollowupCalls).toBe(0);
      expect(result.telemetry.autoApplied).toBe(true);
      expect(runtime.writeBatchCount).toBe(1);
      expect(runtime.lastBatchOperations).toHaveLength(1);
      expect(runtime.lastBatchOperations[0]).toMatchObject({
        kind: "range.write_values_many",
        entries: [{ target: { sheetName: "Data", address: "A4" }, values: [["owner_cash_topup"]] }]
      });
    });

  it("auto-applies exact transfer match updates as one backend-owned grouped operation", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "all transfer to WITSARUT is truck 71-4653",
        mode: "auto",
        target: { sheetName: "May 2026" }
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.taskOutcome).toBe("apply_complete");
      expect(result.maxRecommendedFollowupCalls).toBe(0);
      expect((result.answer as any).kind).toBe("match_update_result");
      expect((result.answer as any).matchColumn.name).toBe("Transfer From/To");
      expect((result.answer as any).updateColumn.name).toBe("Truck ID");
      expect((result.answer as any).predicate).toEqual({ operator: "contains", value: "WITSARUT" });
      expect((result.answer as any).updateValue).toBe("71-4653");
      expect((result.answer as any).matchedRows).toEqual([25, 27, 28, 31, 33, 37, 38, 39, 40, 44, 48]);
      expect(result.agentInstruction).toContain("Do not fetch full rows");
      expect(runtime.writeBatchCount).toBe(1);
      expect(runtime.lastBatchOperations).toHaveLength(1);
      expect(runtime.lastBatchOperations[0]?.kind).toBe("range.write_values_many");
      expect((runtime.lastBatchOperations[0] as any).entries.map((entry: any) => entry.target.address)).toEqual(["C25", "C27", "C28", "C31", "C33", "C37", "C38", "C39", "C40", "C44", "C48"]);
    });

  it("previews exact transfer match updates without full-row handles when auto apply is disabled", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "all transfer to WITSARUT is truck 71-4653",
        mode: "preview_update",
        target: { sheetName: "May 2026" }
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe("match_update_preview");
      expect((preview.answer as any).matchedRowCount).toBe(11);
      expect((preview.answer as any).matchedRows).toEqual([25, 27, 28, 31, 33, 37, 38, 39, 40, 44, 48]);
      expect((preview.answer as any).fullResultUri).toBeUndefined();
      expect(preview.maxRecommendedFollowupCalls).toBe(1);
      expect(preview.agentInstruction).toContain("Do not fetch full rows");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("inherits parent sheet for grouped cell-only patch targets", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Set WITSARUT truck cells",
        mode: "auto",
        target: { sheetName: "May 2026" },
        values: {
          patches: [
            { target: { range: "C31" }, values: [["71-4653"]] },
            { target: { range: "C33" }, values: [["71-4653"]] }
          ]
        }
      });

      expect(result.status).toBe("SUCCESS");
      expect(runtime.writeBatchCount).toBe(1);
      expect(runtime.lastBatchOperations[0]?.kind).toBe("range.write_values_many");
      expect((runtime.lastBatchOperations[0] as any).entries.map((entry: any) => entry.target)).toEqual([
        { workbookId, sheetName: "May 2026", address: "C31" },
        { workbookId, sheetName: "May 2026", address: "C33" }
      ]);
    });

  it("auto-applies semantic section row and column patches after resolving exact cells", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_semantic_patch");
      metadata.sheets.push({
        id: "sheet:Vendor Propose",
        name: "Vendor Propose",
        index: 99,
        usedRange: "A1:P6",
        rowCount: 6,
        columnCount: 16,
        kind: "unknown",
        headers: [],
        tableIds: [],
        sectionIds: ["section:quote:main"],
        summaryBlockIds: [],
        formulaRegionIds: []
      });
      metadata.sections.push({
        id: "section:quote:main",
        sheetName: "Vendor Propose",
        label: "YLTH_CTG_Zone BKK TT _Y2026",
        kind: "table-like",
        range: "A1:P6",
        headerRange: "B3:P3",
        headerRow: 3,
        columns: [
          { name: "Item No.", normalizedName: "item_no", inferredType: "number", role: "identifier", importance: 0.7, index: 1, letter: "B" },
          { name: "Transport Mode", normalizedName: "transport_mode", inferredType: "text", role: "dimension", importance: 0.7, index: 2, letter: "C" },
          { name: "Route", normalizedName: "route", inferredType: "text", role: "description", importance: 0.95, index: 3, letter: "D" },
          { name: "Truck Available (Truck/day)", normalizedName: "truck_available_truckday", inferredType: "number", role: "measure", importance: 0.9, index: 14, letter: "O" },
          { name: "Vendor Propose (THB/trip)", normalizedName: "vendor_propose_thbtrip", inferredType: "currency", role: "amount", importance: 0.9, index: 15, letter: "P" }
        ],
        labels: ["YLTH_CTG_Zone BKK TT _Y2026"],
        rowCount: 6,
        columnCount: 16,
        nonEmptyCellCount: 45,
        confidence: 0.9
      });
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "Set the Klongtoey route truck available to 10 and vendor propose to 7500.",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "write_values" },
        values: {
          semanticPatches: [
            {
              sectionId: "section:quote:main",
              rowMatch: { column: "Route", value: "Klongtoey" },
              columnMatch: "Truck Available",
              value: 10
            },
            {
              sectionId: "section:quote:main",
              rowMatch: { column: "Route", value: "Klongtoey" },
              columnMatch: "Vendor Propose",
              value: 7500
            }
          ]
        }
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.taskOutcome).toBe("apply_complete");
      expect(result.maxRecommendedFollowupCalls).toBe(0);
      expect(result.telemetry.autoApplied).toBe(true);
      expect(result.telemetry.safetyDecision).toBe("auto_apply:scoped_value_edit");
      expect(result.telemetry.internalCallCount).toBeLessThanOrEqual(5);
      expect(runtime.writeBatchCount).toBe(1);
      expect(runtime.lastBatchOperations[0]?.kind).toBe("range.write_values_many");
      expect((runtime.lastBatchOperations[0] as any).entries.map((entry: any) => entry.target.address)).toEqual(["O4", "P4"]);
      const snapshotReads = runtime.snapshotRangesHistory.flat().map((range) => `${range.sheetName}!${range.address}`);
      expect(snapshotReads).toContain("Vendor Propose!A4:P6");
      expect(snapshotReads.filter((range) => range === "Vendor Propose!A4:P6")).toHaveLength(1);
    });

  it("asks for a better row anchor when semantic section patches match no row", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_semantic_patch_missing_row");
      metadata.sheets.push({
        id: "sheet:Vendor Propose",
        name: "Vendor Propose",
        index: 99,
        usedRange: "A1:P6",
        rowCount: 6,
        columnCount: 16,
        kind: "unknown",
        headers: [],
        tableIds: [],
        sectionIds: ["section:quote:main"],
        summaryBlockIds: [],
        formulaRegionIds: []
      });
      metadata.sections.push({
        id: "section:quote:main",
        sheetName: "Vendor Propose",
        label: "YLTH_CTG_Zone BKK TT _Y2026",
        kind: "table-like",
        range: "A1:P6",
        headerRange: "B3:P3",
        headerRow: 3,
        columns: [
          { name: "Route", normalizedName: "route", inferredType: "text", role: "description", importance: 0.95, index: 3, letter: "D" },
          { name: "Vendor Propose (THB/trip)", normalizedName: "vendor_propose_thbtrip", inferredType: "currency", role: "amount", importance: 0.9, index: 15, letter: "P" }
        ],
        labels: ["YLTH_CTG_Zone BKK TT _Y2026"],
        rowCount: 6,
        columnCount: 16,
        nonEmptyCellCount: 45,
        confidence: 0.9
      });
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "Set the Bangna route vendor propose to 7500.",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "write_values" },
        values: {
          semanticPatches: [{
            sectionId: "section:quote:main",
            rowMatch: { column: "Route", value: "Bangna" },
            columnMatch: "Vendor Propose",
            value: 7500
          }]
        }
      });

      expect(result.status).toBe("NEEDS_INPUT");
      expect(result.summary).toContain("could not match row");
      expect(result.nextAction).toBe("ask_user");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("does not auto-apply small writes that would overwrite detected header-like labels", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_header_overwrite_guard");
      metadata.sheets.push({
        id: "sheet:Vendor Propose",
        name: "Vendor Propose",
        index: 99,
        usedRange: "A1:P6",
        rowCount: 6,
        columnCount: 16,
        kind: "unknown",
        headers: [],
        tableIds: [],
        sectionIds: [],
        summaryBlockIds: [],
        formulaRegionIds: []
      });
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "Set O3=10, P3=7500, O4=10, P4=7500, O5=10, P5=8000 for the quotation routes",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "write_values" },
        target: { sheetName: "Vendor Propose", range: "O3:P5" },
        values: valuePatch("Vendor Propose", "O3:P5", [[10, 7500], [10, 7500], [10, 8000]])
      });

      expect(result.status).toBe("PREVIEW_READY");
      expect(result.taskOutcome).toBe("preview_ready");
      expect(result.telemetry.autoApplied).toBeUndefined();
      expect(result.telemetry.safetyDecision).toBe("manual_review:target_looks_like_header_or_clear");
      expect(result.warnings.join(" ")).toContain("header/title/reference");
      expect(result.changes?.[0]).toMatchObject({
        sheetName: "Vendor Propose",
        cell: "O3",
        before: "Truck Available\n(Truck/day)",
        after: 10
      });
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("does not auto-apply null or blank writes unless the request explicitly clears cells", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Update Data B2:C2 with the corrected values",
        intent: { action: "write_values" },
        target: { sheetName: "Data", range: "B2:C2" },
        values: valuePatch("Data", "B2:C2", [[null, 7500]])
      });

      expect(result.status).toBe("PREVIEW_READY");
      expect(result.taskOutcome).toBe("preview_ready");
      expect(result.telemetry.safetyDecision).toBe("manual_review:target_looks_like_header_or_clear");
      expect(result.warnings.join(" ")).toContain("blank/null writes require an explicit clear request");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("prefers sheet-qualified patch ranges over stale table hints", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Update the transaction cells",
        mode: "preview_update",
        values: {
          patches: [
            { target: { tableName: "Transactions", range: "'Data'!B2:C2" }, values: [["Checking", 9000]], reason: "Correct row values" }
          ]
        }
      });
      const applied = await agent.run({
        request: "Apply transaction cells",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations[0]?.kind).toBe("range.write_values_many");
      const entries = (runtime.lastBatchOperations[0] as any).entries;
      expect(entries).toHaveLength(1);
      expect(entries[0].target).toMatchObject({ sheetName: "Data", address: "B2:C2" });
    });

  it("serializes structured apply failure warnings instead of object placeholders", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.batchResultOverride = {
        ok: false,
        warnings: [{
          code: "OPERATION_FAILED",
          message: "Matrix dimensions do not match Data!B2:C2: expected 1 row(s) x 2 column(s), received 1 row(s) x 1 column(s).",
          target: { sheetName: "Data", address: "B2:C2" }
        }],
        error: {
          code: "OPERATION_FAILED",
          message: "Matrix dimensions do not match Data!B2:C2: expected 1 row(s) x 2 column(s), received 1 row(s) x 1 column(s)."
        }
      };
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Update Data B2:C2",
        mode: "preview_update",
        target: { sheetName: "Data", range: "B2:C2" },
        values: valuePatch("Data", "B2:C2", [["Checking", 9000]])
      });
      const applied = await agent.run({
        request: "Apply Data B2:C2",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(applied.status).toBe("VALIDATION_FAILED");
      expect(applied.warnings?.join("\n")).toContain("OPERATION_FAILED: Matrix dimensions do not match Data!B2:C2");
      expect(applied.warnings?.join("\n")).not.toContain("[object Object]");
    });

  it("summarizes large value previews without reading every before-cell", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);
    const metadata = createCachedMetadata("wbctx_large_preview");
    agent.metadataCache.set(metadata);
    const values = Array.from({ length: 21 }, (_row, rowIndex) =>
      Array.from({ length: 10 }, (_cell, columnIndex) => `R${rowIndex + 1}C${columnIndex + 1}`)
    );
    const readCountBefore = runtime.readBatchCount;

    const preview = await agent.run({
      request: "Write large filled matrix",
      mode: "preview_update",
      workbookContextId: metadata.workbookContextId,
      target: { sheetName: "Report", range: "A1:J21" },
      values: valuePatch("Report", "A1:J21", values)
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect(preview.changes).toHaveLength(1);
    expect((preview.changes?.[0]?.after as any).kind).toBe("range_write_summary");
    expect((preview.changes?.[0]?.after as any).cellCount).toBe(210);
    expect(runtime.readBatchCount).toBe(readCountBefore);
  });

  it("caps detailed grouped patch preview reads across small patches", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);
    const values = Array.from({ length: 1 }, () => Array.from({ length: 50 }, (_cell, index) => index));
    const preview = await agent.run({
      request: "Update related wide rows",
      mode: "preview_update",
      values: {
        patches: Array.from({ length: 5 }, (_patch, index) => ({
          target: { sheetName: "Data", range: `A${index + 1}:AX${index + 1}` },
          values,
          reason: `Patch ${index + 1}`
        }))
      }
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect(preview.changes?.length).toBeLessThanOrEqual(201);
  });

  it("applies a pending preview under the agent identity that created it", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      runtime.agentExecutionContext = { agentId: "agent_author", agentName: "Author", clientType: "mcp" };
      const preview = await agent.run({
        request: "Update Data B2",
        mode: "preview_update",
        target: { sheetName: "Data", range: "B2" },
        values: valuePatch("Data", "B2", [[999]])
      });

      runtime.agentExecutionContext = { agentId: "agent_other", agentName: "Other", clientType: "mcp" };
      const applied = await agent.run({
        request: "Apply update",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchRequest?.agentId).toBe("agent_author");
    });

  it("blocks unsafe grouped patches before creating a pending batch", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Update related cells",
        mode: "preview_update",
        values: {
          patches: [
            { target: { sheetName: "Data", range: "B2:C2" }, values: [[1, 2]] },
            { target: { sheetName: "Data", range: "B3:C3" }, values: [["=SUM(A1:A2)", 3]] }
          ]
        }
      });

      expect(result.status).toBe("VALIDATION_FAILED");
      expect(result.summary).toContain("formula-like");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("treats apply results without explicit ok as success when validation passes", async () => {
      const runtime = new FakeAgentRuntime();
      runtime.omitOkOnWrite = true;
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Update Data B2",
        mode: "preview_update",
        target: { sheetName: "Data", range: "B2" },
        values: valuePatch("Data", "B2", [[999]])
      });
      const applied = await agent.run({
        request: "Apply update",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(applied.status).toBe("SUCCESS");
      expect((applied.answer as any).ok).toBe(true);
      expect((applied.answer as any).validationOk).toBe(true);
    });

  it("uses parsed A1 overlap checks for formula-protected regions", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_formula_overlap");
      metadata.formulaRegions = [{ id: "formula:manual", sheetName: "Report", range: "A10:A20", formulaCount: 11 }];
      agent.metadataCache.set(metadata);

      const allowed = await agent.run({
        request: "Update report input",
        mode: "preview_update",
        workbookContextId: metadata.workbookContextId,
        target: { sheetName: "Report", range: "B1" },
        values: valuePatch("Report", "B1", [["ok"]])
      });
      const formulaPreview = await agent.run({
        request: "Update formula area",
        mode: "preview_update",
        workbookContextId: metadata.workbookContextId,
        target: { sheetName: "Report", range: "A12" },
        values: valuePatch("Report", "A12", [["bad"]])
      });

      expect(allowed.status).toBe("PREVIEW_READY");
      expect(formulaPreview.status).toBe("PREVIEW_READY");
      expect(formulaPreview.warnings.join(" ")).toContain("Patch 1 overlaps detected formula regions at Report!A12");
    });

  it("auto-applies scoped auto value edits by default", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Change Data B2 to 999",
        target: { sheetName: "Data", range: "B2" },
        values: valuePatch("Data", "B2", [[999]])
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.mode).toBe("auto");
      expect(result.confirmationToken).toBeUndefined();
      expect(result.nextAction).toBe("answer_now");
      expect(result.taskOutcome).toBe("apply_complete");
      expect(result.maxRecommendedFollowupCalls).toBe(0);
      expect(result.telemetry.autoApplied).toBe(true);
      expect(result.telemetry.safetyDecision).toBe("auto_apply:scoped_value_edit");
      expect(runtime.writeBatchCount).toBe(1);
    });

  it("auto-applies exact OpenCode price edits without a human confirmation turn", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_vendor_price_auto");
      metadata.sheets.push({
        id: "sheet:Vendor Propose",
        name: "Vendor Propose",
        index: 99,
        usedRange: "B2:K10",
        rowCount: 10,
        columnCount: 11,
        kind: "unknown",
        headers: [],
        tableIds: [],
        sectionIds: [],
        summaryBlockIds: [],
        formulaRegionIds: []
      });
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "I want to propose a price of 6100",
        mode: "auto",
        workbookContextId: metadata.workbookContextId,
        intent: { action: "write_values" },
        target: { sheetName: "Vendor Propose", range: "J4" },
        values: valuePatch("Vendor Propose", "J4", [[6100]])
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.taskOutcome).toBe("apply_complete");
      expect(result.nextAction).toBe("answer_now");
      expect(result.maxRecommendedFollowupCalls).toBe(0);
      expect(result.confirmationToken).toBeUndefined();
      expect(result.telemetry.autoApplied).toBe(true);
      expect(result.telemetry.safetyDecision).toBe("auto_apply:scoped_value_edit");
      expect(runtime.writeBatchCount).toBe(1);
      expect(runtime.lastBatchOperations[0]).toMatchObject({
        kind: "range.write_values_many",
        entries: [{ target: { sheetName: "Vendor Propose", address: "J4" }, values: [[6100]] }]
      });
    });

  it("allows callers to opt out of auto-apply for scoped auto value edits", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Change Data B2 to 999",
        target: { sheetName: "Data", range: "B2" },
        values: valuePatch("Data", "B2", [[999]]),
        autoApply: false
      });

      expect(result.status).toBe("PREVIEW_READY");
      expect(result.telemetry.autoApplied).toBeUndefined();
      expect(result.telemetry.safetyDecision).toBe("manual_review:auto_apply_disabled");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("auto-applies small add requests with explicit values and range", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Add a notes block to Report",
        target: { sheetName: "Report", range: "B1:B2" },
        values: valuePatch("Report", "B1:B2", [["Owner"], ["Finance"]])
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.telemetry.safetyDecision).toBe("auto_apply:scoped_value_edit");
      expect(runtime.writeBatchCount).toBe(1);
    });

  it("auto-applies fix/correct wording when explicit values and target are safe", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Fix Data C15 to 71-0409 and correct D15 to Maintenance fee",
        mode: "auto",
        intent: { action: "write_values" },
        target: { sheetName: "Data", range: "C15:D15" },
        values: valuePatch("Data", "C15:D15", [["71-0409", "Maintenance fee"]])
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.taskOutcome).toBe("apply_complete");
      expect(result.maxRecommendedFollowupCalls).toBe(0);
      expect(result.confirmationToken).toBeUndefined();
      expect(result.telemetry.autoApplied).toBe(true);
      expect(result.telemetry.safetyDecision).toBe("auto_apply:scoped_value_edit");
      expect(runtime.writeBatchCount).toBe(1);
    });

  it("auto-applies exact dropdown source-list value corrections as normal value writes", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Dropdown options 71-5226 and 71-5229 are wrong; update the source list values to 700-5226 and 700-5229.",
        mode: "auto",
        intent: { action: "write_values", targetHints: ["Truck ID dropdown source list"] },
        target: { sheetName: "Data", range: "A3:A4" },
        values: valuePatch("Data", "A3:A4", [["700-5226"], ["700-5229"]])
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.taskOutcome).toBe("apply_complete");
      expect(result.telemetry.autoApplied).toBe(true);
      expect(result.telemetry.safetyDecision).toBe("auto_apply:scoped_value_edit");
      expect(runtime.writeBatchCount).toBe(1);
      expect(runtime.lastBatchOperations[0]?.kind).toBe("range.write_values_many");
    });

  it("previews header formatting as a style mutation instead of reading values", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Format the header row on the Data sheet.",
        target: { sheetName: "Data", range: "A1:D1" }
      });

      expect(result.status).toBe("PREVIEW_READY");
      expect((result.answer as any).kind).toBe("style_preview");
      expect(result.proof[0]?.range).toBe("A1:D1");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("previews formula writes through a formula-aware path", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Duplicate the formula from A12 down",
        mode: "preview_update",
        target: { sheetName: "Report", range: "A13:A14" },
        values: { formulas: [["=SUM(B1:B10)"], ["=SUM(B2:B11)"]] }
      });

      expect(result.status).toBe("PREVIEW_READY");
      expect((result.answer as any).kind).toBe("formula_update_preview");
      expect(result.nextAction).toBe("call_apply_update");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("expands one A1 formula across the target range before apply", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Write Payment Variance formula to May 2026",
        mode: "preview_update",
        intent: { action: "write_formulas" },
        values: {
          patches: [{
            target: { sheetName: "May 2026", range: "I2:I4" },
            formulas: [["=H2-G2"]]
          }]
        }
      });
      const applied = await agent.run({
        request: "Apply formula repair",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe("formula_update_preview");
      expect((preview.answer as any).formulaPattern).toBe("=H2-G2");
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations).toHaveLength(1);
      expect(runtime.lastBatchOperations[0]).toMatchObject({
        kind: "range.write_formulas",
        target: { sheetName: "May 2026", address: "I2:I4" },
        formulas: [["=H2-G2"], ["=H3-G3"], ["=H4-G4"]]
      });
    });

  it("groups formula writes and red font styling in one preview", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Write formulas and make the text red",
        mode: "preview_update",
        intent: { action: "write_formulas" },
        values: {
          patches: [{
            target: { sheetName: "May 2026", range: "I2:I3" },
            formulas: [["=H2-G2"]],
            style: { fontColor: "#FF0000" }
          }]
        }
      });
      const applied = await agent.run({
        request: "Apply formula and style repair",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe("formula_update_preview");
      expect((preview.answer as any).operationCount).toBe(2);
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations.map((operation) => operation.kind)).toEqual(["range.write_formulas", "range.write_styles"]);
      expect(runtime.lastBatchOperations[1]).toMatchObject({
        kind: "range.write_styles",
        target: { sheetName: "May 2026", address: "I2:I3" },
        style: { fontColor: "#FF0000" },
        preserveValues: true
      });
    });

  it("accepts shorthand source and destination for formula fill with parent sheet", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Fill formula down",
        mode: "preview_update",
        intent: { action: "fill_formula_down" },
        target: { sheetName: "May 2026" },
        values: { source: "I4", destination: "I2:I244" }
      });
      const applied = await agent.run({
        request: "Apply formula fill",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe("formula_fill_down_preview");
      expect((preview.answer as any)).toMatchObject({ sheetName: "May 2026", range: "I2:I244" });
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.runtimeMethodCalls["formula.fill_down"]).toBe(1);
    });

  it("previews and applies number-format writes through the agent surface", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Format Data C2:C4 as currency",
      mode: "preview_update",
      intent: { action: "write_number_formats" },
      values: {
        patches: [{
          target: { sheetName: "Data", range: "C2:C4" },
          numberFormat: "$#,##0.00"
        }]
      }
    });
    const applied = await agent.run({
      request: "Apply number formats",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("number_format_preview");
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastWriteOperations[0]?.kind).toBe("range.write_number_formats");
    expect((runtime.lastWriteOperations[0] as any).numberFormat).toEqual([["$#,##0.00"], ["$#,##0.00"], ["$#,##0.00"]]);
  });

  it("previews clear-format and autofit-row range operations", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const clearFormats = await agent.run({
      request: "Clear formatting from Data A1:D1",
      mode: "preview_update",
      target: { sheetName: "Data", range: "A1:D1" }
    });
    const autofitRows = await agent.run({
      request: "Autofit row height for Data A1:D4",
      mode: "preview_update",
      target: { sheetName: "Data", range: "A1:D4" }
    });

    expect(clearFormats.status).toBe("PREVIEW_READY");
    expect((clearFormats.answer as any).kind).toBe("clear_formats_preview");
    expect(clearFormats.telemetry.operationRisk).toBe("safe_format");
    expect(autofitRows.status).toBe("PREVIEW_READY");
    expect((autofitRows.answer as any).dimension).toBe("rows");
      expect(autofitRows.telemetry.actionHandlerId).toBe("autofit_rows");
    });

  it("routes border-only removal to style-dimension clearing", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Remove all borders from selected cells",
      mode: "preview_update",
      target: { sheetName: "Data", range: "A3:B10" }
    });
    const applied = await agent.run({
      request: "Apply border cleanup",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("clear_style_dimensions_preview");
    expect((preview.answer as any).dimensions).toEqual(["borders"]);
    expect(preview.telemetry.safetyFingerprintOnly).toBe(true);
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastWriteOperations[0]?.kind).toBe("range.clear_style_dimensions");
    expect((runtime.lastWriteOperations[0] as any).dimensions).toEqual(["borders"]);
  });

  it("routes add-border requests to style writes with border payloads", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Add thin border to every cell in row 2",
      mode: "preview_update",
      target: { sheetName: "Data", range: "A2:D2" }
    });
    const applied = await agent.run({
      request: "Apply border style",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("style_preview");
    expect((preview.answer as any).style.borders).toEqual({ style: "continuous", weight: "thin" });
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastWriteOperations[0]?.kind).toBe("range.write_styles");
    expect((runtime.lastWriteOperations[0] as any).style.borders).toEqual({ style: "continuous", weight: "thin" });
  });

  it("previews and applies explicit range copy operations", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Copy range values",
      mode: "preview_update",
      intent: { action: "copy_range" },
      values: {
        source: { sheetName: "Data", range: "A1:B2" },
        destination: { sheetName: "Report", range: "A1:B2" },
        copyType: "values"
      }
    });
    const applied = await agent.run({
      request: "Apply copy",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("range_copy_preview");
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastWriteOperations[0]?.kind).toBe("range.copy");
    expect(runtime.lastBatchRequest?.operations[0]?.destructiveLevel).toBe("values");
    expect((runtime.lastWriteOperations[0] as any).source.address).toBe("A1:B2");
    expect(runtime.lastWriteOperations[0]?.target.address).toBe("A1:B2");
  });

  it("treats full range clear as structural risk", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Clear range including values and formats",
      mode: "preview_update",
      intent: { action: "clear_range" },
      target: { sheetName: "Data", range: "D4:E5" }
    });
    await agent.run({
      request: "Apply clear",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect(runtime.lastWriteOperations[0]?.kind).toBe("range.clear");
    expect(runtime.lastBatchRequest?.operations[0]?.destructiveLevel).toBe("structure");
  });

  it("uses conservative destructive levels for range copy-all and move operations", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const copyPreview = await agent.run({
      request: "Copy range including formulas and formats",
      mode: "preview_update",
      intent: { action: "copy_range" },
      values: {
        source: { sheetName: "Data", range: "A1:B2" },
        destination: { sheetName: "Report", range: "A1:B2" },
        copyType: "all"
      }
    });
    await agent.run({ request: "Apply copy", mode: "apply_update", operationId: copyPreview.operationId, confirmationToken: copyPreview.confirmationToken });
    expect(runtime.lastBatchRequest?.operations[0]?.destructiveLevel).toBe("structure");

    const movePreview = await agent.run({
      request: "Move range",
      mode: "preview_update",
      intent: { action: "move_range" },
      values: {
        source: { sheetName: "Data", range: "A1:B2" },
        destination: { sheetName: "Report", range: "C1:D2" }
      }
    });
    await agent.run({ request: "Apply move", mode: "apply_update", operationId: movePreview.operationId, confirmationToken: movePreview.confirmationToken });
    expect(runtime.lastBatchRequest?.operations[0]?.destructiveLevel).toBe("structure");
  });

  it("previews and applies safety artifact metadata mutations through the agent state machine", async () => {
    const cases: Array<{
      action: "refresh_snapshot" | "invalidate_snapshot" | "delete_snapshot" | "create_file_backup" | "restore_file_backup" | "prune_backups" | "pin_backup" | "unpin_backup" | "delete_backup";
      values: Record<string, unknown>;
      expectedKind: string;
      expectedRuntimeCall: string;
    }> = [
      { action: "refresh_snapshot", values: { snapshotId: "snapshot_agent_unit" }, expectedKind: "snapshot.refresh_preview", expectedRuntimeCall: "snapshot.refresh" },
      { action: "invalidate_snapshot", values: { snapshotId: "snapshot_agent_unit" }, expectedKind: "snapshot.invalidate_preview", expectedRuntimeCall: "snapshot.invalidate" },
      { action: "delete_snapshot", values: { snapshotId: "snapshot_agent_unit" }, expectedKind: "snapshot.delete_preview", expectedRuntimeCall: "snapshot.delete" },
      { action: "create_file_backup", values: { reason: "Before risky edit", pin: true }, expectedKind: "backup_create_file_preview", expectedRuntimeCall: "backup.create_file" },
      { action: "restore_file_backup", values: { backupId: "backup_file_unit", mode: "open-as-new" }, expectedKind: "backup_restore_file_preview", expectedRuntimeCall: "backup.restore_file" },
      { action: "prune_backups", values: { kind: "file-copy", maxBackupsPerWorkbook: 2 }, expectedKind: "backup_prune_preview", expectedRuntimeCall: "backup.prune" },
      { action: "pin_backup", values: { backupId: "backup_agent_unit" }, expectedKind: "backup.pin_preview", expectedRuntimeCall: "backup.pin" },
      { action: "unpin_backup", values: { backupId: "backup_agent_unit" }, expectedKind: "backup.unpin_preview", expectedRuntimeCall: "backup.unpin" },
      { action: "delete_backup", values: { backupId: "backup_agent_unit" }, expectedKind: "backup.delete_preview", expectedRuntimeCall: "backup.delete" }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: `Prepare ${testCase.action}`,
        mode: "preview_update",
        intent: { action: testCase.action },
        values: testCase.values
      });
      const applied = await agent.run({
        request: `Apply ${testCase.action}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe(testCase.expectedKind);
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.runtimeMethodCalls[testCase.expectedRuntimeCall]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    }
  });

  it("previews and applies workbook lifecycle mutations through backend orchestration", async () => {
    const localConfig = { version: 1, workbookId: "workbook_agent_unit", exportedAt: "2026-06-18T00:00:00.000Z", source: "open-workbook-local-config", templates: [], regions: [] };
    const cases: Array<{
      action: "restore_workbook_backup" | "import_local_config" | "embed_local_config" | "import_embedded_local_config" | "close_workbook";
      values: Record<string, unknown>;
      expectedKind: string;
      expectedRuntimeCall: string;
    }> = [
      { action: "restore_workbook_backup", values: { backupId: "backup_agent_unit" }, expectedKind: "workbook_restore_backup_preview", expectedRuntimeCall: "workbook.restore_backup" },
      { action: "import_local_config", values: { config: localConfig }, expectedKind: "workbook_import_local_config_preview", expectedRuntimeCall: "workbook.import_local_config" },
      { action: "embed_local_config", values: { includePermissions: true }, expectedKind: "workbook_embed_local_config_preview", expectedRuntimeCall: "workbook.embed_local_config" },
      { action: "import_embedded_local_config", values: { overwrite: true }, expectedKind: "workbook_import_embedded_local_config_preview", expectedRuntimeCall: "workbook.import_embedded_local_config" },
      { action: "close_workbook", values: { closeBehavior: "SkipSave" }, expectedKind: "workbook_close_preview", expectedRuntimeCall: "workbook.close" }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const preview = await agent.run({
        request: `Prepare ${testCase.action}`,
        mode: "preview_update",
        intent: { action: testCase.action },
        values: testCase.values
      });
      const applied = await agent.run({
        request: `Apply ${testCase.action}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe(testCase.expectedKind);
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.runtimeMethodCalls[testCase.expectedRuntimeCall]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    }
  });

  it("previews and applies formula operation intents through backend orchestration", async () => {
    const cases: Array<{
      action: "copy_formula_patterns" | "fill_formula_down" | "fill_formula_right" | "repair_formula_patterns" | "convert_formulas_to_values";
      expectedKind: string;
      expectedRuntimeCall: string;
      input: Parameters<AgentOrchestrator["run"]>[0];
    }> = [
      {
        action: "copy_formula_patterns",
        expectedKind: "formula_copy_patterns_preview",
        expectedRuntimeCall: "formula.copy_patterns",
        input: {
          request: "Copy formula patterns",
          mode: "preview_update",
          intent: { action: "copy_formula_patterns" },
          values: {
            source: { sheetName: "Report", range: "B2" },
            destination: { sheetName: "Report", range: "B3:B5" }
          }
        }
      },
      {
        action: "fill_formula_down",
        expectedKind: "formula_fill_down_preview",
        expectedRuntimeCall: "formula.fill_down",
        input: {
          request: "Fill formula down",
          mode: "preview_update",
          intent: { action: "fill_formula_down" },
          values: {
            source: { sheetName: "Report", range: "B2" },
            destination: { sheetName: "Report", range: "B3:B5" }
          }
        }
      },
      {
        action: "fill_formula_right",
        expectedKind: "formula_fill_right_preview",
        expectedRuntimeCall: "formula.fill_right",
        input: {
          request: "Fill formula right",
          mode: "preview_update",
          intent: { action: "fill_formula_right" },
          values: {
            source: { sheetName: "Report", range: "B2" },
            destination: { sheetName: "Report", range: "C2:E2" }
          }
        }
      },
      {
        action: "repair_formula_patterns",
        expectedKind: "formula_repair_patterns_preview",
        expectedRuntimeCall: "formula.repair_patterns",
        input: {
          request: "Repair formula patterns from template",
          mode: "preview_update",
          intent: { action: "repair_formula_patterns" },
          target: { sheetName: "Report" },
          values: { templateId: "template_unit" }
        }
      },
      {
        action: "convert_formulas_to_values",
        expectedKind: "formula_convert_to_values_preview",
        expectedRuntimeCall: "formula.convert_to_values",
        input: {
          request: "Convert formulas to values",
          mode: "preview_update",
          intent: { action: "convert_formulas_to_values" },
          target: { sheetName: "Report", range: "B2" }
        }
      }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run(testCase.input);
      const applied = await agent.run({
        request: `Apply ${testCase.action}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe(testCase.expectedKind);
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.runtimeMethodCalls[testCase.expectedRuntimeCall]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    }
  });

  it("previews and applies named item mutations through backend orchestration", async () => {
    const cases: Array<{
      action: "create_name" | "update_name" | "delete_name";
      expectedKind: string;
      expectedRuntimeCall: string;
      values: Record<string, unknown>;
      target?: Parameters<AgentOrchestrator["run"]>[0]["target"];
    }> = [
      { action: "create_name", expectedKind: "name_create_preview", expectedRuntimeCall: "names.create", values: { name: "NewInput", comment: "Unit input" }, target: { sheetName: "Report", range: "B2:B3" } },
      { action: "update_name", expectedKind: "name_update_preview", expectedRuntimeCall: "names.update", values: { name: "RevenueTotal", formula: "=Report!$B$2" } },
      { action: "delete_name", expectedKind: "name_delete_preview", expectedRuntimeCall: "names.delete", values: { name: "RevenueTotal" } }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: `Prepare ${testCase.action}`,
        mode: "preview_update",
        intent: { action: testCase.action },
        values: testCase.values,
        ...(testCase.target ? { target: testCase.target } : {})
      });
      const applied = await agent.run({
        request: `Apply ${testCase.action}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe(testCase.expectedKind);
      expect(applied.status).toBe("SUCCESS");
      const runtimeCallCount =
        runtime.runtimeMethodCalls[testCase.expectedRuntimeCall] ??
        runtime.tableMethodCalls.filter((call) => call.method === testCase.expectedRuntimeCall).length;
      expect(runtimeCallCount).toBe(1);
    }
  });

  it("previews and applies region mutations through backend orchestration", async () => {
    const cases: Array<{
      action: "register_region" | "clear_region_values" | "write_region_values" | "fill_region";
      expectedKind: string;
      expectedRuntimeCall: string;
      input: Parameters<AgentOrchestrator["run"]>[0];
    }> = [
      {
        action: "register_region",
        expectedKind: "region_register_preview",
        expectedRuntimeCall: "region.register",
        input: {
          request: "Register input region",
          mode: "preview_update",
          intent: { action: "register_region" },
          target: { sheetName: "Report", range: "B1:B3" },
          values: { regionName: "InputRegion", kind: "data", description: "Editable inputs" }
        }
      },
      {
        action: "clear_region_values",
        expectedKind: "region_clear_values_preview",
        expectedRuntimeCall: "region.clear_values",
        input: {
          request: "Clear region values",
          mode: "preview_update",
          intent: { action: "clear_region_values" },
          values: { regionName: "InputRegion" }
        }
      },
      {
        action: "write_region_values",
        expectedKind: "region_write_values_preview",
        expectedRuntimeCall: "region.write_values",
        input: {
          request: "Write region values",
          mode: "preview_update",
          intent: { action: "write_region_values" },
          values: { regionName: "InputRegion", values: [["Owner"], ["Finance"], ["Ready"]] }
        }
      },
      {
        action: "fill_region",
        expectedKind: "region_fill_preview",
        expectedRuntimeCall: "region.fill",
        input: {
          request: "Fill region values",
          mode: "preview_update",
          intent: { action: "fill_region" },
          values: { regionName: "InputRegion", rows: [["Owner"], ["Finance"], ["Ready"]], clearFirst: true }
        }
      }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run(testCase.input);
      const applied = await agent.run({
        request: `Apply ${testCase.action}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe(testCase.expectedKind);
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.runtimeMethodCalls[testCase.expectedRuntimeCall]).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    }
  });

  it("previews and applies template mutations through backend orchestration", async () => {
    const cases: Array<{
      action: "register_template" | "unregister_template" | "clear_template_data_regions" | "fill_template_regions" | "repair_sheet_from_template";
      expectedKind: string;
      expectedRuntimeCall: string;
      expectedWriteOperationKind?: string;
      input: Parameters<AgentOrchestrator["run"]>[0];
    }> = [
      {
        action: "register_template",
        expectedKind: "template_register_preview",
        expectedRuntimeCall: "template.register",
        input: {
          request: "Register template",
          mode: "preview_update",
          intent: { action: "register_template" },
          target: { sheetName: "Report" },
          values: { name: "Report Template", dataRegions: ["B2:B3"] }
        }
      },
      {
        action: "unregister_template",
        expectedKind: "template_unregister_preview",
        expectedRuntimeCall: "template.unregister",
        input: {
          request: "Unregister template",
          mode: "preview_update",
          intent: { action: "unregister_template" },
          values: { templateId: "template_unit" }
        }
      },
      {
        action: "clear_template_data_regions",
        expectedKind: "template_clear_data_regions_preview",
        expectedRuntimeCall: "template.repair_sheet",
        input: {
          request: "Clear template data regions",
          mode: "preview_update",
          intent: { action: "clear_template_data_regions" },
          target: { sheetName: "Report" },
          values: { templateId: "template_unit" }
        }
      },
      {
        action: "fill_template_regions",
        expectedKind: "template_fill_regions_preview",
        expectedRuntimeCall: "batch.apply",
        expectedWriteOperationKind: "range.write_values",
        input: {
          request: "Fill template data regions",
          mode: "preview_update",
          intent: { action: "fill_template_regions" },
          target: { sheetName: "Report" },
          values: { templateId: "template_unit", regionValues: { "B2:B3": [["North"], ["South"]] } }
        }
      },
      {
        action: "repair_sheet_from_template",
        expectedKind: "template_repair_preview",
        expectedRuntimeCall: "template.repair_sheet",
        input: {
          request: "Repair sheet from template",
          mode: "preview_update",
          intent: { action: "repair_sheet_from_template" },
          target: { sheetName: "Report" },
          values: { templateId: "template_unit", repair: ["styles", "formulas"] }
        }
      }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const preview = await agent.run(testCase.input);
      const applied = await agent.run({
        request: `Apply ${testCase.action}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe(testCase.expectedKind);
      expect(applied.status).toBe("SUCCESS");
      if (testCase.expectedWriteOperationKind) {
        expect(runtime.writeBatchCount).toBe(1);
        expect(runtime.lastWriteOperations[0]?.kind).toBe(testCase.expectedWriteOperationKind);
      } else {
        const runtimeCallCount =
          runtime.runtimeMethodCalls[testCase.expectedRuntimeCall] ??
          runtime.tableMethodCalls.filter((call) => call.method === testCase.expectedRuntimeCall).length;
        expect(runtimeCallCount).toBe(1);
      }
    }
  });

  it("previews and applies style template mutations through backend orchestration", async () => {
    const cases: Array<{
      action: "copy_style_from_template" | "repair_style_consistency";
      expectedKind: string;
      expectedRuntimeCall: string;
      input: Parameters<AgentOrchestrator["run"]>[0];
    }> = [
      {
        action: "copy_style_from_template",
        expectedKind: "style_copy_preview",
        expectedRuntimeCall: "style.copy_dimensions",
        input: {
          request: "Copy style from template",
          mode: "preview_update",
          intent: { action: "copy_style_from_template" },
          values: {
            source: { sheetName: "Data", range: "A1:B4" },
            destination: { sheetName: "Report", range: "A1:B4" },
            dimensions: ["fills", "fonts"]
          }
        }
      },
      {
        action: "repair_style_consistency",
        expectedKind: "style_repair_consistency_preview",
        expectedRuntimeCall: "style.repair_consistency",
        input: {
          request: "Repair style consistency",
          mode: "preview_update",
          intent: { action: "repair_style_consistency" },
          target: { sheetName: "Report" },
          values: { templateId: "template_unit" }
        }
      }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const preview = await agent.run(testCase.input);
      const applied = await agent.run({
        request: `Apply ${testCase.action}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe(testCase.expectedKind);
      expect(applied.status).toBe("SUCCESS");
      const runtimeCallCount =
        runtime.runtimeMethodCalls[testCase.expectedRuntimeCall] ??
        runtime.tableMethodCalls.filter((call) => call.method === testCase.expectedRuntimeCall).length;
      expect(runtimeCallCount).toBe(1);
    }
  });

  it("previews grouped style copies as one apply operation", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Copy same style from source to targets",
      mode: "preview_update",
      intent: { action: "copy_style_from_template" },
      values: {
        styleCopies: [
          { source: { sheetName: "Data", range: "A1:B1" }, destination: { sheetName: "Report", range: "A1:B1" }, dimensions: ["fills", "fonts"] },
          { source: { sheetName: "Data", range: "A2:B2" }, destination: { sheetName: "Report", range: "C1:D1" }, dimensions: ["fills", "fonts"] }
        ]
      }
    });
    const applied = await agent.run({
      request: "Apply grouped style copies",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("style_copy_many_preview");
    expect(applied.status).toBe("SUCCESS");
    expect((applied.answer as any).backupIds).toEqual(["backup_style_many"]);
    expect((applied.answer as any).rollbackAvailable).toBe(true);
    expect(runtime.runtimeMethodCalls["style.copy_dimensions_many"]).toBe(1);
    expect(runtime.runtimeMethodCalls["style.copy_dimensions"]).toBeUndefined();
  });

  it("previews and applies a styled table replacement as one workflow", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const headers = ["Dated", "Loading Date", "Qty", "Attention", "From (FM)", "Description", "Received (Tons)", "Volume"];
    const row = ["20/6/26", "20/6/26", "5X40'HQ", "5X40'HQ", "SC88", "Loading at Sai Noi Factory", 32.5, "5X40'HQ UPGRADE"];
    const preview = await agent.run({
      request: "Rotate booking fields to headers and replace the old vertical table with same style",
      mode: "preview_update",
      intent: { action: "replace_range_with_styled_table" },
      target: { sheetName: "Report", range: "A1:H2" },
      values: {
        headers,
        row,
        clearRange: "A1:B25",
        headerStyleSource: { sheetName: "Data", range: "A1:D1" },
        bodyStyleSource: { sheetName: "Data", range: "A2:D2" },
        dimensions: ["fills", "fonts", "borders", "alignment"],
        autofit: true
      }
    });
    const applied = await agent.run({
      request: "Apply styled table replacement",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("replace_range_with_styled_table_preview");
    expect((preview.answer as any).clearCount).toBe(1);
    expect((preview.answer as any).styleCopyCount).toBe(4);
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.writeBatchCount).toBe(1);
    expect(runtime.lastWriteOperations.map((operation) => operation.kind)).toEqual(["range.clear", "range.write_values", "range.autofit_columns"]);
    expect((runtime.lastWriteOperations[0] as any).applyTo).toBe("all");
    expect(runtime.runtimeMethodCalls["style.copy_dimensions_many"]).toBe(1);
    expect(runtime.runtimeMethodCalls["style.copy_dimensions"]).toBeUndefined();
  });

  it("preserves column widths by default when replacing a styled table", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Rotate booking fields to headers and replace the old vertical table with same style",
      mode: "preview_update",
      intent: { action: "replace_range_with_styled_table" },
      target: { sheetName: "Report", range: "A1:C2" },
      values: {
        headers: ["Dated", "Loading Date", "Qty"],
        row: ["20/6/26", "20/6/26", "5X40'HQ"],
        clearRange: "A1:B25",
        headerStyleSource: { sheetName: "Data", range: "A1:C1" },
        bodyStyleSource: { sheetName: "Data", range: "A2:C2" },
        dimensions: ["fills", "fonts", "borders", "alignment"]
      }
    });
    const applied = await agent.run({
      request: "Apply styled table replacement",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastWriteOperations.map((operation) => operation.kind)).toEqual(["range.clear", "range.write_values"]);
  });

  it("routes generic OCR field/value data to the styled table replacement workflow", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Convert extracted invoice field/value OCR data into a horizontal styled table",
      mode: "preview_update",
      target: { sheetName: "Report", range: "A1:D2" },
      values: {
        headers: ["Invoice", "Date", "Customer", "Total"],
        row: ["INV-100", "2026-06-20", "ACME", 2500],
        clearRange: "A1:D20",
        headerStyleSource: { sheetName: "Data", range: "A1:D1" },
        bodyStyleSource: { sheetName: "Data", range: "A2:D2" }
      }
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("replace_range_with_styled_table_preview");
    expect((preview.answer as any).styleCopyCount).toBe(2);
  });

  it("redirects fragmented style-copy previews to a grouped workflow", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);
    const metadata = createCachedMetadata("wbctx_fragmented_style");
    agent.metadataCache.set(metadata);

    const base = {
      mode: "preview_update" as const,
      workbookContextId: metadata.workbookContextId,
      values: {
        source: { sheetName: "Data", range: "A2:D2" }
      }
    };
    await agent.run({ ...base, request: "Copy body style to Booking A2:D2", values: { ...base.values, destination: { sheetName: "Report", range: "A2:D2" } } });
    const redirected = await agent.run({ ...base, request: "Copy body style to Booking E2:H2", values: { ...base.values, destination: { sheetName: "Report", range: "E2:H2" } } });

    expect(redirected.status).toBe("NEEDS_WORKFLOW_REDIRECT");
    expect(redirected.nextAction).toBe("call_preview_update");
    expect((redirected.answer as any).kind).toBe("workflow_redirect");
    expect((redirected.answer as any).suggestedIntentAction).toBe("copy_style_from_template");
    expect((redirected.answer as any).suggestedValues.styleCopies).toHaveLength(2);
    expect(redirected.warnings).toContain("Fragmented style-copy previews were redirected to a grouped workflow.");
  });

  it("redirects repeated value-only cleanup because stale formats can remain", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);
    const metadata = createCachedMetadata("wbctx_fragmented_clear");
    agent.metadataCache.set(metadata);

    await agent.run({
      request: "Clear leftover data in A3:X10",
      mode: "preview_update",
      workbookContextId: metadata.workbookContextId,
      intent: { action: "clear_values" },
      target: { sheetName: "Report", range: "A3:X10" }
    });
    const redirected = await agent.run({
      request: "Clear leftover data in A11:X18",
      mode: "preview_update",
      workbookContextId: metadata.workbookContextId,
      intent: { action: "clear_values" },
      target: { sheetName: "Report", range: "A11:X18" }
    });

    expect(redirected.status).toBe("NEEDS_WORKFLOW_REDIRECT");
    expect((redirected.answer as any).suggestedIntentAction).toBe("clear_range");
    expect((redirected.answer as any).suggestedValues.clearRange).toBe("A3:X18");
    expect(redirected.warnings).toContain("Fragmented value-only cleanup was redirected because it can leave stale formatting.");
  });

  it("redirects adjacent value-write previews to grouped patches", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);
    const metadata = createCachedMetadata("wbctx_fragmented_write_values");
    agent.metadataCache.set(metadata);

    await agent.run({
      request: "Write first extracted value",
      mode: "preview_update",
      workbookContextId: metadata.workbookContextId,
      target: { sheetName: "Report", range: "A1" },
      values: valuePatch("Report", "A1", [["A"]])
    });
    const redirected = await agent.run({
      request: "Write second extracted value",
      mode: "preview_update",
      workbookContextId: metadata.workbookContextId,
      target: { sheetName: "Report", range: "B1" },
      values: valuePatch("Report", "B1", [["B"]])
    });

    expect(redirected.status).toBe("PREVIEW_READY");
    expect((redirected.answer as any).kind).toBe("multi_range_preview");
    expect((redirected.answer as any).patchCount).toBe(1);
  });

  it("redirects repeated format previews to write_styles_many", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);
    const metadata = createCachedMetadata("wbctx_fragmented_format");
    agent.metadataCache.set(metadata);

    await agent.run({
      request: "Make A1 bold",
      mode: "preview_update",
      workbookContextId: metadata.workbookContextId,
      intent: { action: "format_range" },
      target: { sheetName: "Report", range: "A1" }
    });
    const redirected = await agent.run({
      request: "Make B1 bold",
      mode: "preview_update",
      workbookContextId: metadata.workbookContextId,
      intent: { action: "format_range" },
      target: { sheetName: "Report", range: "B1" }
    });

    expect(redirected.status).toBe("NEEDS_WORKFLOW_REDIRECT");
    expect((redirected.answer as any).suggestedIntentAction).toBe("write_styles_many");
    expect((redirected.answer as any).suggestedValues.entries).toHaveLength(2);
    expect(redirected.warnings).toContain("Fragmented format previews were redirected to a grouped style workflow.");
  });

  describe("OpenCode production regressions", () => {
    it("applies explicit black header styling without falling back to default header colors", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Set header row A1:X1 on Data sheet to black fill (#000000) with white font color (#FFFFFF), bold text, center aligned.",
        mode: "preview_update",
        intent: { action: "format_range" },
        target: { sheetName: "Data", range: "A1:X1" }
      });
      const applied = await agent.run({
        request: "Apply black header style.",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe("style_preview");
      expect((preview.answer as any).style).toMatchObject({
        fillColor: "#000000",
        fontColor: "#FFFFFF",
        fontBold: true,
        horizontalAlignment: "center"
      });
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations[0]).toMatchObject({
        kind: "range.write_styles",
        target: { sheetName: "Data", address: "A1:X1" },
        style: {
          fillColor: "#000000",
          fontColor: "#FFFFFF",
          fontBold: true,
          horizontalAlignment: "center"
        },
        preserveValues: true
      });
    });

    it("applies flattened multi-style entries as one write_styles_many batch", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Set yellow fill #FFFF00 on rows 2, 4, and 7 of Booking sheet for full row A:X.",
        mode: "preview_update",
        intent: { action: "write_styles_many" },
        target: { sheetName: "Booking" },
        values: {
          entries: [
            { range: "A2:X2", sheetName: "Booking", fillColor: "#FFFF00" },
            { range: "A4:X4", sheetName: "Booking", fillColor: "#FFFF00" },
            { range: "A7:X7", sheetName: "Booking", fillColor: "#FFFF00" }
          ]
        }
      });
      const applied = await agent.run({
        request: "Apply yellow row highlights.",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.writeBatchCount).toBe(1);
      expect(runtime.lastBatchOperations).toEqual([
        expect.objectContaining({
          kind: "range.write_styles_many",
          entries: [
            { target: expect.objectContaining({ sheetName: "Booking", address: "A2:X2" }), style: { fillColor: "#FFFF00" }, preserveValues: true },
            { target: expect.objectContaining({ sheetName: "Booking", address: "A4:X4" }), style: { fillColor: "#FFFF00" }, preserveValues: true },
            { target: expect.objectContaining({ sheetName: "Booking", address: "A7:X7" }), style: { fillColor: "#FFFF00" }, preserveValues: true }
          ]
        })
      ]);
    });

    it("prepends real merge operations when batched style entries ask to merge header spans", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Merge row 1 grouped header spans and center align them on Invoices.",
        mode: "preview_update",
        intent: { action: "write_styles_many" },
        target: { sheetName: "Invoices" },
        values: {
          entries: [
            { sheetName: "Invoices", range: "A1:B1", style: { horizontalAlignment: "center", verticalAlignment: "center" } },
            { sheetName: "Invoices", range: "C1:F1", style: { horizontalAlignment: "center", verticalAlignment: "center" } },
            { sheetName: "Invoices", range: "G1:N1", style: { horizontalAlignment: "center", verticalAlignment: "center" } },
            { sheetName: "Invoices", range: "O1:O1", style: { horizontalAlignment: "center", verticalAlignment: "center" } }
          ]
        }
      });
      const applied = await agent.run({
        request: "Apply grouped header merge alignment.",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any)).toMatchObject({ kind: "merge_and_write_styles_many_preview", mergeCount: 3, rangeCount: 4 });
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations.map((operation) => operation.kind)).toEqual([
        "range.merge",
        "range.merge",
        "range.merge",
        "range.write_styles_many"
      ]);
      expect(runtime.lastBatchOperations.slice(0, 3)).toEqual([
        expect.objectContaining({ kind: "range.merge", target: expect.objectContaining({ sheetName: "Invoices", address: "A1:B1" }) }),
        expect.objectContaining({ kind: "range.merge", target: expect.objectContaining({ sheetName: "Invoices", address: "C1:F1" }) }),
        expect.objectContaining({ kind: "range.merge", target: expect.objectContaining({ sheetName: "Invoices", address: "G1:N1" }) })
      ]);
      expect(runtime.lastBatchOperations.at(-1)).toMatchObject({
        kind: "range.write_styles_many",
        entries: [
          expect.objectContaining({ target: expect.objectContaining({ address: "A1:B1" }), style: expect.objectContaining({ horizontalAlignment: "center", verticalAlignment: "center" }) }),
          expect.objectContaining({ target: expect.objectContaining({ address: "C1:F1" }), style: expect.objectContaining({ horizontalAlignment: "center", verticalAlignment: "center" }) }),
          expect.objectContaining({ target: expect.objectContaining({ address: "G1:N1" }), style: expect.objectContaining({ horizontalAlignment: "center", verticalAlignment: "center" }) }),
          expect.objectContaining({ target: expect.objectContaining({ address: "O1:O1" }), style: expect.objectContaining({ horizontalAlignment: "center", verticalAlignment: "center" }) })
        ]
      });
    });

    it("supports explicit multi-range merge payloads with default center alignment", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Merge these row 1 header spans and center them vertically and horizontally.",
        mode: "preview_update",
        target: { sheetName: "Invoices" },
        values: {
          merges: [
            { sheetName: "Invoices", range: "A1:B1" },
            { sheetName: "Invoices", range: "C1:F1" },
            { sheetName: "Invoices", range: "G1:N1" }
          ]
        }
      });
      const applied = await agent.run({
        request: "Apply grouped header merges.",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any)).toMatchObject({ kind: "merge_ranges_preview", mergeCount: 3, styledRangeCount: 3 });
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations.map((operation) => operation.kind)).toEqual([
        "range.merge",
        "range.merge",
        "range.merge",
        "range.write_styles_many"
      ]);
      expect(runtime.lastBatchOperations.at(-1)).toMatchObject({
        kind: "range.write_styles_many",
        entries: [
          expect.objectContaining({ style: expect.objectContaining({ horizontalAlignment: "center", verticalAlignment: "center", wrapText: true }) }),
          expect.objectContaining({ style: expect.objectContaining({ horizontalAlignment: "center", verticalAlignment: "center", wrapText: true }) }),
          expect.objectContaining({ style: expect.objectContaining({ horizontalAlignment: "center", verticalAlignment: "center", wrapText: true }) })
        ]
      });
    });

  it("applies inserted columns through the structural batch operation", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Add new columns next to Qty on Data.",
        mode: "preview_update",
        intent: { action: "insert_columns" },
        target: { sheetName: "Data", range: "C:D" }
      });
      const applied = await agent.run({
        request: "Apply inserted columns.",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe("range.insert_columns_preview");
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations[0]).toMatchObject({
        kind: "range.insert_columns",
        destructiveLevel: "structure",
        target: { sheetName: "Data", address: "C1:D4" }
      });
    });

    it("applies range column swaps with a normalized columnOrder payload", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Swap columns A and B in Data",
        mode: "preview_update",
        intent: { action: "reorder_range_columns" },
        target: { sheetName: "Data", range: "A1:B4" },
        values: { order: [2, 1] }
      });
      const applied = await agent.run({
        request: "Apply column swap.",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).columnOrder).toEqual([2, 1]);
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations[0]).toMatchObject({
        kind: "range.reorder_columns",
        destructiveLevel: "structure",
        target: { sheetName: "Data", address: "A1:B4" },
        columnOrder: [2, 1]
      });
    });

    it("applies table column swaps through the table host method", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Move Status before Date in the Transactions table.",
        mode: "preview_update",
        intent: { action: "reorder_table_columns" },
        target: { tableName: "Transactions" },
        values: { columns: ["Status", "Date", "Account", "Amount"] }
      });
      const applied = await agent.run({
        request: "Apply table column order.",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe("table_reorder_columns_preview");
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.tableMethodCalls).toContainEqual({
        method: "table.reorder_columns",
        request: expect.objectContaining({
          tableName: "Transactions",
          columnOrder: ["Status", "Date", "Account", "Amount"]
        })
      });
    });

    it("applies dropdown validation from agent-friendly options", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Add select list to Data D2:D4.",
        mode: "preview_update",
        intent: { action: "write_data_validation" },
        values: {
          patches: [{
            target: { sheetName: "Data", range: "D2:D4" },
            options: ["20GP", "40GP", "40HQ"],
            inCellDropDown: true
          }]
        }
      });
      const applied = await agent.run({
        request: "Apply dropdown validation.",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect(((preview.answer as any).entries?.[0]?.validation ?? (preview.answer as any).validation)).toEqual({
        type: "list",
        source: ["20GP", "40GP", "40HQ"],
        inCellDropDown: true,
        ignoreBlanks: true
      });
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations[0]).toMatchObject({
        kind: "range.write_data_validation",
        target: { sheetName: "Data", address: "D2:D4" },
        validation: {
          type: "list",
          source: ["20GP", "40GP", "40HQ"],
          inCellDropDown: true,
          ignoreBlanks: true
        }
      });
    });

    it("preserves and normalizes formula sources for dropdown validation", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Update Transaction Type dropdown source range.",
        mode: "preview_update",
        intent: { action: "write_data_validation" },
        values: {
          patches: [{
            target: { sheetName: "May 2026", range: "E2:E244" },
            validation: { type: "list", source: "=Dropdown Lists!$B$2:$B$28", inCellDropDown: true }
          }]
        }
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect(((preview.answer as any).entries?.[0]?.validation ?? (preview.answer as any).validation)).toMatchObject({
        type: "list",
        source: "='Dropdown Lists'!$B$2:$B$28",
        inCellDropDown: true,
        ignoreBlanks: true
      });
      const operation = (agent.dumpOperations()[0] as any).action.operations[0];
      expect(operation).toMatchObject({
        kind: "range.write_data_validation",
        validation: { source: "='Dropdown Lists'!$B$2:$B$28" }
      });
    });

    it("applies formula conditional formatting from formula and style fields", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Add formula color on Data A2:D4 when column D is 40HQ.",
        mode: "preview_update",
        intent: { action: "write_conditional_formatting" },
        values: {
          patches: [{
            target: { sheetName: "Data", range: "A2:D4" },
            conditionalFormatting: {
              formula: "=$D2=\"40HQ\"",
              style: { fillColor: "#FFFF00", fontColor: "#000000" }
            }
          }]
        }
      });
      const applied = await agent.run({
        request: "Apply formula conditional formatting.",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe("write_conditional_formatting_preview");
      expect((preview.answer as any).kind).not.toBe("sheet_create_preview");
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations[0]).toMatchObject({
        kind: "range.write_conditional_formatting",
        target: { sheetName: "Data", address: "A2:D4" },
        rule: {
          type: "custom",
          formula: "=$D2=\"40HQ\"",
          style: { fillColor: "#FFFF00", fontColor: "#000000" }
        }
      });
    });
  });

  it("honors explicit black header colors instead of default header styling", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Set header row A1:X1 on Data sheet to black fill (#000000) with white font color (#FFFFFF), bold text, center aligned.",
      mode: "preview_update",
      intent: { action: "format_range" },
      target: { sheetName: "Data", range: "A1:X1" }
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).style).toMatchObject({
      fillColor: "#000000",
      fontColor: "#FFFFFF",
      fontBold: true
    });
  });

  it("honors explicit header fill hex when the request also names a color family", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Set HR A1:E1 to dark blue fill (#1F4E78), white font (#FFFFFF), bold text, center aligned.",
      mode: "preview_update",
      intent: { action: "format_range" },
      target: { sheetName: "Data", range: "A1:E1" }
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).style).toMatchObject({
      fillColor: "#1F4E78",
      fontColor: "#FFFFFF",
      fontBold: true,
      horizontalAlignment: "center"
    });
  });

  it("accepts flattened write_styles_many entries from agents", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Set yellow fill #FFFF00 on rows 2, 4, and 7 of Booking sheet for full row A:X.",
      mode: "preview_update",
      intent: { action: "write_styles_many" },
      target: { sheetName: "Booking" },
      values: {
        entries: [
          { range: "A2:X2", sheetName: "Booking", fillColor: "#FFFF00" },
          { range: "A4:X4", sheetName: "Booking", fillColor: "#FFFF00" },
          { range: "A7:X7", sheetName: "Booking", fillColor: "#FFFF00" }
        ]
      }
    });
    const applied = await agent.run({
      request: "Apply yellow highlights.",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastBatchOperations[0]).toMatchObject({
      kind: "range.write_styles_many",
      entries: [
        { target: { sheetName: "Booking", address: "A2:X2" }, style: { fillColor: "#FFFF00" } },
        { target: { sheetName: "Booking", address: "A4:X4" }, style: { fillColor: "#FFFF00" } },
        { target: { sheetName: "Booking", address: "A7:X7" }, style: { fillColor: "#FFFF00" } }
      ]
    });
  });

  it("accepts top-level style patches from agents and expands whole-column width targets", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Set column widths on Data sheet.",
      mode: "preview_update",
      intent: { action: "format_range" },
      target: { sheetName: "Data" },
      patches: [
        { target: { sheetName: "Data", range: "A:A" }, style: { columnWidth: 14 } },
        { target: { sheetName: "Data", range: "B:B" }, style: { columnWidth: 12 } }
      ]
    } as any);
    const applied = await agent.run({
      request: "Apply column widths.",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("write_styles_many_preview");
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastBatchOperations[0]).toMatchObject({
      kind: "range.write_styles_many",
      entries: [
        { target: { sheetName: "Data", address: "A1:A4" }, style: { columnWidth: 73.5 } },
        { target: { sheetName: "Data", address: "B1:B4" }, style: { columnWidth: 63 } }
      ]
    });
  });

  it("routes style-shaped values.patches to write_styles_many instead of writing objects into cells", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Set column widths on Data sheet for all columns in one batch",
      mode: "preview_update",
      intent: { action: "format_range" },
      values: {
        patches: [
          { target: { sheetName: "Data", range: "A1" }, values: [[{ columnWidth: 14 }]] },
          { target: { sheetName: "Data", range: "B1" }, values: [[{ columnWidth: 12 }]] }
        ]
      }
    });
    const applied = await agent.run({
      request: "Apply column widths.",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("write_styles_many_preview");
    expect((preview.answer as any).rangeCount).toBe(2);
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastBatchOperations[0]).toMatchObject({
      kind: "range.write_styles_many",
      entries: [
        { target: { sheetName: "Data", address: "A1" }, style: { columnWidth: 73.5 } },
        { target: { sheetName: "Data", address: "B1" }, style: { columnWidth: 63 } }
      ]
    });
    expect(runtime.lastBatchOperations[0]?.kind).not.toBe("range.write_values_many");
  });

  it("rejects empty style previews instead of applying a no-op", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Set column widths on Data sheet but forgot structured style values.",
      mode: "preview_update",
      intent: { action: "format_range" },
      target: { sheetName: "Data" }
    });

    expect(preview.status).toBe("NEEDS_INPUT");
    expect(preview.summary).toContain("Style update needs at least one supported style property");
    expect(runtime.lastBatchOperations.some((operation) => operation.kind === "range.write_styles")).toBe(false);
    expect(runtime.lastBatchOperations.some((operation) => operation.kind === "range.write_styles_many")).toBe(false);
  });

  it("previews dropdown data validation as a validation operation", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Set data validation on Data D2:D4 as a list dropdown. Allowed values: 20GP, 40GP, 40HQ.",
      mode: "preview_update",
      intent: { action: "write_data_validation" },
      values: {
        patches: [{
          target: { sheetName: "Data", range: "D2:D4" },
          validation: { type: "List", formula1: "20GP,40GP,40HQ" }
        }]
      }
    });
    const applied = await agent.run({
      request: "Apply dropdown validation.",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("write_data_validation_preview");
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastBatchOperations[0]).toMatchObject({
      kind: "range.write_data_validation",
      target: { sheetName: "Data", address: "D2:D4" },
      validation: { type: "list", source: ["20GP", "40GP", "40HQ"] }
    });
  });

  it("normalizes unquoted sheet range formula1 for dropdown data validation", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Update May 2026 transaction type dropdown source.",
      mode: "preview_update",
      intent: { action: "write_data_validation" },
      values: {
        patches: [{
          target: { sheetName: "May 2026", range: "E2:E244" },
          validation: { type: "list", formula1: "Dropdown Lists!$B$2:$B$28", inCellDropDown: true }
        }]
      }
    });
    const applied = await agent.run({
      request: "Apply dropdown validation.",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect(((preview.answer as any).entries?.[0]?.validation ?? (preview.answer as any).validation).source).toBe("='Dropdown Lists'!$B$2:$B$28");
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastBatchOperations[0]).toMatchObject({
      kind: "range.write_data_validation",
      target: { sheetName: "May 2026", address: "E2:E244" },
      validation: { type: "list", source: "='Dropdown Lists'!$B$2:$B$28" }
    });
  });

  it("previews multi-range dropdown validation entries without broadening to the table", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Add dropdown data validation to column A and column E only.",
      mode: "preview_update",
      target: { sheetName: "Data", tableName: "Transactions" },
      values: {
        validation: { type: "list", source: ["Open", "Closed"], inCellDropDown: true },
        entries: [
          { sheetName: "Data", range: "A3:A1002", validation: { type: "list", source: ["Open", "Closed"], inCellDropDown: true } },
          { sheetName: "Data", range: "E3:E1002", validation: { type: "list", source: ["Open", "Closed"], inCellDropDown: true } }
        ]
      }
    });
    const applied = await agent.run({
      request: "Apply dropdown validation.",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("write_data_validation_preview");
    expect((preview.answer as any).entries.map((entry: any) => entry.range)).toEqual(["A3:A1002", "E3:E1002"]);
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastBatchOperations[0]).toMatchObject({
      kind: "range.write_data_validation",
      entries: [
        { target: { sheetName: "Data", address: "A3:A1002" }, validation: { type: "list", source: ["Open", "Closed"] } },
        { target: { sheetName: "Data", address: "E3:E1002" }, validation: { type: "list", source: ["Open", "Closed"] } }
      ]
    });
    expect((runtime.lastBatchOperations[0] as any).target.address).not.toBe("A1:Z1000");
  });

  it("previews formula conditional formatting instead of sheet creation", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Add conditional formatting rule on Data sheet range A2:D4. Formula =$D2=\"40HQ\" should fill the entire row with yellow color #FFFF00.",
      mode: "preview_update",
      target: { sheetName: "Data", range: "A2:D4" },
      values: { formula: "=$D2=\"40HQ\"", style: { fillColor: "#FFFF00" } }
    });
    const applied = await agent.run({
      request: "Apply conditional formatting.",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("write_conditional_formatting_preview");
    expect((preview.answer as any).kind).not.toBe("sheet_create_preview");
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastBatchOperations[0]).toMatchObject({
      kind: "range.write_conditional_formatting",
      target: { sheetName: "Data", address: "A2:D4" },
      rule: { type: "custom", formula: "=$D2=\"40HQ\"", style: { fillColor: "#FFFF00" } }
    });
  });

  it("previews plain range column reorder for swap requests", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Swap columns A and B in Data",
      mode: "preview_update",
      intent: { action: "reorder_range_columns" },
      target: { sheetName: "Data", range: "A1:B4" },
      values: { columnOrder: [2, 1] }
    });
    const applied = await agent.run({
      request: "Apply column swap.",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("range.reorder_columns_preview");
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastBatchOperations[0]).toMatchObject({
      kind: "range.reorder_columns",
      target: { sheetName: "Data", address: "A1:B4" },
      columnOrder: [2, 1]
    });
  });

  it("redirects repeated formula fills to one larger formula workflow", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);
    const metadata = createCachedMetadata("wbctx_fragmented_formula_fill");
    agent.metadataCache.set(metadata);

    const base = {
      request: "Fill formula down",
      mode: "preview_update" as const,
      workbookContextId: metadata.workbookContextId,
      intent: { action: "fill_formula_down" as const }
    };
    await agent.run({
      ...base,
      values: {
        source: { sheetName: "Report", range: "A1" },
        destination: { sheetName: "Report", range: "A2:A5" }
      }
    });
    const redirected = await agent.run({
      ...base,
      values: {
        source: { sheetName: "Report", range: "A1" },
        destination: { sheetName: "Report", range: "A6:A10" }
      }
    });

    expect(redirected.status).toBe("NEEDS_WORKFLOW_REDIRECT");
    expect((redirected.answer as any).suggestedIntentAction).toBe("fill_formula_down");
    expect((redirected.answer as any).suggestedValues.destination).toEqual({ sheetName: "Report", range: "A2:A10" });
    expect(redirected.warnings).toContain("Fragmented formula-fill-down previews were redirected to a grouped formula workflow.");
  });

  it("reports partial workflow failure with rollback evidence", async () => {
    const runtime = new FakeAgentRuntime();
    runtime.failStyleCopyOnCall = 2;
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Replace table and copy style",
      mode: "preview_update",
      intent: { action: "replace_range_with_styled_table" },
      target: { sheetName: "Report", range: "A1:B2" },
      values: {
        headers: ["Dated", "Qty"],
        row: ["20/6/26", "5X40'HQ"],
        clearRange: "A1:B20",
        headerStyleSource: { sheetName: "Data", range: "A1:B1" },
        bodyStyleSource: { sheetName: "Data", range: "A2:B2" },
        dimensions: ["fills", "fonts"]
      }
    });
    const applied = await agent.run({
      request: "Apply styled table replacement",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect(applied.status).toBe("VALIDATION_FAILED");
    expect((applied.answer as any).partialFailure).toBe(true);
    expect((applied.answer as any).rollbackAvailable).toBe(true);
    expect((applied.answer as any).backupIds).toContain("backup_style_many");
    expect((applied.answer as any).stepResults).toHaveLength(2);
    expect(applied.warnings).toContain("Style copy failed");
  });

  it("rejects mismatched direct style copy ranges before apply", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Copy same style from Data to Report",
      mode: "preview_update",
      intent: { action: "copy_style_from_template" },
      values: {
        source: { sheetName: "Data", range: "A1:D1" },
        destination: { sheetName: "Report", range: "A1:H1" },
        dimensions: ["fills", "fonts"]
      }
    });

    expect(preview.status).toBe("NEEDS_INPUT");
    expect(preview.summary).toContain("same dimensions");
    expect(runtime.runtimeMethodCalls["style.copy_dimensions"]).toBeUndefined();
  });

  it("returns cached apply results when the same operation is retried", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Write values",
      mode: "preview_update",
      intent: { action: "write_values" },
      target: { sheetName: "Report", range: "A1:B1" },
      values: valuePatch("Report", "A1:B1", [["A", "B"]])
    });
    const first = await agent.run({
      request: "Apply values",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });
    const second = await agent.run({
      request: "Apply values again",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(first.status).toBe("SUCCESS");
    expect(second.status).toBe("SUCCESS");
    expect(runtime.writeBatchCount).toBe(1);
  });

  it("previews and applies repair capabilities through backend orchestration", async () => {
    const cases: Array<{
      action: "repair_style_from_template" | "repair_formulas_from_template" | "repair_table_structure";
      expectedKind: string;
      expectedRuntimeCall: string;
      input: Parameters<AgentOrchestrator["run"]>[0];
    }> = [
      {
        action: "repair_style_from_template",
        expectedKind: "style_repair_consistency_preview",
        expectedRuntimeCall: "style.repair_consistency",
        input: {
          request: "Repair styles from template",
          mode: "preview_update",
          intent: { action: "repair_style_from_template" },
          target: { sheetName: "Report" },
          values: { templateId: "template_unit" }
        }
      },
      {
        action: "repair_formulas_from_template",
        expectedKind: "formula_repair_patterns_preview",
        expectedRuntimeCall: "formula.repair_patterns",
        input: {
          request: "Repair formulas from template",
          mode: "preview_update",
          intent: { action: "repair_formulas_from_template" },
          target: { sheetName: "Report" },
          values: { templateId: "template_unit" }
        }
      },
      {
        action: "repair_table_structure",
        expectedKind: "table_copy_structure_preview",
        expectedRuntimeCall: "table.copy_structure",
        input: {
          request: "Repair table structure",
          mode: "preview_update",
          intent: { action: "repair_table_structure" },
          target: { tableName: "Transactions" },
          values: { targetSheetName: "Report", targetAddress: "A1:D4" }
        }
      }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const preview = await agent.run(testCase.input);
      const applied = await agent.run({
        request: `Apply ${testCase.action}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe(testCase.expectedKind);
      expect(applied.status).toBe("SUCCESS");
      const runtimeCallCount =
        runtime.runtimeMethodCalls[testCase.expectedRuntimeCall] ??
        runtime.tableMethodCalls.filter((call) => call.method === testCase.expectedRuntimeCall).length;
      expect(runtimeCallCount).toBe(1);
    }
  });

  it("previews and applies cleaning mutations through backend orchestration", async () => {
    const cases: Array<{
      action: "normalize_headers" | "trim_whitespace" | "remove_duplicates" | "parse_dates" | "parse_numbers" | "standardize_currency" | "fill_missing_values" | "split_column" | "merge_columns";
      expectedRuntimeCall: string;
      values?: Record<string, unknown>;
    }> = [
      { action: "normalize_headers", expectedRuntimeCall: "clean.normalize_headers", values: { headerRowIndex: 0 } },
      { action: "trim_whitespace", expectedRuntimeCall: "clean.trim_whitespace" },
      { action: "remove_duplicates", expectedRuntimeCall: "clean.remove_duplicates", values: { hasHeader: true, keyColumns: [0, 1] } },
      { action: "parse_dates", expectedRuntimeCall: "clean.parse_dates" },
      { action: "parse_numbers", expectedRuntimeCall: "clean.parse_numbers" },
      { action: "standardize_currency", expectedRuntimeCall: "clean.standardize_currency" },
      { action: "fill_missing_values", expectedRuntimeCall: "clean.fill_missing_values", values: { strategy: "zero" } },
      { action: "split_column", expectedRuntimeCall: "clean.split_column", values: { columnIndex: 1, delimiter: "-", targetAddress: "E1:F4" } },
      { action: "merge_columns", expectedRuntimeCall: "clean.merge_columns", values: { columnIndexes: [1, 2], separator: " ", targetAddress: "E1:E4" } }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const preview = await agent.run({
        request: `Clean data with ${testCase.action}`,
        mode: "preview_update",
        intent: { action: testCase.action },
        target: { sheetName: "Data", range: "A1:D4" },
        values: testCase.values
      });
      const applied = await agent.run({
        request: `Apply ${testCase.action}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe("cleaning_preview");
      expect((preview.answer as any).action).toBe(testCase.action);
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.runtimeMethodCalls[testCase.expectedRuntimeCall]).toBe(1);
    }
  });

  it("previews and applies backend-compiled column value transforms without model-sized matrices", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Add acct- prefix to all Account values",
      mode: "preview_update",
      intent: { action: "transform_values" },
      target: { sheetName: "Data", tableName: "Transactions", column: "Account" },
      values: { operation: "add_prefix", prefix: "acct-" }
    });
    const applied = await agent.run({
      request: "Apply account prefix transform",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("transform_values_preview");
    expect((preview.answer as any).changedCount).toBe(3);
    expect((preview.answer as any).target.range).toBe("B2:B4");
    expect((preview.answer as any).examples).toHaveLength(3);
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastBatchOperations[0]?.kind).toBe("range.write_values_many");
    expect((runtime.lastBatchOperations[0] as any).entries).toEqual([{
      target: { workbookId, sheetName: "Data", address: "B2:B4" },
      values: [["acct-A-100"], ["acct-A-200"], ["acct-A-300"]],
      preserveFormats: true
    }]);
  });

  it("keeps backend-compiled broad transforms preview-first in auto mode", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Add acct- prefix to all Account values",
      mode: "auto",
      intent: { action: "transform_values" },
      target: { sheetName: "Data", tableName: "Transactions", column: "Account" },
      values: { operation: "add_prefix", prefix: "acct-" }
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect(preview.taskOutcome).toBe("preview_ready");
    expect(preview.telemetry.autoApplied).not.toBe(true);
    expect(preview.telemetry.safetyDecision).toBe("manual_review:broad_compiled_transform");
    expect(runtime.writeBatchCount).toBe(0);
  });

  it("previews and applies row-aware copy-if-blank derivations from source columns", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Fill blank Account from Amount for this row",
      mode: "preview_update",
      intent: { action: "derive_values" },
      target: { sheetName: "Data", range: "B3" },
      values: { operation: "copy_if_blank", sourceRange: "C3" }
    });
    const applied = await agent.run({
      request: "Apply blank account derivation",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("derive_values_preview");
    expect((preview.answer as any).changedCount).toBe(1);
    expect((preview.answer as any).rowAlignment).toEqual({ type: "same_row", rows: 1 });
    expect((preview.answer as any).examples[0]).toMatchObject({ row: 3, source: { C: 200 } });
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastBatchOperations[0]?.kind).toBe("range.write_values_many");
    expect((runtime.lastBatchOperations[0] as any).entries).toEqual([{
      target: { workbookId, sheetName: "Data", address: "B3:B3" },
      values: [[200]],
      preserveFormats: true
    }]);
  });

  it("previews and applies lookup-map derivations from another sheet", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Update Status from Account using Customer Master lookup",
      mode: "preview_update",
      intent: { action: "derive_values" },
      target: { sheetName: "Data", tableName: "Transactions", column: "Status" },
      values: {
        operation: "lookup_map",
        sourceColumn: "Account",
        lookupSheetName: "Customer Master",
        lookupKeyColumn: "Account",
        lookupValueColumn: "Tier"
      }
    });
    const applied = await agent.run({
      request: "Apply lookup derivation",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("derive_values_preview");
    expect((preview.answer as any).rowAlignment).toEqual({ type: "lookup_map", rows: 3, lookupRows: 3 });
    expect((preview.answer as any).changedCount).toBe(3);
    expect((preview.answer as any).examples[0]).toMatchObject({
      row: 2,
      source: { Account: "A-100" },
      after: { Status: "Gold" }
    });
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastBatchOperations[0]?.kind).toBe("range.write_values_many");
    expect((runtime.lastBatchOperations[0] as any).entries).toEqual([{
      target: { workbookId, sheetName: "Data", address: "D2:D4" },
      values: [["Gold"], ["Silver"], ["Bronze"]],
      preserveFormats: true
    }]);
  });

  it("previews row-aware formula-like payment variance derivations from cash and actual columns", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Preview Payment Variance as Actual Amount minus Cash Amount for May 2026",
      mode: "preview_update",
      intent: { action: "derive_values" },
      target: { sheetName: "May 2026", tableName: "May2026_Transaction_Filter", column: "Payment Variance" },
      values: {
        operation: "formula_like",
        formula: "actual_minus_cash",
        sourceColumns: ["Cash Amount", "Actual Amount"]
      }
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("derive_values_preview");
    expect((preview.answer as any).target.range).toBe("I2:I244");
    expect((preview.answer as any).target.header).toBe("Payment Variance");
    expect((preview.answer as any).sources.map((source: any) => source.header)).toEqual(["Cash Amount", "Actual Amount"]);
    expect((preview.answer as any).rowAlignment).toEqual({ type: "same_row", rows: 243 });
    expect((preview.answer as any).examples[0]).toMatchObject({
      row: 2,
      source: { "Cash Amount": "2211.21", "Actual Amount": "2211.21" },
      after: { "Payment Variance": 0 }
    });
    expect(runtime.writeBatchCount).toBe(0);
  });

  it("previews Payment Variance formula derivation without skipping blank source rows", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Payment Variance should be formula Actual Amount minus Cash Amount for May 2026 and text red",
      mode: "preview_update",
      intent: { action: "derive_values" },
      target: { sheetName: "May 2026", tableName: "May2026_Transaction_Filter", column: "Payment Variance" },
      values: {
        operation: "formula_like",
        outputMode: "formula",
        style: { fontColor: "#FF0000" }
      }
    });
    const applied = await agent.run({
      request: "Apply Payment Variance formula repair",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("formula_update_preview");
    expect((preview.answer as any).target.range).toBe("I2:I244");
    expect((preview.answer as any).sources.map((source: any) => source.header)).toEqual(["Cash Amount", "Actual Amount"]);
    expect((preview.answer as any).changedCells).toBe(243);
    expect((preview.answer as any).examples[0]).toMatchObject({
      row: 2,
      after: { "Payment Variance": "=H2-G2" }
    });
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastBatchOperations.map((operation) => operation.kind)).toEqual(["range.write_formulas", "range.write_styles"]);
    expect(runtime.lastBatchOperations[0]).toMatchObject({
      kind: "range.write_formulas",
      target: { sheetName: "May 2026", address: "I2:I244" }
    });
    expect((runtime.lastBatchOperations[0] as any).formulas[242]).toEqual(["=H244-G244"]);
  });

  it("previews settlement bundles with payment variance formulas and separate reconciliation/detail notes", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Make May 2026 settlement consistent with Apr 2026 for row 2",
      mode: "preview_update",
      intent: { action: "settle_reconciliation" },
      target: { sheetName: "May 2026" },
      values: {
        referenceSheetName: "Apr 2026",
        rowUpdates: [{
          row: 2,
          reconciliationNote: "Settled against owner top-up",
          detailNotes: "Owner fund top-up matched to bank transfer"
        }]
      }
    });
    const applied = await agent.run({
      request: "Apply settlement bundle",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect(preview.taskOutcome).toBe("preview_ready");
    expect((preview.answer as any).kind).toBe("settlement_bundle_preview");
    expect((preview.answer as any).target.header).toBe("Payment Variance");
    expect((preview.answer as any).sources.map((source: any) => source.header)).toEqual(["Cash Amount", "Actual Amount"]);
    expect((preview.answer as any).noteTargets.map((target: any) => target.header)).toEqual(["Reconciliation Note", "Detail Notes"]);
    expect((preview.answer as any).reference.paymentVariance).toMatchObject({
      range: "I2:I244",
      formulaExample: "=H2-G2",
      formulaCell: "I2"
    });
    expect((preview.answer as any).examples[0]).toMatchObject({
      row: 2,
      source: { "Cash Amount": "2211.21", "Actual Amount": "2211.21" },
      after: {
        "Payment Variance": "=H2-G2",
        "Reconciliation Note": "Settled against owner top-up",
        "Detail Notes": "Owner fund top-up matched to bank transfer"
      }
    });
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastBatchOperations.map((operation) => operation.kind)).toEqual(["range.write_formulas", "range.write_values_many"]);
    expect((runtime.lastBatchOperations[0] as any).target).toEqual({ workbookId, sheetName: "May 2026", address: "I2:I2" });
    expect((runtime.lastBatchOperations[0] as any).formulas).toEqual([["=H2-G2"]]);
    expect((runtime.lastBatchOperations[1] as any).entries).toEqual([
      {
        target: { workbookId, sheetName: "May 2026", address: "J2:J2" },
        values: [["Settled against owner top-up"]],
        preserveFormats: true
      },
      {
        target: { workbookId, sheetName: "May 2026", address: "M2:M2" },
        values: [["Owner fund top-up matched to bank transfer"]],
        preserveFormats: true
      }
    ]);
  });

  it("previews workbook structure sheet prefix transforms as one batch plan", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Add FY26 - prefix to Data and Report sheets",
      mode: "preview_update",
      intent: { action: "transform_sheets" },
      values: {
        operation: "add_prefix",
        prefix: "FY26 - ",
        sheets: ["Data", "Report"]
      }
    });
    const applied = await agent.run({
      request: "Apply sheet prefix plan",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("transform_sheets_preview");
    expect((preview.answer as any).changedCount).toBe(2);
    expect((preview.answer as any).examples).toEqual([
      { from: "Data", to: "FY26 - Data" },
      { from: "Report", to: "FY26 - Report" }
    ]);
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastBatchOperations.map((operation) => operation.kind)).toEqual(["sheet.rename", "sheet.rename"]);
    expect(runtime.lastBatchOperations.map((operation: any) => [operation.sheetName, operation.newSheetName])).toEqual([
      ["Data", "FY26 - Data"],
      ["Report", "FY26 - Report"]
    ]);
  });

  it("previews formula recalculation as a workbook calculate operation", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Recalculate formulas",
      mode: "preview_update",
      intent: { action: "recalculate_formulas" }
    });
    const applied = await agent.run({
      request: "Apply formula recalc",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect((preview.answer as any).kind).toBe("formula.recalculate_preview");
    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastBatchOperations[0]?.kind).toBe("workbook.calculate");
  });

  it("routes every covered range-core capability to its internal operation", async () => {
    const cases: Array<{
      capabilityName: string;
      expectedOperationKind: string;
      input: Parameters<AgentOrchestrator["run"]>[0];
    }> = [
      {
        capabilityName: "excel.range.write_values",
        expectedOperationKind: "range.write_values_many",
        input: {
          request: "Update Data B2",
          mode: "preview_update",
          target: { sheetName: "Data", range: "B2" },
          values: valuePatch("Data", "B2", [[123]])
        }
      },
      {
        capabilityName: "excel.range.write_formulas",
        expectedOperationKind: "range.write_formulas",
        input: {
          request: "Write formula to Report A12",
          mode: "preview_update",
          intent: { action: "write_formulas" },
          values: {
            patches: [{
              target: { sheetName: "Report", range: "A12" },
              formulas: [["=SUM(B1:B10)"]]
            }]
          }
        }
      },
      {
        capabilityName: "excel.range.write_number_formats",
        expectedOperationKind: "range.write_number_formats",
        input: {
          request: "Format Data C2 as currency",
          mode: "preview_update",
          intent: { action: "write_number_formats" },
          values: {
            patches: [{
              target: { sheetName: "Data", range: "C2" },
              numberFormat: "$#,##0.00"
            }]
          }
        }
      },
      {
        capabilityName: "excel.range.write_styles",
        expectedOperationKind: "range.write_styles",
        input: {
          request: "Format the header row on Data",
          mode: "preview_update",
          intent: { action: "format_range" },
          target: { sheetName: "Data", range: "A1:D1" }
        }
      },
      {
        capabilityName: "excel.range.clear",
        expectedOperationKind: "range.clear",
        input: {
          request: "Clear the whole target",
          mode: "preview_update",
          intent: { action: "clear_range" },
          target: { sheetName: "Data", range: "D4" }
        }
      },
      {
        capabilityName: "excel.range.clear_values_keep_format",
        expectedOperationKind: "range.clear_values_keep_format",
        input: {
          request: "Clear values",
          mode: "preview_update",
          intent: { action: "clear_values" },
          target: { sheetName: "Data", range: "B3" }
        }
      },
      {
        capabilityName: "excel.range.clear_values",
        expectedOperationKind: "range.clear_values",
        input: {
          request: "Clear raw values",
          mode: "preview_update",
          intent: { action: "clear_values_raw" },
          target: { sheetName: "Data", range: "B3" }
        }
      },
      {
        capabilityName: "excel.range.clear_formats",
        expectedOperationKind: "range.clear_formats",
        input: {
          request: "Clear formatting",
          mode: "preview_update",
          intent: { action: "clear_formats" },
          target: { sheetName: "Data", range: "A1:D1" }
        }
      },
      {
        capabilityName: "excel.range.copy",
        expectedOperationKind: "range.copy",
        input: {
          request: "Copy range",
          mode: "preview_update",
          intent: { action: "copy_range" },
          values: {
            source: { sheetName: "Data", range: "A1:B2" },
            destination: { sheetName: "Report", range: "A1:B2" }
          }
        }
      },
      {
        capabilityName: "excel.range.write_styles_many",
        expectedOperationKind: "range.write_styles_many",
        input: {
          request: "Style multiple ranges",
          mode: "preview_update",
          intent: { action: "write_styles_many" },
          values: {
            entries: [
              { sheetName: "Data", range: "A1:D1", style: { bold: true } },
              { sheetName: "Report", range: "B2:C2", style: { fillColor: "#fff2cc" } }
            ]
          }
        }
      },
      {
        capabilityName: "excel.range.move",
        expectedOperationKind: "range.move",
        input: {
          request: "Move range",
          mode: "preview_update",
          intent: { action: "move_range" },
          values: {
            source: { sheetName: "Data", range: "A1:B2" },
            destination: { sheetName: "Report", range: "C1:D2" }
          }
        }
      },
      {
        capabilityName: "excel.range.autofit_columns",
        expectedOperationKind: "range.autofit_columns",
        input: {
          request: "Autofit columns",
          mode: "preview_update",
          intent: { action: "autofit" },
          target: { sheetName: "Data", range: "A:D" }
        }
      },
      {
        capabilityName: "excel.range.autofit_rows",
        expectedOperationKind: "range.autofit_rows",
        input: {
          request: "Autofit rows",
          mode: "preview_update",
          intent: { action: "autofit_rows" },
          target: { sheetName: "Data", range: "A1:D4" }
        }
      },
      {
        capabilityName: "excel.range.insert_rows",
        expectedOperationKind: "range.insert_rows",
        input: {
          request: "Insert rows",
          mode: "preview_update",
          intent: { action: "insert_rows" },
          target: { sheetName: "Data", range: "2:3" }
        }
      },
      {
        capabilityName: "excel.range.delete_rows",
        expectedOperationKind: "range.delete_rows",
        input: {
          request: "Delete rows",
          mode: "preview_update",
          intent: { action: "delete_rows" },
          target: { sheetName: "Data", range: "2:3" }
        }
      },
      {
        capabilityName: "excel.range.delete_rows.corrected_from_bad_column_intent",
        expectedOperationKind: "range.delete_rows",
        input: {
          request: "Please delete row 2",
          mode: "preview_update",
          intent: { action: "delete_columns" },
          target: { sheetName: "Data", range: "2:2" }
        }
      },
      {
        capabilityName: "excel.range.insert_columns",
        expectedOperationKind: "range.insert_columns",
        input: {
          request: "Insert columns",
          mode: "preview_update",
          intent: { action: "insert_columns" },
          target: { sheetName: "Data", range: "B:C" }
        }
      },
      {
        capabilityName: "excel.range.delete_columns",
        expectedOperationKind: "range.delete_columns",
        input: {
          request: "Delete columns",
          mode: "preview_update",
          intent: { action: "delete_columns" },
          target: { sheetName: "Data", range: "B:C" }
        }
      },
      {
        capabilityName: "excel.range.hide_columns",
        expectedOperationKind: "range.hide_columns",
        input: {
          request: "Hide columns",
          mode: "preview_update",
          intent: { action: "hide_columns" },
          target: { sheetName: "Data", range: "B:C" }
        }
      },
      {
        capabilityName: "excel.range.unhide_columns",
        expectedOperationKind: "range.unhide_columns",
        input: {
          request: "Unhide columns",
          mode: "preview_update",
          intent: { action: "unhide_columns" },
          target: { sheetName: "Data", range: "B:C" }
        }
      },
      {
        capabilityName: "excel.range.merge",
        expectedOperationKind: "range.merge",
        input: {
          request: "Merge range",
          mode: "preview_update",
          intent: { action: "merge_range" },
          target: { sheetName: "Report", range: "A1:B1" },
          values: { across: true }
        }
      },
      {
        capabilityName: "excel.range.unmerge",
        expectedOperationKind: "range.unmerge",
        input: {
          request: "Unmerge range",
          mode: "preview_update",
          intent: { action: "unmerge_range" },
          target: { sheetName: "Report", range: "A1:B1" }
        }
      }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const preview = await agent.run(testCase.input);

      expect(preview.status, testCase.capabilityName).toBe("PREVIEW_READY");

      const applied = await agent.run({
        request: `Apply ${testCase.capabilityName}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(applied.status, testCase.capabilityName).toBe("SUCCESS");
      expect(runtime.lastBatchOperations[0]?.kind, testCase.capabilityName).toBe(testCase.expectedOperationKind);
    }
  });

  it("routes every covered table-core mutation to its runtime method", async () => {
    const cases: Array<{
      capabilityName: string;
      expectedMethod: string;
      input: Parameters<AgentOrchestrator["run"]>[0];
    }> = [
      {
        capabilityName: "excel.table.append_rows",
        expectedMethod: "table.append_rows",
        input: {
          request: "Append transaction rows to the Transactions table",
          mode: "preview_update",
          intent: { action: "append_table_rows" },
          target: { tableName: "Transactions" },
          values: { rows: [["2026-06-01", "A-100", 123, "Open"]] }
        }
      },
      {
        capabilityName: "excel.table.update_rows",
        expectedMethod: "table.update_rows",
        input: {
          request: "Update transaction table row",
          mode: "preview_update",
          intent: { action: "update_table_rows" },
          target: { tableName: "Transactions" },
          values: { rows: [{ index: 0, values: ["2026-06-01", "A-100", 999, "Closed"] }] }
        }
      },
      {
        capabilityName: "excel.table.create",
        expectedMethod: "table.create",
        input: {
          request: "Create a table on Report B1:D4",
          mode: "preview_update",
          intent: { action: "create_table" },
          target: { sheetName: "Report", range: "B1:D4" },
          values: { tableName: "ReportInputs", hasHeaders: true, style: "TableStyleMedium2" }
        }
      },
      {
        capabilityName: "excel.table.resize",
        expectedMethod: "table.resize",
        input: {
          request: "Resize Transactions table",
          mode: "preview_update",
          intent: { action: "resize_table" },
          target: { tableName: "Transactions" },
          values: { address: "A1:D10" }
        }
      },
      {
        capabilityName: "excel.table.reorder_columns",
        expectedMethod: "table.reorder_columns",
        input: {
          request: "Reorder Transactions columns",
          mode: "preview_update",
          intent: { action: "reorder_table_columns" },
          target: { tableName: "Transactions" },
          values: { columnOrder: ["Status", "Date", "Account", "Amount"] }
        }
      },
      {
        capabilityName: "excel.table.clear_data_keep_formulas",
        expectedMethod: "table.clear_data_keep_formulas",
        input: {
          request: "Clear table data rows",
          mode: "preview_update",
          intent: { action: "clear_table_data" },
          target: { tableName: "Transactions" }
        }
      },
      {
        capabilityName: "excel.table.clear_filters",
        expectedMethod: "table.clear_filters",
        input: {
          request: "Clear table filters",
          mode: "preview_update",
          intent: { action: "clear_table_filters" },
          target: { tableName: "Transactions" }
        }
      },
      {
        capabilityName: "excel.table.apply_filters",
        expectedMethod: "table.apply_filters",
        input: {
          request: "Filter Transactions table",
          mode: "preview_update",
          intent: { action: "filter_range" },
          target: { tableName: "Transactions" },
          values: { filters: [{ column: "Status", criteria: { filterOn: "Values", values: ["Open"] } }] }
        }
      },
      {
        capabilityName: "excel.table.sort",
        expectedMethod: "table.sort",
        input: {
          request: "Sort Transactions by amount",
          mode: "preview_update",
          intent: { action: "sort_table" },
          target: { tableName: "Transactions" }
        }
      },
      {
        capabilityName: "excel.table.apply_view",
        expectedMethod: "table.apply_view",
        input: {
          request: "Filter Transactions to open and sort by amount descending",
          mode: "preview_update",
          intent: { action: "apply_table_view" },
          target: { tableName: "Transactions" },
          values: {
            filters: [{ column: "Status", criteria: { filterOn: "Values", values: ["Open"] } }],
            sort: { fields: [{ key: 2, ascending: false }] }
          }
        }
      },
      {
        capabilityName: "excel.table.set_total_row",
        expectedMethod: "table.set_total_row",
        input: {
          request: "Show total row on Transactions table",
          mode: "preview_update",
          intent: { action: "set_table_total_row" },
          target: { tableName: "Transactions" },
          values: { showTotals: true }
        }
      },
      {
        capabilityName: "excel.table.set_style",
        expectedMethod: "table.set_style",
        input: {
          request: "Set table style",
          mode: "preview_update",
          intent: { action: "set_table_style" },
          target: { tableName: "Transactions" },
          values: { style: "TableStyleMedium2" }
        }
      },
      {
        capabilityName: "excel.table.copy_structure",
        expectedMethod: "table.copy_structure",
        input: {
          request: "Copy Transactions table structure",
          mode: "preview_update",
          intent: { action: "copy_table_structure" },
          target: { tableName: "Transactions" },
          values: { targetSheetName: "Report", targetAddress: "A10:D12", newTableName: "TransactionsCopy" }
        }
      }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const preview = await agent.run(testCase.input);

      expect(preview.status, testCase.capabilityName).toBe("PREVIEW_READY");

      const applied = await agent.run({
        request: `Apply ${testCase.capabilityName}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(applied.status, testCase.capabilityName).toBe("SUCCESS");
      expect(runtime.tableMethodCalls.at(-1)?.method, testCase.capabilityName).toBe(testCase.expectedMethod);
    }
  });

  it("rejects incomplete table append rows before creating a preview", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Append a row to Transactions with one missing value.",
      mode: "preview_update",
      target: { tableName: "Transactions" },
      values: { rows: [["2026-01-15", "Globex", "Support"]] }
    });

    expect(preview.status).toBe("VALIDATION_FAILED");
    expect(preview.summary).toContain("has 3 value(s)");
    expect(preview.summary).toContain("4 column(s)");
    expect(preview.nextAction).toBe("ask_user");
    expect(runtime.tableMethodCalls.length).toBe(0);
  });

  it("routes every covered sheet-core mutation to its internal operation", async () => {
    const cases: Array<{
      capabilityName: string;
      expectedOperationKind: string;
      input: Parameters<AgentOrchestrator["run"]>[0];
    }> = [
      {
        capabilityName: "excel.sheet.create",
        expectedOperationKind: "sheet.create",
        input: {
          request: "Create a new sheet",
          mode: "preview_update",
          intent: { action: "create_sheet" },
          values: { sheetName: "New Report" }
        }
      },
      {
        capabilityName: "excel.sheet.copy",
        expectedOperationKind: "sheet.copy",
        input: {
          request: "Copy the Report sheet",
          mode: "preview_update",
          intent: { action: "copy_sheet" },
          target: { sheetName: "Report" },
          values: { newSheetName: "Report Copy" }
        }
      },
      {
        capabilityName: "excel.sheet.rename",
        expectedOperationKind: "sheet.rename",
        input: {
          request: "Rename the Report sheet",
          mode: "preview_update",
          intent: { action: "rename_sheet" },
          target: { sheetName: "Report" },
          values: { newSheetName: "Report Renamed" }
        }
      },
      {
        capabilityName: "excel.sheet.delete",
        expectedOperationKind: "sheet.delete",
        input: {
          request: "Delete the Report sheet",
          mode: "preview_update",
          intent: { action: "delete_sheet" },
          target: { sheetName: "Report" }
        }
      },
      {
        capabilityName: "excel.sheet.hide",
        expectedOperationKind: "sheet.hide",
        input: {
          request: "Hide the Report sheet",
          mode: "preview_update",
          intent: { action: "hide_sheet" },
          target: { sheetName: "Report" }
        }
      },
      {
        capabilityName: "excel.sheet.unhide",
        expectedOperationKind: "sheet.unhide",
        input: {
          request: "Unhide the Report sheet",
          mode: "preview_update",
          intent: { action: "unhide_sheet" },
          target: { sheetName: "Report" }
        }
      },
      {
        capabilityName: "excel.sheet.protect",
        expectedOperationKind: "sheet.protect",
        input: {
          request: "Protect the Report sheet",
          mode: "preview_update",
          intent: { action: "protect_sheet" },
          target: { sheetName: "Report" },
          values: { password: "unit" }
        }
      },
      {
        capabilityName: "excel.sheet.unprotect",
        expectedOperationKind: "sheet.unprotect",
        input: {
          request: "Unprotect the Report sheet",
          mode: "preview_update",
          intent: { action: "unprotect_sheet" },
          target: { sheetName: "Report" },
          values: { password: "unit" }
        }
      },
      {
        capabilityName: "excel.sheet.clear",
        expectedOperationKind: "sheet.clear",
        input: {
          request: "Clear the Report sheet",
          mode: "preview_update",
          intent: { action: "clear_sheet" },
          target: { sheetName: "Report" },
          values: { applyTo: "contents" }
        }
      },
      {
        capabilityName: "excel.sheet.set_tab_color",
        expectedOperationKind: "sheet.set_tab_color",
        input: {
          request: "Set the Report sheet tab color",
          mode: "preview_update",
          intent: { action: "set_sheet_tab_color" },
          target: { sheetName: "Report" },
          values: { color: "#4472C4" }
        }
      }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const preview = await agent.run(testCase.input);

      expect(preview.status, testCase.capabilityName).toBe("PREVIEW_READY");

      const applied = await agent.run({
        request: `Apply ${testCase.capabilityName}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(applied.status, testCase.capabilityName).toBe("SUCCESS");
      expect(runtime.lastBatchOperations[0]?.kind, testCase.capabilityName).toBe(testCase.expectedOperationKind);
    }
  });

  it("resolves active/current sheet phrases for destructive sheet operations", async () => {
    const runtime = new FakeAgentRuntime();
    const agent = new AgentOrchestrator(runtime as any);

    const preview = await agent.run({
      request: "Delete the current sheet",
      mode: "preview_update",
      intent: { action: "delete_sheet" }
    });

    expect(preview.status).toBe("PREVIEW_READY");
    expect(preview.summary).toBe("Prepared sheet deletion for Data.");

    const applied = await agent.run({
      request: "Apply delete current sheet",
      mode: "apply_update",
      operationId: preview.operationId,
      confirmationToken: preview.confirmationToken
    });

    expect(applied.status).toBe("SUCCESS");
    expect(runtime.lastBatchOperations[0]).toMatchObject({ kind: "sheet.delete", sheetName: "Data" });
  });

  it("routes covered workbook snapshot and backup actions to runtime methods", async () => {
    const cases: Array<{
      capabilityName: string;
      intentAction: "create_snapshot" | "create_backup";
      expectedMethod: string;
    }> = [
      { capabilityName: "excel.workbook.snapshot", intentAction: "create_snapshot", expectedMethod: "workbook.snapshot" },
      { capabilityName: "excel.workbook.create_backup", intentAction: "create_backup", expectedMethod: "workbook.create_backup" }
    ];

    for (const testCase of cases) {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const preview = await agent.run({
        request: `Create ${testCase.capabilityName}`,
        mode: "preview_update",
        intent: { action: testCase.intentAction },
        target: { sheetName: "Data", range: "A1:D4" },
        values: { reason: `Unit ${testCase.capabilityName}` }
      });
      const applied = await agent.run({
        request: `Apply ${testCase.capabilityName}`,
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status, testCase.capabilityName).toBe("PREVIEW_READY");
      expect(applied.status, testCase.capabilityName).toBe("SUCCESS");
      expect(runtime.runtimeMethodCalls[testCase.expectedMethod], testCase.capabilityName).toBe(1);
      expect(runtime.writeBatchCount, testCase.capabilityName).toBe(0);
    }
  });

  it("previews duplicate sheet template cleanup as one pending operation", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Can you duplicate Report sheet, remove data and keep only template?",
        target: { sheetName: "Report" },
        values: { dataRegions: ["B2:B3"] }
      });
      const applied = await agent.run({
        request: "Apply template cleanup",
        mode: "apply_update",
        operationId: result.operationId,
        confirmationToken: result.confirmationToken
      });

      expect(result.status).toBe("PREVIEW_READY");
      expect((result.answer as any).kind).toBe("template_cleanup_preview");
      expect((result.answer as any).operationKind).toBe("sheet.copy_clean_data_regions");
      expect((result.answer as any).dataRegions).toEqual(["B2:B3"]);
      expect((result.answer as any).sourceSheetName).toBe("Report");
      expect(result.nextAction).toBe("call_apply_update");
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.writeBatchCount).toBe(1);
      expect(runtime.lastBatchOperations).toHaveLength(1);
      expect(runtime.lastBatchOperations[0]?.kind).toBe("sheet.copy_clean_data_regions");
    });

  it("previews and applies table appends through the agent surface", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Append transaction rows to the Transactions table",
        mode: "preview_update",
        target: { tableName: "Transactions" },
        values: { rows: [["2026-06-01", 204, "71-4653", "Company gas top-up"]] }
      });
      const applied = await agent.run({
        request: "Apply append",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe("table_append_preview");
      expect((preview.answer as any).tableName).toBe("Transactions");
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.appendTableRowCount).toBe(1);
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("does not auto-apply ambiguous natural-language updates without canonical patch values", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Update financial 2026"
      });

      expect(result.status).toBe("NEEDS_INPUT");
      expect(result.summary).toContain("Preview needs structured values");
      expect(result.nextAction).toBe("ask_user");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("does not auto-apply formula-sensitive requests as value writes", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Fix formula in Report A12",
        target: { sheetName: "Report", range: "A12" },
        values: valuePatch("Report", "A12", [[100]])
      });

      expect(result.status).toBe("NEEDS_INPUT");
      expect(result.nextAction).toBe("manual_review");
      expect(result.telemetry.safetyDecision).toBe("manual_review:advanced_workflow");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("blocks formula-like values in generic value previews", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Change Data B2",
        mode: "preview_update",
        target: { sheetName: "Data", range: "B2" },
        values: valuePatch("Data", "B2", [["=SUM(A1:A2)"]])
      });

      expect(result.status).toBe("VALIDATION_FAILED");
      expect(result.nextAction).toBe("manual_review");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("does not auto-apply sparse broad writes", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Change Report B1:C4",
        target: { sheetName: "Report", range: "B1:C4" },
        values: { patches: [{ target: { sheetName: "Report", range: "B1:C4" }, values: [[1]] }] }
      });

      expect(result.status).toBe("VALIDATION_FAILED");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("rejects legacy direct values.values updates with a canonical patch example", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Update Data B2",
        mode: "preview_update",
        intent: { action: "write_values" },
        target: { sheetName: "Data", range: "B2" },
        values: { values: [[123]] }
      });

      expect(result.status).toBe("NEEDS_INPUT");
      expect((result.answer as any).kind).toBe("canonical_patch_required");
      expect((result.answer as any).code).toBe("CANONICAL_PATCH_REQUIRED");
      expect((result.answer as any).example.values.patches[0]).toMatchObject({
        target: { sheetName: "Sales", range: "E2" },
        values: [["Reviewed"]]
      });
      expect(result.warnings).toContain("CANONICAL_PATCH_REQUIRED");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("rejects legacy top-level structured payloads for patchable direct mutation families", async () => {
      const cases: Array<Parameters<AgentOrchestrator["run"]>[0]> = [
        {
          request: "Write formula to Report A12",
          mode: "preview_update",
          intent: { action: "write_formulas" },
          target: { sheetName: "Report", range: "A12" },
          values: { formulas: [["=SUM(B1:B10)"]] }
        },
        {
          request: "Format Data C2 as currency",
          mode: "preview_update",
          intent: { action: "write_number_formats" },
          target: { sheetName: "Data", range: "C2" },
          values: { numberFormat: "$#,##0.00" }
        },
        {
          request: "Style Data A1:D1",
          mode: "preview_update",
          intent: { action: "format_range" },
          target: { sheetName: "Data", range: "A1:D1" },
          values: { style: { fillColor: "#1A3C6E" } }
        },
        {
          request: "Add dropdown to Data D2:D4",
          mode: "preview_update",
          intent: { action: "write_data_validation" },
          target: { sheetName: "Data", range: "D2:D4" },
          values: { options: ["20GP", "40GP", "40HQ"] }
        },
        {
          request: "Add conditional formatting",
          mode: "preview_update",
          intent: { action: "write_conditional_formatting" },
          target: { sheetName: "Data", range: "A2:D4" },
          values: { formula: "=$D2=\"40HQ\"", style: { fillColor: "#FFFF00" } }
        }
      ];

      for (const input of cases) {
        const runtime = new FakeAgentRuntime();
        const agent = new AgentOrchestrator(runtime as any);
        const result = await agent.run(input);

        expect(result.status, input.intent?.action).toBe("NEEDS_INPUT");
        expect((result.answer as any).kind, input.intent?.action).toBe("canonical_patch_required");
        expect(result.warnings, input.intent?.action).toContain("CANONICAL_PATCH_REQUIRED");
        expect(runtime.writeBatchCount, input.intent?.action).toBe(0);
      }
    });

  it("reports explicit route metadata for auto mutation routing", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const result = await agent.run({
        request: "Update Data B2",
        values: { patches: [{ target: { sheetName: "Data", range: "B2" }, values: [[123]] }] }
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.telemetry.routeMode).toBe("preview_update");
      expect(result.telemetry.routeMatchedRule).toBe("mutation.keyword");
      expect(result.telemetry.routeConfidence).toBeGreaterThan(0);
      expect(result.telemetry.operationRisk).toBe("small_value_write");
      expect(result.telemetry.targetFingerprintStatus).toBe("matched");
      expect(result.telemetry.autoApplied).toBe(true);
    });

  it("carries operation risk telemetry through apply", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Format header row on Data",
        mode: "preview_update",
        target: { sheetName: "Data", range: "A1:D1" }
      });
      const applied = await agent.run({
        request: "Apply style update",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect(preview.telemetry.operationRisk).toBe("safe_format");
      expect(applied.status).toBe("SUCCESS");
      expect(applied.telemetry.operationRisk).toBe("safe_format");
      expect((applied.answer as any).operationRisk).toBe("safe_format");
    });

  it("routes natural autofilter and border requests to range handlers instead of sheet creation", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const filter = await agent.run({
        request: "Add autofilter to the header row of Booking sheet range A1:X7.",
        mode: "preview_update",
        target: { sheetName: "Data", range: "A1:D4" }
      });
      const borders = await agent.run({
        request: "Add borders to Booking sheet range A1:X7. Set EdgeBottom and EdgeRight as thin continuous black borders.",
        mode: "preview_update",
        target: { sheetName: "Data", range: "A1:D4" }
      });
      const explicitFormat = await new AgentOrchestrator(new FakeAgentRuntime() as any).run({
        request: "format_range action: Apply borders to range Data!A1:D4. Set EdgeBottom and EdgeRight borders.",
        mode: "preview_update",
        target: { sheetName: "Data", range: "A1:D4" }
      });

      expect(filter.status).toBe("PREVIEW_READY");
      expect((filter.answer as any).kind).toBe("filter_preview");
      expect(filter.telemetry.actionHandlerId).toBe("filter_range");
      expect(filter.telemetry.operationRisk).toBe("safe_format");
      expect((filter.answer as any).kind).not.toBe("sheet_create_preview");

      expect(borders.status).toBe("PREVIEW_READY");
      expect((borders.answer as any).kind).toBe("style_preview");
      expect(borders.telemetry.actionHandlerId).toBe("format_range");
      expect((borders.answer as any).kind).not.toBe("sheet_create_preview");

      expect(explicitFormat.status).toBe("PREVIEW_READY");
      expect((explicitFormat.answer as any).kind).toBe("style_preview");
      expect(explicitFormat.telemetry.actionHandlerId).toBe("format_range");
    });

  it("redirects lookup-only filter_range requests to query_rows", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const redirected = await agent.run({
        request: "Show rows where Status = Open",
        mode: "preview_update",
        intent: { action: "filter_range" },
        target: { sheetName: "Data", tableName: "Transactions" },
        values: { where: [{ column: "Status", op: "=", value: "Open" }] }
      });

      expect(redirected.status).toBe("NEEDS_WORKFLOW_REDIRECT");
      expect((redirected.answer as any)).toMatchObject({
        kind: "query_rows_redirect",
        suggestedIntentAction: "query_rows"
      });
      expect(redirected.suggestedOperation).toMatchObject({
        mode: "answer",
        intent: { action: "query_rows" },
        values: { where: [{ column: "Status", op: "=", value: "Open" }] }
      });
      expect(redirected.agentInstruction).toContain("filter_range only when the user explicitly asks");
      expect(runtime.writeBatchCount).toBe(0);
    });

  it("clears range autofilters without requiring table filter criteria", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Remove all filters from Booking sheet range A1:X7.",
        mode: "preview_update",
        target: { sheetName: "Data", range: "A1:D4" }
      });
      const applied = await agent.run({
        request: "Apply clear filters",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe("filter_clear_preview");
      expect(preview.telemetry.actionHandlerId).toBe("filter_range");
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.lastBatchOperations[0]).toMatchObject({ kind: "range.clear_autofilter", target: { sheetName: "Data", address: "A1:D4" } });
      expect(runtime.tableMethodCalls).toHaveLength(0);
    });

  it("treats clear_table_filters with a range target as range autofilter clearing", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Clear all applied filters",
        mode: "preview_update",
        intent: { action: "clear_table_filters" },
        target: { sheetName: "Data", range: "A1:D4" }
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe("filter_clear_preview");
      expect((preview.changes ?? [])[0]).toMatchObject({ sheetName: "Data", range: "A1:D4", after: "filters cleared" });
    });

  it("keeps parse_dates patch targets exact instead of broadening to the whole sheet", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);

      const preview = await agent.run({
        request: "Parse text dates in Data columns A and C rows 2-4.",
        mode: "preview_update",
        intent: { action: "parse_dates" },
        target: { sheetName: "Data", range: "A1:D4" },
        values: {
          patches: [
            { target: { sheetName: "Data", range: "A2:A4" }, numberFormat: "dd/mm/yyyy" },
            { target: { sheetName: "Data", range: "C2:C4" }, numberFormat: "dd/mm/yyyy" }
          ]
        }
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).kind).toBe("cleaning_preview");
      expect((preview.answer as any).grouped).toBe(true);
      expect((preview.answer as any).requests.map((request: any) => request.address)).toEqual(["A2:A4", "C2:C4"]);
      expect(preview.proof.map((proof) => proof.range)).toEqual(["A2:A4", "C2:C4"]);
      expect(preview.changes?.map((change) => change.range)).toEqual(["A2:A4", "C2:C4"]);

      const applied = await agent.run({
        request: "Apply exact date parsing.",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(applied.status).toBe("SUCCESS");
      expect(runtime.runtimeMethodCalls["clean.parse_dates"]).toBe(2);
    });

  it("normalizes agent-friendly table filter shorthands before apply", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      agent.metadataCache.set(createCachedMetadata("wbctx_filter_alias"));

      const preview = await agent.run({
        request: "Filter the Transactions table to Status Open",
        mode: "preview_update",
        workbookContextId: "wbctx_filter_alias",
        intent: { action: "filter_range" },
        target: { tableName: "Transactions" },
        values: { filters: [{ column: "Status", filterType: "text", value: "Open" }] }
      });
      const applied = await agent.run({
        request: "Apply filter",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.tableMethodCalls.at(-1)).toEqual({
        method: "table.apply_filters",
        request: {
          workbookId: "workbook_agent_unit",
          tableName: "Transactions",
          filters: [{ column: "Status", criteria: { filterOn: "Values", values: ["Open"] } }]
        }
      });
    });

  it("normalizes criterion table filters before apply", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      agent.metadataCache.set(createCachedMetadata("wbctx_filter_criterion"));

      const preview = await agent.run({
        request: "Filter the Transactions table to Status Open",
        mode: "preview_update",
        workbookContextId: "wbctx_filter_criterion",
        intent: { action: "filter_range" },
        target: { tableName: "Transactions" },
        values: { filters: [{ column: "Status", criterion: "Open" }] }
      });
      const applied = await agent.run({
        request: "Apply filter",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(applied.status).toBe("SUCCESS");
      expect(runtime.tableMethodCalls.at(-1)?.request.filters[0].criteria).toEqual({ filterOn: "Values", values: ["Open"] });
    });

  it("accepts rowStart and rowEnd aliases for compact table reads", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      agent.metadataCache.set(createCachedMetadata("wbctx_table_page_alias"));

      const result = await agent.run({
        request: "Read rows 2 through 2 from Transactions",
        mode: "answer",
        workbookContextId: "wbctx_table_page_alias",
        target: { tableName: "Transactions" },
        values: { rowStart: 2, rowEnd: 2 }
      });

      expect(result.status).toBe("SUCCESS");
      expect((result.answer as any).rowOffset).toBe(1);
      expect((result.answer as any).rowLimit).toBe(1);
      expect((result.answer as any).values).toEqual([["2026-06-02", "A-101", 456, "Closed"]]);
      const resultId = String((result.answer as any).resultUri).split("/").pop()!;
      expect((agent.getResultResource(resultId) as any).answer.values).toEqual([["2026-06-02", "A-101", 456, "Closed"]]);
    });

  it("auto-routes structured values with an explicit target to a scoped write even without intent.action", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_auto_structured_value_write");
      metadata.sheets = [
        ...metadata.sheets,
        {
          id: "sheet:Dropdown Lists",
          name: "Dropdown Lists",
          index: 9,
          usedRange: "A1:B29",
          rowCount: 29,
          columnCount: 2,
          kind: "reference",
          headers: [],
          tableIds: [],
          sectionIds: [],
          summaryBlockIds: [],
          formulaRegionIds: []
        }
      ];
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "",
        mode: "auto",
        workbookContextId: metadata.workbookContextId,
        target: { sheetName: "Dropdown Lists", range: "B29" },
        values: valuePatch("Dropdown Lists", "B29", [["owner_cash_topup"]])
      });

      expect(result.status).toBe("SUCCESS");
      expect(result.taskOutcome).toBe("apply_complete");
      expect(result.telemetry.routeMode).toBe("preview_update");
      expect(runtime.lastBatchOperations[0]).toMatchObject({
        kind: "range.write_values_many",
        entries: [{ target: { sheetName: "Dropdown Lists", address: "B29" }, values: [["owner_cash_topup"]] }]
      });
    });

  it("rejects malformed direct value matrices before auto-applying structured writes", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      const metadata = createCachedMetadata("wbctx_auto_structured_value_shape");
      metadata.sheets = [
        ...metadata.sheets,
        {
          id: "sheet:Dropdown Lists",
          name: "Dropdown Lists",
          index: 9,
          usedRange: "A1:B29",
          rowCount: 29,
          columnCount: 2,
          kind: "reference",
          headers: [],
          tableIds: [],
          sectionIds: [],
          summaryBlockIds: [],
          formulaRegionIds: []
        }
      ];
      agent.metadataCache.set(metadata);

      const result = await agent.run({
        request: "",
        mode: "auto",
        workbookContextId: metadata.workbookContextId,
        target: { sheetName: "Dropdown Lists", range: "B29" },
        values: valuePatch("Dropdown Lists", "B29", [[null, "owner_cash_topup"]])
      });

      expect(result.status).toBe("VALIDATION_FAILED");
      expect(result.summary).toContain("shape mismatch");
      expect(result.telemetry.routeMode).toBe("preview_update");
      expect(runtime.writeBatchCount).toBe(0);
      expect((result.answer as any)?.kind).not.toBe("range_profile");
    });

  it("normalizes structured table sort values", async () => {
      const runtime = new FakeAgentRuntime();
      const agent = new AgentOrchestrator(runtime as any);
      agent.metadataCache.set(createCachedMetadata("wbctx_sort_alias"));

      const preview = await agent.run({
        request: "Sort Transactions by status descending",
        mode: "preview_update",
        workbookContextId: "wbctx_sort_alias",
        intent: { action: "sort_table" },
        target: { tableName: "Transactions" },
        values: { sortBy: "Status", direction: "descending" }
      });
      const applied = await agent.run({
        request: "Apply sort",
        mode: "apply_update",
        operationId: preview.operationId,
        confirmationToken: preview.confirmationToken
      });

      expect(preview.status).toBe("PREVIEW_READY");
      expect((preview.answer as any).sortField).toBe("Status");
      expect(applied.status).toBe("SUCCESS");
      expect(runtime.tableMethodCalls.at(-1)?.request.fields).toEqual([{ key: 3, ascending: false }]);
    });
});
