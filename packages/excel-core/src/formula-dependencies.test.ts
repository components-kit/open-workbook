import { describe, expect, it } from "vitest";
import type { FormulaPatternResponse, TableInfo, WorkbookId } from "@components-kit/open-workbook-protocol";
import { buildFormulaDependencyGraph, traceDependents, tracePrecedents } from "./formula-dependencies.js";

describe("formula dependency graph", () => {
  it("extracts sheet-local and cross-sheet precedents", () => {
    const workbookId = "workbook_formula_graph" as WorkbookId;
    const graph = buildFormulaDependencyGraph(patterns(workbookId));

    const precedents = tracePrecedents(graph, "Report", "B2");

    expect(precedents.nodes.map((node) => `${node.sheetName}!${node.address}`).sort()).toEqual(["Raw Data!C3:D4", "Report!A1"]);
    expect(graph.edges).toHaveLength(2);
  });

  it("finds dependents for a referenced range", () => {
    const workbookId = "workbook_formula_dependents" as WorkbookId;
    const graph = buildFormulaDependencyGraph(patterns(workbookId));

    const dependents = traceDependents(graph, "Raw Data", "C4");

    expect(dependents.nodes.map((node) => `${node.sheetName}!${node.address}`)).toEqual(["Report!B2"]);
  });

  it("extracts structured table references as table nodes", () => {
    const workbookId = "workbook_formula_structured" as WorkbookId;
    const graph = buildFormulaDependencyGraph(patterns(workbookId, "=SUM(Transactions[Amount])+COUNTIFS(Transactions[Status],\"Open\")"));

    const tableNodes = graph.nodes.filter((node) => node.kind === "table");

    expect(tableNodes.map((node) => `${node.tableName}[${node.structuredReference}]`).sort()).toEqual(["Transactions[Amount]", "Transactions[Status]"]);
    expect(graph.warnings.some((warning) => warning.code === "FORMULA_STRUCTURED_REFERENCES_PARSED")).toBe(true);
  });

  it("extracts external workbook references as external nodes", () => {
    const workbookId = "workbook_formula_external" as WorkbookId;
    const graph = buildFormulaDependencyGraph(patterns(workbookId, "='[Prior.xlsx]Summary'!$A$1+SUM(A1:A2)"));

    const externalNodes = graph.nodes.filter((node) => node.kind === "external");
    const localPrecedents = graph.nodes.filter((node) => node.kind === "range" && node.formula === undefined);

    expect(externalNodes).toHaveLength(1);
    expect(externalNodes[0]?.externalWorkbook).toBe("Prior.xlsx");
    expect(externalNodes[0]?.externalReference).toBe("Summary!A1");
    expect(localPrecedents.map((node) => `${node.sheetName}!${node.address}`)).toEqual(["Report!A1:A2"]);
    expect(graph.warnings.some((warning) => warning.code === "FORMULA_EXTERNAL_REFERENCES_PARSED")).toBe(true);
  });

  it("resolves structured table references to data-body ranges when table metadata is provided", () => {
    const workbookId = "workbook_formula_table_resolve" as WorkbookId;
    const graph = buildFormulaDependencyGraph(patterns(workbookId, "=SUM(Transactions[Amount])"), {
      tables: [transactionsTable(workbookId)]
    });

    const rangeNodes = graph.nodes.filter((node) => node.kind === "range" && node.formula === undefined);
    const tableNodes = graph.nodes.filter((node) => node.kind === "table");

    expect(tableNodes).toHaveLength(1);
    expect(rangeNodes.map((node) => `${node.sheetName}!${node.address}`)).toEqual(["Transactions!C2:C10"]);
  });

  it("resolves structured table special items and column spans", () => {
    const workbookId = "workbook_formula_table_special" as WorkbookId;
    const graph = buildFormulaDependencyGraph(
      patterns(workbookId, "=SUM(Transactions[[#Headers],[Amount]])+SUM(Transactions[[#Totals],[Amount]])+SUM(Transactions[[#All],[Amount]:[Memo]])"),
      {
        tables: [transactionsTable(workbookId)]
      }
    );

    const rangeNodes = graph.nodes
      .filter((node) => node.kind === "range" && node.formula === undefined)
      .map((node) => `${node.sheetName}!${node.address}`)
      .sort();

    expect(rangeNodes).toEqual(["Transactions!C1", "Transactions!C11", "Transactions!C1:D11"]);
  });

  it("extracts whole-column and spill references conservatively", () => {
    const workbookId = "workbook_formula_column_spill" as WorkbookId;
    const graph = buildFormulaDependencyGraph(patterns(workbookId, "=SUM(A:C)+SUM(E5#)"));

    const rangeNodes = graph.nodes.filter((node) => node.kind === "range" && node.formula === undefined);

    expect(rangeNodes.map((node) => `${node.sheetName}!${node.address}`).sort()).toEqual(["Report!A:C", "Report!E5"]);
    expect(graph.warnings.some((warning) => warning.code === "FORMULA_SPILL_RANGE_UNRESOLVED")).toBe(true);
  });

  it("expands spill references when spill metadata is available", () => {
    const workbookId = "workbook_formula_spill_expand" as WorkbookId;
    const graph = buildFormulaDependencyGraph({
      ...patterns(workbookId, "=SUM(E5#)"),
      spillRanges: [{ sheetName: "Report", anchorAddress: "E5", spillAddress: "E5:G12" }]
    });

    const rangeNodes = graph.nodes.filter((node) => node.kind === "range" && node.formula === undefined);

    expect(rangeNodes.map((node) => `${node.sheetName}!${node.address}`)).toEqual(["Report!E5:G12"]);
    expect(graph.warnings.some((warning) => warning.code === "FORMULA_SPILL_RANGE_UNRESOLVED")).toBe(false);
  });
});

function patterns(workbookId: WorkbookId, formula = "=SUM(A1,'Raw Data'!$C$3:$D$4)"): FormulaPatternResponse {
  return {
    workbookId,
    sheetName: "Report",
    address: "B2:B2",
    capturedAt: new Date().toISOString(),
    rowCount: 1,
    columnCount: 1,
    formulaCount: 1,
    formulas: [[formula]],
    patternMatrix: [["hash"]],
    patterns: [],
    cells: [
      {
        rowIndex: 0,
        columnIndex: 0,
        formula,
        patternHash: "hash"
      }
    ],
    warnings: []
  };
}

function transactionsTable(workbookId: WorkbookId): TableInfo {
  return {
    workbookId,
    tableName: "Transactions",
    sheetName: "Transactions",
    address: "A1:D11",
    rowCount: 11,
    columnCount: 4,
    showHeaders: true,
    showTotals: true,
    columns: [
      { index: 0, name: "Date" },
      { index: 1, name: "Status" },
      { index: 2, name: "Amount" },
      { index: 3, name: "Memo" }
    ]
  };
}
