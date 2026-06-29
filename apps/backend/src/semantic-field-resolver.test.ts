import { describe, expect, it } from "vitest";
import { createCachedMetadata } from "./agent-orchestrator.test-support.js";
import { resolveSemanticField, resolveSemanticFields } from "./semantic-field-resolver.js";

describe("semantic field resolver", () => {
  it("resolves exact and synonym-backed field terms within a table scope", () => {
    const metadata = createCachedMetadata("wbctx_field_resolver");

    expect(resolveSemanticField(metadata, "payment status", { sheetName: "Data", tableName: "Transactions" })).toMatchObject({
      ambiguous: false,
      best: {
        field: "Status",
        sheetName: "Data",
        tableName: "Transactions",
        columnLetter: "D",
        evidence: expect.arrayContaining(["status semantic role match"])
      }
    });
    expect(resolveSemanticField(metadata, "total cost", { sheetName: "Data", tableName: "Transactions" }).best).toMatchObject({
      field: "Amount",
      columnLetter: "C"
    });
  });

  it("uses target scope to avoid cross-sheet candidates when requested", () => {
    const metadata = createCachedMetadata("wbctx_field_scope");
    const scoped = resolveSemanticField(metadata, "tier", { sheetName: "Customer Master" });

    expect(scoped.best).toMatchObject({
      field: "Tier",
      sheetName: "Customer Master",
      range: "A1:B1"
    });
    expect(scoped.candidates.every((candidate) => candidate.sheetName === "Customer Master")).toBe(true);
  });

  it("resolves multiple query terms in input order", () => {
    const metadata = createCachedMetadata("wbctx_field_multi");
    const resolutions = resolveSemanticFields(metadata, ["date", "customer", "amount"], { sheetName: "Data", tableName: "Transactions" });

    expect(resolutions.map((resolution) => resolution.best?.field)).toEqual(["Date", "Account", "Amount"]);
  });
});
