import { describe, expect, it } from "vitest";
import { createCachedMetadata } from "./agent-orchestrator.test-support.js";
import { WorkbookMetadataCache } from "./workbook-metadata-cache.js";

describe("WorkbookMetadataCache context state", () => {
  it("tracks facet freshness separately from metadata entries", () => {
    const cache = new WorkbookMetadataCache();
    const metadata = createCachedMetadata("wbctx_cache_facets");

    cache.set(metadata);
    const initial = cache.getContextState(metadata.workbookContextId);
    const stale = cache.markFacetsStale(metadata.workbookContextId, ["values", "aggregates"], ["Data!D2"]);

    expect(initial?.contextVersion).toBe(1);
    expect(initial?.freshness).toMatchObject({
      status: "fresh",
      staleFacets: [],
      confidence: 1
    });
    expect(stale?.contextVersion).toBe(2);
    expect(stale?.freshness).toMatchObject({
      status: "mostly_fresh",
      staleFacets: ["values", "aggregates"],
      staleRanges: ["Data!D2"]
    });
    expect(stale?.freshness.freshFacets).not.toContain("values");
  });

  it("stores operation journal entries with context version and cloned changes", () => {
    const cache = new WorkbookMetadataCache();
    const metadata = createCachedMetadata("wbctx_cache_journal");
    cache.set(metadata);
    cache.markFacetsStale(metadata.workbookContextId, ["values"], ["Data!D2"], 1000);

    const entry = cache.appendJournalEntry(metadata.workbookContextId, {
      operationId: "op_journal",
      affectedRanges: ["Data!D2"],
      affectedFacets: ["values"],
      invalidatedFacets: ["aggregates", "formulaResults"],
      preservedFacets: ["schema", "headers", "fieldContext", "validation"],
      changes: [{ sheetName: "Data", range: "D2", before: "Open", after: "Closed" }],
      cacheAction: "recorded"
    }, 2000);
    const state = cache.getContextState(metadata.workbookContextId);

    expect(entry).toMatchObject({
      operationId: "op_journal",
      workbookContextId: metadata.workbookContextId,
      contextVersion: 2,
      appliedAt: 2000,
      cacheAction: "recorded"
    });
    expect(state?.journal).toHaveLength(1);
    state!.journal[0]!.affectedRanges.push("mutated");
    expect(cache.getContextState(metadata.workbookContextId)?.journal[0]?.affectedRanges).toEqual(["Data!D2"]);
  });

  it("checks required facet freshness without treating the whole context as stale", () => {
    const cache = new WorkbookMetadataCache();
    const metadata = createCachedMetadata("wbctx_cache_freshness_check");
    cache.set(metadata);
    cache.markFacetsStale(metadata.workbookContextId, ["values", "aggregates"], ["Data!D2"]);

    expect(cache.checkFacetFreshness(metadata.workbookContextId, ["schema", "headers", "validation"])).toMatchObject({
      status: "fresh",
      requiresRead: false,
      freshRequiredFacets: ["schema", "headers", "validation"],
      staleRequiredFacets: []
    });
    expect(cache.checkFacetFreshness(metadata.workbookContextId, ["schema", "values", "aggregates"])).toMatchObject({
      status: "mostly_fresh",
      requiresRead: true,
      freshRequiredFacets: ["schema"],
      staleRequiredFacets: ["values", "aggregates"],
      staleRanges: ["Data!D2"]
    });
  });

  it("plans narrow refreshes for only missing or stale required facets", () => {
    const cache = new WorkbookMetadataCache();
    const metadata = createCachedMetadata("wbctx_cache_refresh_plan");
    cache.set(metadata);
    cache.markFacetsStale(metadata.workbookContextId, ["values", "aggregates"], ["Data!B2"]);

    expect(cache.planContextRefresh(metadata.workbookContextId, ["schema", "headers", "values"])).toMatchObject({
      status: "mostly_fresh",
      requiredFacets: ["schema", "headers", "values"],
      cacheFacets: ["schema", "headers"],
      liveFacets: ["values"],
      missingFacets: [],
      staleFacets: ["values"],
      staleRanges: ["Data!B2"],
      requiresRead: true,
      readStrategy: "read_stale_facets",
      reason: expect.stringContaining("stale context facets"),
      confidence: 2 / 3
    });
    expect(cache.planContextRefresh("missing_ctx", ["schema", "validation"])).toMatchObject({
      status: "stale",
      cacheFacets: [],
      liveFacets: ["schema", "validation"],
      missingFacets: ["schema", "validation"],
      requiresRead: true,
      readStrategy: "read_missing_facets"
    });
  });

  it("records optimistic value updates without making the values facet stale", () => {
    const cache = new WorkbookMetadataCache();
    const metadata = createCachedMetadata("wbctx_cache_optimistic_values");
    cache.set(metadata);

    const updated = cache.applyOptimisticValueChanges(metadata.workbookContextId, "op_patch", [
      { sheetName: "Data", cell: "B2", range: "B2", before: 123, after: 999 },
      { sheetName: "Data", range: "C2", before: "Open", after: "Closed" },
      { sheetName: "Data", range: "D2:D500", after: { kind: "range_write_summary", cellCount: 499 } }
    ], 3000);
    cache.markFacetsStale(metadata.workbookContextId, ["aggregates", "formulaResults"], ["Data!B2"]);

    expect(updated).toHaveLength(2);
    expect(cache.getOptimisticValue(metadata.workbookContextId, "Data", "B2")).toMatchObject({
      sheetName: "Data",
      range: "B2",
      value: 999,
      before: 123,
      operationId: "op_patch",
      updatedAt: 3000
    });
    expect(cache.getOptimisticValue(metadata.workbookContextId, "Data", "D2:D500")).toBeUndefined();
    expect(cache.getContextState(metadata.workbookContextId)?.freshness.staleFacets).not.toContain("values");
    expect(cache.checkFacetFreshness(metadata.workbookContextId, ["values"])).toMatchObject({
      status: "fresh",
      requiresRead: false
    });
  });

  it("removes context state when metadata is deleted", () => {
    const cache = new WorkbookMetadataCache();
    const metadata = createCachedMetadata("wbctx_cache_delete");
    cache.set(metadata);

    expect(cache.getContextState(metadata.workbookContextId)).toBeDefined();
    cache.deleteByContextId(metadata.workbookContextId);

    expect(cache.getContextState(metadata.workbookContextId)).toBeUndefined();
  });
});
