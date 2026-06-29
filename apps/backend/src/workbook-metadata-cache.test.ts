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

  it("removes context state when metadata is deleted", () => {
    const cache = new WorkbookMetadataCache();
    const metadata = createCachedMetadata("wbctx_cache_delete");
    cache.set(metadata);

    expect(cache.getContextState(metadata.workbookContextId)).toBeDefined();
    cache.deleteByContextId(metadata.workbookContextId);

    expect(cache.getContextState(metadata.workbookContextId)).toBeUndefined();
  });
});
