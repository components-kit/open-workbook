import { OPEN_WORKBOOK_VERSION } from "@components-kit/open-workbook-protocol";
import { describe, expect, it } from "vitest";
import { OfficeJsEngine } from "./office-js-engine.js";

describe("OfficeJsEngine version", () => {
  it("uses the shared package version", async () => {
    const engine = new OfficeJsEngine();

    expect(engine.version).toBe(OPEN_WORKBOOK_VERSION);
    await expect(engine.getCapabilities()).resolves.toMatchObject({
      engine: {
        version: OPEN_WORKBOOK_VERSION
      }
    });
  });
});

