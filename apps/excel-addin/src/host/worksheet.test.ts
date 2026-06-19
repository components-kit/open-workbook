import { describe, expect, it } from "vitest";
import { BATCH_OPERATION_KINDS, getHostMethod } from "./registry.js";

describe("worksheet host operations", () => {
  it("are routed through the batch executor", () => {
    expect(getHostMethod("operation.execute_batch")).toBeDefined();
    expect(BATCH_OPERATION_KINDS).toContain("sheet.create");
    expect(BATCH_OPERATION_KINDS).toContain("sheet.set_tab_color");
  });
});
