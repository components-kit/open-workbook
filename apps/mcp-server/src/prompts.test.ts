import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

describe("MCP prompts", () => {
  it("documents the generic field/value styled-table workflow as one preview/apply path", () => {
    const source = readFileSync(new URL("./prompts.ts", import.meta.url), "utf8");

    expect(source).toContain("excel.prompts.field_value_image_to_styled_table");
    expect(source).toContain("excel.prompts.booking_image_to_styled_table");
    expect(source).toContain("field/value data from OCR, screenshots, forms, invoices, shipment documents, or booking images");
    expect(source).toContain("replace_range_with_styled_table");
    expect(source).toContain("Do not split clear, value write, autofit, and style copy");
  });
});
