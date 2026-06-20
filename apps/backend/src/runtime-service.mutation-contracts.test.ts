import { readFileSync } from "node:fs";
import { describe, expect, it } from "vitest";

describe("RuntimeService mutation contracts", () => {
  it("routes direct add-in workbook mutations through transaction envelopes", () => {
    const source = readFileSync(new URL("./runtime-service.ts", import.meta.url), "utf8");
    const required = [
      { runtimeMethod: "deleteChart", addinMethod: "chart.delete" },
      { runtimeMethod: "repairSheetFromTemplate", addinMethod: "template.repair" }
    ];

    for (const requirement of required) {
      const body = methodBody(source, requirement.runtimeMethod);

      expect(body).toContain(`client.request("${requirement.addinMethod}"`);
      expect(body).toContain("return this.applyDirectTransaction(");
      expect(body.indexOf("return this.applyDirectTransaction(")).toBeLessThan(body.indexOf(`client.request("${requirement.addinMethod}"`));
    }
  });
});

function methodBody(source: string, methodName: string): string {
  const start = source.indexOf(`async ${methodName}(`);
  expect(start).toBeGreaterThanOrEqual(0);
  const nextMethod = source.indexOf("\n  async ", start + 1);
  return source.slice(start, nextMethod === -1 ? source.length : nextMethod);
}
