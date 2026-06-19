import { describe, expect, it } from "vitest";
import { cellCount, cellCountFromAddress, columnNameToNumber, formatA1Address, formatA1Cell, numberToColumnName, parseA1Address, rangesOverlap, stripSheetName, tryParseA1Address } from "./range-address.js";

describe("range address utilities", () => {
  it("converts column names and numbers", () => {
    expect(columnNameToNumber("A")).toBe(1);
    expect(columnNameToNumber("Z")).toBe(26);
    expect(columnNameToNumber("AA")).toBe(27);
    expect(numberToColumnName(27)).toBe("AA");
  });

  it("parses single-cell and rectangular A1 ranges", () => {
    expect(parseA1Address("B2")).toMatchObject({
      startRow: 2,
      startColumn: 2,
      endRow: 2,
      endColumn: 2
    });
    expect(parseA1Address("'Accounting Jan'!A1:D20")).toMatchObject({
      sheetName: "Accounting Jan",
      startRow: 1,
      startColumn: 1,
      endRow: 20,
      endColumn: 4
    });
  });

  it("formats and counts ranges", () => {
    const parsed = parseA1Address("'Accounting Jan'!A1:D20");
    expect(formatA1Address(parsed)).toBe("'Accounting Jan'!A1:D20");
    expect(cellCount("A1:D20")).toBe(80);
  });

  it("supports forgiving range helper calls", () => {
    expect(stripSheetName("'Accounting Jan'!A1:D20")).toBe("A1:D20");
    expect(tryParseA1Address("not a range")).toBeUndefined();
    expect(cellCountFromAddress("'Accounting Jan'!A1:D20")).toBe(80);
    expect(cellCountFromAddress("not a range")).toBeUndefined();
    expect(rangesOverlap("A1:C3", "C3:D4")).toBe(true);
    expect(rangesOverlap("A1:C3", "D4:E5")).toBe(false);
  });

  it("formats one-based cell positions as A1 addresses", () => {
    expect(formatA1Cell(1, 1)).toBe("A1");
    expect(formatA1Cell(4, 2)).toBe("B4");
    expect(formatA1Cell(1048576, 16384)).toBe("XFD1048576");
  });

  it("rejects invalid one-based cell positions", () => {
    expect(() => formatA1Cell(0, 1)).toThrow("Invalid row number");
    expect(() => formatA1Cell(1, 0)).toThrow("Invalid column number");
  });
});
