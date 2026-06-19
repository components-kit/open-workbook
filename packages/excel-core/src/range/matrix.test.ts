import { describe, expect, it } from "vitest";
import { cloneMatrix, matrixCellCount, padMatrixRows, rectangularize } from "./matrix.js";

describe("matrix helpers", () => {
  it("counts cells across ragged rows", () => {
    expect(matrixCellCount([[1, 2], [3]])).toBe(3);
  });

  it("clones row arrays", () => {
    const input = [[1], [2]];
    const cloned = cloneMatrix(input);
    cloned[0]![0] = 9;
    expect(input[0]![0]).toBe(1);
  });

  it("rectangularizes ragged rows", () => {
    expect(rectangularize([[1], [2, 3]])).toEqual([[1, null], [2, 3]]);
  });

  it("pads rows and columns", () => {
    expect(padMatrixRows([[1]], 2, 3)).toEqual([[1, null, null], [null, null, null]]);
  });
});
