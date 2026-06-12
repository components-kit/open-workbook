import { describe, expect, it } from "vitest";
import { chunkMatrixRows } from "./chunking.js";

describe("chunkMatrixRows", () => {
  it("keeps small matrices in one chunk", () => {
    const chunks = chunkMatrixRows(
      [
        [1, 2],
        [3, 4]
      ],
      10
    );

    expect(chunks).toEqual([
      {
        rowOffset: 0,
        rows: [
          [1, 2],
          [3, 4]
        ]
      }
    ]);
  });

  it("splits large matrices by whole rows", () => {
    const chunks = chunkMatrixRows(
      [
        [1, 2],
        [3, 4],
        [5, 6],
        [7, 8],
        [9, 10]
      ],
      4
    );

    expect(chunks.map((chunk) => chunk.rowOffset)).toEqual([0, 2, 4]);
    expect(chunks.map((chunk) => chunk.rows)).toEqual([
      [
        [1, 2],
        [3, 4]
      ],
      [
        [5, 6],
        [7, 8]
      ],
      [[9, 10]]
    ]);
  });

  it("uses one row per chunk when a row exceeds the limit", () => {
    const chunks = chunkMatrixRows(
      [
        [1, 2, 3],
        [4, 5, 6]
      ],
      2
    );

    expect(chunks.map((chunk) => chunk.rowOffset)).toEqual([0, 1]);
  });

  it("rejects ragged matrices", () => {
    expect(() => chunkMatrixRows([[1], [2, 3]], 10)).toThrow("Matrix must be rectangular");
  });
});
