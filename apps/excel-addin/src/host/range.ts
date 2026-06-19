import {
  findBlankCells,
  findFormulaErrors,
  readRangeComments,
  readRangeConditionalFormatting,
  readRangeDataValidation,
  readRangeHyperlinks,
  readRangeMergedCells,
  readRangeNotes,
  searchRange
} from "./executor-core.js";

export const rangeHostOperations = {
  readRangeHyperlinks: (params: unknown) => readRangeHyperlinks(params as Parameters<typeof readRangeHyperlinks>[0]),
  readRangeComments: (params: unknown) => readRangeComments(params as Parameters<typeof readRangeComments>[0]),
  readRangeNotes: (params: unknown) => readRangeNotes(params as Parameters<typeof readRangeNotes>[0]),
  readRangeMergedCells: (params: unknown) => readRangeMergedCells(params as Parameters<typeof readRangeMergedCells>[0]),
  readRangeDataValidation: (params: unknown) => readRangeDataValidation(params as Parameters<typeof readRangeDataValidation>[0]),
  readRangeConditionalFormatting: (params: unknown) => readRangeConditionalFormatting(params as Parameters<typeof readRangeConditionalFormatting>[0]),
  searchRange: (params: unknown) => searchRange(params as Parameters<typeof searchRange>[0]),
  findBlankCells: (params: unknown) => findBlankCells(params as Parameters<typeof findBlankCells>[0]),
  findFormulaErrors: (params: unknown) => findFormulaErrors(params as Parameters<typeof findFormulaErrors>[0])
};
