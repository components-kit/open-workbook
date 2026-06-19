import {
  convertFormulasToValues,
  copyFormulaPatterns,
  fillFormulaPattern,
  readFormulaPatterns
} from "./executor-core.js";

export const formulaHostOperations = {
  readFormulaPatterns: (params: unknown) => readFormulaPatterns(params as Parameters<typeof readFormulaPatterns>[0]),
  copyFormulaPatterns: (params: unknown) => copyFormulaPatterns(params as Parameters<typeof copyFormulaPatterns>[0]),
  fillFormulaPattern: (params: unknown) => fillFormulaPattern(params as Parameters<typeof fillFormulaPattern>[0]),
  convertFormulasToValues: (params: unknown) => convertFormulasToValues(params as Parameters<typeof convertFormulasToValues>[0])
};
