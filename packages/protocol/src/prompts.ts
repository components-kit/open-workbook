import type { CatalogStatus } from "./tools.js";

export interface PromptContract {
  name: string;
  status: CatalogStatus;
  description: string;
}

const STABLE_PROMPTS = new Set([
  "excel.prompts.create_next_month_sheet",
  "excel.prompts.clean_current_sheet",
  "excel.prompts.fix_formula_errors",
  "excel.prompts.format_like_template",
  "excel.prompts.field_value_image_to_styled_table",
  "excel.prompts.booking_image_to_styled_table",
  "excel.prompts.validate_report_before_saving",
  "excel.prompts.create_summary_report"
]);

const PROMPT_NAMES = [
  "excel.prompts.create_next_month_sheet",
  "excel.prompts.clean_current_sheet",
  "excel.prompts.fix_formula_errors",
  "excel.prompts.format_like_template",
  "excel.prompts.field_value_image_to_styled_table",
  "excel.prompts.booking_image_to_styled_table",
  "excel.prompts.validate_report_before_saving",
  "excel.prompts.create_summary_report"
] as const;

export const PromptCatalog: PromptContract[] = PROMPT_NAMES.map((name) => ({
  name,
  status: STABLE_PROMPTS.has(name) ? "stable" : "planned",
  description: `Open Workbook prompt ${name}.`
}));

export type PromptName = (typeof PROMPT_NAMES)[number];
