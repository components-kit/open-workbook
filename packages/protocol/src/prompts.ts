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
  "excel.prompts.validate_report_before_saving",
  "excel.prompts.create_summary_report"
]);

const UNSUPPORTED_PROMPTS = new Set([
  "excel.prompts.import_receipts_to_table",
  "excel.prompts.import_invoices_to_table",
  "excel.prompts.reconcile_statement",
  "excel.prompts.create_driver_payroll",
  "excel.prompts.import_fuel_slips",
  "excel.prompts.calculate_fuel_consumption",
  "excel.prompts.create_customer_transport_report",
  "excel.prompts.reconcile_job_payments"
]);

const PROMPT_NAMES = [
  "excel.prompts.create_next_month_sheet",
  "excel.prompts.clean_current_sheet",
  "excel.prompts.fix_formula_errors",
  "excel.prompts.import_receipts_to_table",
  "excel.prompts.import_invoices_to_table",
  "excel.prompts.reconcile_statement",
  "excel.prompts.format_like_template",
  "excel.prompts.validate_report_before_saving",
  "excel.prompts.create_summary_report",
  "excel.prompts.create_driver_payroll",
  "excel.prompts.import_fuel_slips",
  "excel.prompts.calculate_fuel_consumption",
  "excel.prompts.create_customer_transport_report",
  "excel.prompts.reconcile_job_payments"
] as const;

export const PromptCatalog: PromptContract[] = PROMPT_NAMES.map((name) => ({
  name,
  status: STABLE_PROMPTS.has(name) ? "stable" : UNSUPPORTED_PROMPTS.has(name) ? "unsupported" : "planned",
  description: `Open Workbook prompt ${name}.`
}));

export type PromptName = (typeof PROMPT_NAMES)[number];
