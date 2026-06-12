export interface ToolContract {
  name: string;
  title: string;
  mutatesWorkbook: boolean;
  requiresConfirmation: boolean;
  description: string;
}

export const InitialToolContracts: ToolContract[] = [
  {
    name: "excel.runtime.get_status",
    title: "Get runtime status",
    mutatesWorkbook: false,
    requiresConfirmation: false,
    description: "Return backend, MCP, and add-in connection status."
  },
  {
    name: "excel.runtime.get_active_context",
    title: "Get active Excel context",
    mutatesWorkbook: false,
    requiresConfirmation: false,
    description: "Return active workbook, sheet, selected range, and capability summary."
  },
  {
    name: "excel.workbook.create_backup",
    title: "Create workbook backup",
    mutatesWorkbook: false,
    requiresConfirmation: false,
    description: "Create a local workbook-copy backup before risky operations."
  },
  {
    name: "excel.range.read_full",
    title: "Read complete range state",
    mutatesWorkbook: false,
    requiresConfirmation: false,
    description: "Read values, formulas, display text, formats, validation, comments, and metadata for a range."
  },
  {
    name: "excel.batch.validate",
    title: "Validate batch",
    mutatesWorkbook: false,
    requiresConfirmation: false,
    description: "Validate a proposed batch without previewing or applying."
  },
  {
    name: "excel.batch.dry_run",
    title: "Dry-run batch",
    mutatesWorkbook: false,
    requiresConfirmation: false,
    description: "Compile a batch and return estimated effects, warnings, backups, and diffs."
  },
  {
    name: "excel.batch.apply",
    title: "Apply batch",
    mutatesWorkbook: true,
    requiresConfirmation: true,
    description: "Apply a validated batch through the backup and diff lifecycle."
  },
  {
    name: "excel.plan.create",
    title: "Create workbook plan",
    mutatesWorkbook: false,
    requiresConfirmation: false,
    description: "Create a reversible plan from operations."
  },
  {
    name: "excel.plan.apply",
    title: "Apply workbook plan",
    mutatesWorkbook: true,
    requiresConfirmation: true,
    description: "Apply a previewed plan if target-region fingerprints still match."
  },
  {
    name: "excel.plan.rollback",
    title: "Rollback workbook plan",
    mutatesWorkbook: true,
    requiresConfirmation: true,
    description: "Rollback an applied plan using operation snapshots or workbook-copy backups."
  },
  {
    name: "excel.template.register",
    title: "Register template",
    mutatesWorkbook: false,
    requiresConfirmation: false,
    description: "Capture a reusable template fingerprint from a workbook sheet or region."
  },
  {
    name: "excel.template.create_sheet_from_template",
    title: "Create sheet from template",
    mutatesWorkbook: true,
    requiresConfirmation: true,
    description: "Create a sheet from a registered template while preserving structure, styles, formulas, and filters."
  }
];
