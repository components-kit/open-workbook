# Tool Surface

The long-term namespace is broad, but implementation should grow from reliable primitives.

## First Tools

- `excel.runtime.get_status`
- `excel.runtime.get_active_context`
- `excel.workbook.list_open_workbooks`
- `excel.workbook.create_backup`
- `excel.workbook.restore_backup`
- `excel.workbook.snapshot`
- `excel.range.read_full`
- `excel.batch.validate`
- `excel.batch.dry_run`
- `excel.batch.apply`
- `excel.plan.create`
- `excel.plan.preview`
- `excel.plan.apply`
- `excel.plan.rollback`
- `excel.diff.summarize`
- `excel.template.register`
- `excel.template.create_sheet_from_template`
- `excel.template.validate_sheet_against_template`
- `excel.style.validate_consistency`
- `excel.formula.validate_against_template`

## Resources

- `excel://runtime/status`
- `excel://workbooks`
- `excel://workbooks/{workbook_id}/map`
- `excel://workbooks/{workbook_id}/sheets`
- `excel://workbooks/{workbook_id}/templates`
- `excel://workbooks/{workbook_id}/snapshots/{snapshot_id}`
- `excel://workbooks/{workbook_id}/plans/{plan_id}/diff`

## Rule

No mutating tool should directly write to Excel. It must create or apply a plan through the backup and batch lifecycle.

## Implemented MCP Tools

The current stdio MCP server exposes:

- `excel.runtime.get_status`
- `excel.runtime.get_active_context`
- `excel.workbook.list_open_workbooks`
- `excel.range.read_full`
- `excel.batch.validate`
- `excel.batch.dry_run`
- `excel.batch.apply`
- `excel.plan.create`
- `excel.plan.preview`
- `excel.plan.apply`

`excel.batch.apply` and `excel.plan.apply` route through backend snapshots, backup records, target-region conflict checks, and add-in Office.js execution.
