# Permissions and Cleaning

Permissions and cleaning are implemented together because bulk data cleanup must respect workbook safety rules.

## Permissions

`excel.permissions.*` manages runtime policy:

- `excel.permissions.get`: returns current policy, scope, and locked regions.
- `excel.permissions.set`: updates write, destructive-action, workbook-action, macro, confirmation, and scope settings.
- `excel.permissions.require_confirmation`: sets destructive levels that require a confirmation token.
- `excel.permissions.set_scope`: restricts mutations to a workbook, sheets, or registered regions.
- `excel.permissions.allow_destructive_actions`: allows or blocks structure/workbook destructive actions.
- `excel.permissions.allow_macro_execution`: records macro policy. Macro execution is not implemented.
- `excel.permissions.lock_regions`: blocks writes that overlap registered regions.
- `excel.permissions.unlock_regions`: removes region locks.

Policy checks run before an apply-mode batch reaches Excel. Direct table mutations also check write scope and locked regions.

Default runtime policy allows value/format writes, blocks structure/workbook destructive actions, blocks workbook actions, and has no confirmation requirement until configured.

## Cleaning

`excel.clean.*` reads values from Excel, transforms them in the backend, then writes through the standard batch path when mutation is needed.

Mutating cleaners:

- `excel.clean.normalize_headers`
- `excel.clean.trim_whitespace`
- `excel.clean.remove_duplicates`
- `excel.clean.parse_dates`
- `excel.clean.parse_numbers`
- `excel.clean.standardize_currency`
- `excel.clean.fill_missing_values`
- `excel.clean.split_column`
- `excel.clean.merge_columns`

Read-only cleaners:

- `excel.clean.detect_header_row`
- `excel.clean.detect_outliers`
- `excel.clean.fuzzy_match`

Because writes use `range.write_values`, cleaning preserves formatting and participates in backup, telemetry, permission, and rollback behavior.

## Current Limits

The permission state and locked region list are persisted in the daemon state file under `.open-workbook/state` by default, or under `OPEN_WORKBOOK_STATE_DIR` when configured.

Cleaning tools currently operate on values only. Formula-preserving table cleanup should use table-specific tools or template repair paths.
