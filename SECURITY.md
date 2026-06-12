# Security Policy

Open Workbook handles sensitive spreadsheet data. The default security posture is local-first and permission-gated.

## Data Handling

- Workbook snapshots, backups, diffs, and telemetry stay local by default.
- Integrations that send workbook content to external services must be explicit and documented.
- Logs must not include cell values unless a user enables diagnostic value logging.

## Reporting

Please report security issues privately to the maintainers before public disclosure.
