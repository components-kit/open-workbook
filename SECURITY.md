# Security Policy

Open Workbook handles sensitive spreadsheet data. The default security posture is local-first and permission-gated.

## Data Handling

- Workbook snapshots, backups, diffs, and telemetry stay local by default.
- Integrations that send workbook content to external services must be explicit and documented.
- Logs must not include cell values unless a user enables diagnostic value logging.

## Reporting

Please report security issues privately before public disclosure. Use GitHub private vulnerability reporting or contact the maintainers through the repository security policy.

Include:

- affected version or commit
- steps to reproduce
- whether workbook content, local files, credentials, model-provider tokens, or MCP client configuration can be exposed or modified
- suggested mitigation, if known
