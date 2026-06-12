# Contributing

Open Workbook is built around workbook safety. Contributions should preserve these invariants:

- Do not add write tools that bypass plans, snapshots, permissions, and diffs.
- Do not introduce per-cell Office.js loops for operations that can be batched.
- Do not treat styles, formulas, filters, tables, or print settings as optional for template workflows.
- Keep sensitive workbook data local by default.

## Development

```bash
corepack pnpm install
corepack pnpm check
corepack pnpm test
corepack pnpm build
node packages/cli/dist/index.js doctor
```

Run the local runtime from source:

```bash
node packages/cli/dist/index.js mcp
node packages/cli/dist/index.js addin serve
```

## Pull Request Expectations

- Include tests for protocol, planning, backup, template, or engine behavior touched by the change.
- Document new public tools or resources in `docs/tool-surface.md`.
- Include performance notes for new read/write paths.
- Update package README files when changing module responsibilities or public exports.
