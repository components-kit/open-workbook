# Scripts

Repository scripts are grouped by maintenance workflow. Public entrypoints stay in the root `package.json`; call those commands instead of invoking these files directly unless you are debugging a specific helper.

## Folders

- `build/`: build cleanup and executable bit fixes used by `pnpm build`.
- `validate/`: release-gate contract checks for the one-tool MCP surface, docs, skills, and package metadata.
- `release/`: npm package dry-run and publish helpers.
- `dev/`: local runtime helpers. Sideload commands are handled by the CLI.
- `docs/`: generated documentation helpers such as `llms-full.txt`.
- `reports/`: diagnostic reports that are useful during planning but are not release gates.
- `lib/`: small shared helpers for repo-maintenance scripts.

Test runners, smoke tests, E2E flows, fixtures, and benchmarks live under `tests/`.
