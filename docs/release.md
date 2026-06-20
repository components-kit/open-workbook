# Release Process

This document describes how to prepare Open Workbook releases. npm publishing is a separate step and can be done after a tagged release has been reviewed.

## Version Policy

Open Workbook publishes all public packages at the same version:

- `@components-kit/open-workbook`
- `@components-kit/open-workbook-mcp-server`
- `@components-kit/open-workbook-backend`
- `@components-kit/open-workbook-protocol`
- `@components-kit/open-workbook-excel-core`
- `@components-kit/open-workbook-office-js-engine`

The Excel add-in workspace package remains private and is bundled into the CLI package assets.

## Pre-Release Checklist

1. Confirm `package.json` and all workspace package versions match the intended release.
2. Update `CHANGELOG.md` with the release date and material user-facing changes.
3. Run the full verification suite:

```bash
corepack pnpm verify
```

4. Run the release E2E gate and review report-only quality diagnostics when making agent workflow claims:

```bash
corepack pnpm test:e2e
corepack pnpm test:e2e:report
corepack pnpm test:e2e:agent-surface
corepack pnpm test:e2e:agent-workflow
corepack pnpm test:e2e:office-agent:behavior
corepack pnpm test:e2e:agent:quality
corepack pnpm test:e2e:agent:quality:compare
corepack pnpm test:e2e:agent:quality:gate
```

5. Run package dry runs:

```bash
corepack pnpm pack:dry-run
```

6. Confirm the CLI reports healthy packaged assets:

```bash
node packages/cli/dist/index.js doctor
node packages/cli/dist/index.js setup --dry-run
node packages/cli/dist/index.js upgrade --dry-run
node packages/cli/dist/index.js sideload manifest --out /tmp/open-workbook.xml
```

7. Confirm `setup --dry-run` and `upgrade --dry-run` print the intended local stdio MCP launch command:

```text
npx -y @components-kit/open-workbook@latest mcp
npx skills add components-kit/open-workbook --skill open-workbook-skills
```

8. For host-readiness claims, open desktop Excel with the sideloaded add-in connected and run the opt-in live smoke:

```bash
OPEN_WORKBOOK_LIVE_E2E=1 corepack pnpm test:e2e:live:mac
OPEN_WORKBOOK_LIVE_E2E=1 corepack pnpm test:e2e:live:windows
OPEN_WORKBOOK_LIVE_E2E=1 corepack pnpm test:e2e:live:mac -- --deep
```

9. Confirm generated manifests include the expected taskpane URL and backend URL.
10. Review `git diff` and commit the release prep changes.
11. Create an annotated tag:

```bash
git tag -a v0.1.0 -m "Release v0.1.0"
```

12. Push the release commit and tag when ready:

```bash
git push origin main
git push origin v0.1.0
```

## npm Publishing Later

Publish from a clean checkout after the tag has been pushed and reviewed. Use pnpm so workspace dependencies are rewritten to package versions in packed artifacts.

```bash
corepack pnpm install --frozen-lockfile
corepack pnpm verify
corepack pnpm pack:dry-run
corepack pnpm publish:npm
```

Do not publish `@components-kit/open-workbook-excel-addin`; it is private by design.

If npm requires a one-time password, forward it to every publish command:

```bash
corepack pnpm publish:npm -- --otp 123456
```
