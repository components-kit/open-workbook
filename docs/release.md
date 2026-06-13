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

4. Run package dry runs:

```bash
corepack pnpm pack:dry-run
```

5. Confirm the CLI reports healthy packaged assets:

```bash
node packages/cli/dist/index.js doctor
node packages/cli/dist/index.js setup --dry-run
node packages/cli/dist/index.js sideload manifest --out /tmp/open-workbook.xml
```

6. Confirm `setup --dry-run` prints the intended npm install shape:

```text
npx -y @components-kit/open-workbook@latest mcp
npx skills add components-kit/open-workbook --skill open-workbook-excel
```

7. Confirm generated manifests include the expected taskpane URL and backend URL.
8. Review `git diff` and commit the release prep changes.
9. Create an annotated tag:

```bash
git tag -a v0.1.0 -m "Release v0.1.0"
```

10. Push the release commit and tag when ready:

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
