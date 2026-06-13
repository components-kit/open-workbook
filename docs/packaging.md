# Packaging And Publishing

Open Workbook is designed for local-first, non-AppSource distribution. The public entry point is `@component-kit/open-workbook`, which can be run through `npx` and carries the built MCP server, Excel add-in assets, and fallback agent instructions. The primary skill install path is skills.sh against `skills/open-workbook-excel`.

## Packages

Publishable packages:

- `@component-kit/open-workbook`: user-facing CLI and bundled runtime assets
- `@component-kit/open-workbook-mcp-server`: MCP stdio server
- `@component-kit/open-workbook-backend`: local backend broker and runtime service
- `@component-kit/open-workbook-protocol`: shared contracts, tool catalog, resources, and prompts
- `@component-kit/open-workbook-excel-core`: planning, backups, snapshots, templates, permissions, and range utilities
- `@component-kit/open-workbook-office-js-engine`: Office.js engine interface and defaults

Private workspace package:

- `@component-kit/open-workbook-excel-addin`: source for the sideloaded add-in; bundled into `@component-kit/open-workbook` assets rather than published as a standalone install target

## Build

```bash
corepack pnpm install
corepack pnpm build
corepack pnpm test
```

The root build:

1. Compiles all TypeScript projects.
2. Copies built MCP server, add-in assets, and fallback instructions into `packages/cli/assets`.
3. Ensures CLI binaries are executable.

## CLI Asset Bundle

`@component-kit/open-workbook` resolves runtime assets in this order:

1. Source checkout paths, for contributors.
2. Packaged `packages/cli/assets`, for installed users.
3. Installed package dependencies, as a fallback for package-manager layouts.

This lets the same command shape work from source and from an installed package:

```bash
owb mcp
owb addin serve
owb setup
owb instructions
owb sideload mac
owb sideload windows
owb sideload manifest
owb doctor
owb paths
```

## Source Install Smoke Test

```bash
node packages/cli/dist/index.js doctor
node packages/cli/dist/index.js paths
node packages/cli/dist/index.js setup --dry-run
node packages/cli/dist/index.js instructions
node packages/cli/dist/index.js sideload manifest --out /tmp/open-workbook.xml
```

## Package Dry Run

```bash
pnpm pack --dry-run ./packages/cli
```

The tarball should include:

- `packages/cli/dist`
- `packages/cli/assets/mcp-server/dist`
- `packages/cli/assets/excel-addin/dist`
- `packages/cli/assets/excel-addin/public`
- `packages/cli/assets/excel-addin/scripts`
- `packages/cli/assets/excel-addin/manifest.xml`
- `packages/cli/assets/instructions/open-workbook-excel/SKILL.md`
- `packages/cli/assets/instructions/open-workbook-excel/references`
- `packages/cli/README.md`

## Publish Checklist

Before publishing:

1. Confirm package versions are aligned.
2. Run `corepack pnpm verify`.
3. Run `node packages/cli/dist/index.js doctor`.
4. Run `corepack pnpm pack:dry-run`.
5. Confirm `node packages/cli/dist/index.js setup --dry-run` prints the generic `npx @component-kit/open-workbook@latest mcp` config and `npx skills add components-kit/open-workbook --skill open-workbook-excel`.
6. Confirm generated manifests include the expected taskpane URL and `backendUrl`.
7. Confirm `@component-kit/open-workbook-excel-addin` remains private.
8. Confirm npm publish access is public for publishable scoped packages.

`corepack pnpm verify` runs `scripts/validate-package-metadata.mjs`, which enforces the package version, repository, public/private publish intent, `dist` entrypoints, README presence, and publish access rules above. Use pnpm for packing/publishing so workspace dependencies are rewritten to publishable semver ranges.

## End-User Install Shape

The intended user flow after package publishing:

```bash
npx -y @component-kit/open-workbook setup
```

Users paste the printed MCP config into their agent UI:

```json
{
  "mcpServers": {
    "open-workbook": {
      "command": "npx",
      "args": ["-y", "@component-kit/open-workbook@latest", "mcp"]
    }
  }
}
```

They install the skill with:

```bash
npx skills add components-kit/open-workbook --skill open-workbook-excel
```

For OpenCode:

```bash
npx skills add components-kit/open-workbook --skill open-workbook-excel -a opencode -g -y
```

## Native Installer Shape

Native installers or administrator scripts can wrap the same commands:

- install Node/runtime assets or embed a Node runtime
- place `owb` on `PATH` or configure an equivalent `npx` command
- register an optional auto-start process from `owb service manifest`
- guide or automate the Excel manifest trust step where the platform allows it

Excel still requires user or admin trust approval for the add-in manifest outside AppSource.
