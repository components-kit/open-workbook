# Packaging And Publishing

Open Workbook is designed for local-first, non-AppSource distribution. The public entry point is the `owb` CLI, which carries the built MCP server and Excel add-in assets.

## Packages

Publishable packages:

- `@open-workbook/cli`: user-facing CLI and bundled runtime assets
- `@open-workbook/mcp-server`: MCP stdio server
- `@open-workbook/backend`: local backend broker and runtime service
- `@open-workbook/protocol`: shared contracts, tool catalog, resources, and prompts
- `@open-workbook/excel-core`: planning, backups, snapshots, templates, permissions, and range utilities
- `@open-workbook/office-js-engine`: Office.js engine interface and defaults

Private workspace package:

- `@open-workbook/excel-addin`: source for the sideloaded add-in; bundled into `@open-workbook/cli` assets rather than published as a standalone install target

## Build

```bash
corepack pnpm install
corepack pnpm build
corepack pnpm test
```

The root build:

1. Compiles all TypeScript projects.
2. Copies built MCP server and add-in assets into `packages/cli/assets`.
3. Ensures CLI binaries are executable.

## CLI Asset Bundle

`@open-workbook/cli` resolves runtime assets in this order:

1. Source checkout paths, for contributors.
2. Packaged `packages/cli/assets`, for installed users.
3. Installed package dependencies, as a fallback for package-manager layouts.

This lets the same command shape work from source and from an installed package:

```bash
owb mcp
owb addin serve
owb sideload mac
owb sideload windows
owb sideload manifest
owb opencode config
owb doctor
owb paths
```

## Source Install Smoke Test

```bash
node packages/cli/dist/index.js doctor
node packages/cli/dist/index.js paths
node packages/cli/dist/index.js opencode config --id open-workbook --command "node packages/cli/dist/index.js"
node packages/cli/dist/index.js sideload manifest --out /tmp/open-workbook.xml
```

## Package Dry Run

```bash
npm pack --dry-run ./packages/cli
```

The tarball should include:

- `packages/cli/dist`
- `packages/cli/assets/mcp-server/dist`
- `packages/cli/assets/excel-addin/dist`
- `packages/cli/assets/excel-addin/public`
- `packages/cli/assets/excel-addin/scripts`
- `packages/cli/assets/excel-addin/manifest.xml`
- `packages/cli/README.md`

## Publish Checklist

Before publishing:

1. Confirm package versions are aligned.
2. Run `corepack pnpm verify`.
3. Run `node packages/cli/dist/index.js doctor`.
4. Run `corepack pnpm pack:dry-run`.
5. Confirm generated manifests include the expected taskpane URL and `backendUrl`.
6. Confirm `@open-workbook/excel-addin` remains private.
7. Confirm npm publish access is public for publishable scoped packages.

## End-User Install Shape

The intended user flow after package publishing:

```bash
npm install -g @open-workbook/cli
owb doctor
owb opencode config --id open-workbook
owb addin serve
owb sideload mac
```

Windows users should run:

```bash
owb sideload windows --out open-workbook.xml
```

and copy the generated manifest into a trusted shared-folder add-in catalog.

## Native Installer Shape

Native installers or administrator scripts can wrap the same commands:

- install Node/runtime assets or embed a Node runtime
- place `owb` on `PATH`
- register an optional auto-start process from `owb service manifest`
- guide or automate the Excel manifest trust step where the platform allows it

Excel still requires user or admin trust approval for the add-in manifest outside AppSource.
