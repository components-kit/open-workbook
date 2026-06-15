# Local Excel Add-in Sideloading

Open Workbook does not require Microsoft AppSource. Users can sideload the Excel add-in locally and run the MCP server on their machine.

## Runtime Model

- The MCP command starts the local backend automatically.
- The Excel add-in connects to that backend at `ws://127.0.0.1:37845/addin`.
- The add-in taskpane is served locally at `http://localhost:37846/taskpane.html`; `owb mcp` starts that local taskpane server by default.
- The source manifest is `apps/excel-addin/manifest.xml`; CLI sideload commands generate a runtime manifest with the active taskpane and backend URLs.

## Development Sideload

Run both processes:

```bash
corepack pnpm dev:mcp
corepack pnpm dev:addin
```

Then sideload a generated manifest into Excel. The source manifest is useful for development, but generated manifests include the active taskpane and backend URLs.

You can inspect the generated manifest with:

```bash
node packages/cli/dist/index.js sideload manifest --development
```

Development manifests use a separate Office add-in ID, display as `OpenWorkbook Local`, and write to `open-workbook-local.xml` on macOS. This keeps source sideloading from replacing the production `OpenWorkbook` manifest installed by `npx -y @components-kit/open-workbook@latest setup` or `upgrade`.

## macOS Manual Sideload

For Excel on macOS, copy the manifest to the Office add-ins sideload folder:

```bash
corepack pnpm sideload:mac
```

With the CLI directly:

```bash
owb sideload mac --development
```

Then restart Excel and open the add-in from the ribbon or Insert > Add-ins, depending on the Excel version.

## Windows Manual Sideload

For Windows desktop Excel, use a trusted shared-folder catalog:

1. Create a folder for add-in manifests, such as `C:\open-workbook-addins`.
2. Share that folder in Windows and note its UNC path, such as `\\YOUR-PC\open-workbook-addins`.
3. Generate the Open Workbook manifest and copy it into the shared folder.
4. In Excel, open Trust Center settings and add the UNC path as a trusted add-in catalog.
5. Select `Show in Menu`, restart Excel, and insert the shared-folder add-in.

The helper command writes a generated manifest and prints the expected catalog setup:

```bash
corepack pnpm sideload:windows
```

With the CLI directly:

```bash
owb sideload windows --development --out open-workbook-local.xml
```

Custom local ports can be passed to every manifest-producing command:

```bash
owb sideload manifest --addin-url http://localhost:37846 --backend-url ws://127.0.0.1:37845/addin
```

Omit `--development` when you intentionally want the production add-in identity.

For local desktop sideloading, the taskpane and backend use loopback HTTP and WebSocket endpoints. Ribbon and Developer Add-ins branding images are static PNG assets served from ComponentsKit over HTTPS so Excel can load branded icons without requiring local certificates.

## User Install Shape

For non-store open-source distribution, the expected install flow is:

1. Run `npx -y @components-kit/open-workbook setup`.
2. Add the printed MCP launch command to the agent UI's local stdio MCP configuration.
3. Install the skill with `npx skills add components-kit/open-workbook --skill open-workbook-excel`.
4. Open the agent UI, then Excel, then the Open Workbook add-in.

`owb service manifest` remains available for advanced deployments that want startup services. Excel still requires user or admin trust approval for the add-in manifest.
