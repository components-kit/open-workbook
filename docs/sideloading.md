# Local Excel Add-in Sideloading

Open Workbook does not require Microsoft AppSource. Users can sideload the Excel add-in locally and run the MCP server on their machine.

## Runtime Model

- The MCP server starts the local backend automatically.
- The Excel add-in connects to that backend at `ws://127.0.0.1:37845/addin`.
- The add-in taskpane is served locally at `http://localhost:37846/taskpane.html`.
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
node packages/cli/dist/index.js sideload manifest
```

## macOS Manual Sideload

For Excel on macOS, copy the manifest to the Office add-ins sideload folder:

```bash
corepack pnpm sideload:mac
```

With the CLI directly:

```bash
owb sideload mac
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
owb sideload windows --out open-workbook.xml
```

Custom local ports can be passed to every manifest-producing command:

```bash
owb sideload manifest --addin-url http://127.0.0.1:37846 --backend-url ws://127.0.0.1:37845/addin
```

For local desktop sideloading, the default URLs use loopback HTTP and WebSocket endpoints. HTTPS serving is available when you provide a trusted local certificate:

```bash
owb addin serve --https \
  --tls-cert ./certs/open-workbook.local.pem \
  --tls-key ./certs/open-workbook.local-key.pem

OPEN_WORKBOOK_ADDIN_HTTPS=1 \
owb sideload manifest --out open-workbook.xml
```

The CLI does not install or trust certificates automatically. Use a local tool such as `mkcert` or your organization's certificate process, then pass the certificate paths to the add-in server.

## User Install Shape

For non-store open-source distribution, the expected install flow is:

1. Install/configure the MCP server in the MCP client.
2. Run or auto-start the local add-in asset server.
3. Sideload the manifest once.
4. Open Excel; the add-in connects to the MCP-started backend.

`owb service manifest` can generate launchd, systemd user, or Windows scheduled-task wrappers for auto-starting the add-in asset server. Excel still requires user or admin trust approval for the add-in manifest.
