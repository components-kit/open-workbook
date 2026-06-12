# Local Excel Add-in Sideloading

Open Workbook does not require Microsoft AppSource. Users can sideload the Excel add-in locally and run the MCP server on their machine.

## Runtime Model

- The MCP server starts the local backend automatically.
- The Excel add-in connects to that backend at `ws://127.0.0.1:37845/addin`.
- The add-in taskpane is served locally at `http://127.0.0.1:37846/taskpane.html`.
- The sideload manifest is `apps/excel-addin/manifest.xml`.

## Development Sideload

Run both processes:

```bash
corepack pnpm dev:mcp
corepack pnpm dev:addin
```

Then sideload `apps/excel-addin/manifest.xml` into Excel.

## macOS Manual Sideload

For Excel on macOS, copy the manifest to the Office add-ins sideload folder:

```bash
corepack pnpm sideload:mac
```

Then restart Excel and open the add-in from the ribbon or Insert > Add-ins, depending on the Excel version.

## Windows Manual Sideload

For Windows desktop Excel, use a trusted catalog share:

1. Create a folder for add-in manifests.
2. Copy `apps/excel-addin/manifest.xml` into that folder.
3. In Excel, open Trust Center settings and add that folder as a trusted add-in catalog.
4. Restart Excel and insert the shared-folder add-in.

The helper command prints the source manifest path and expected catalog setup:

```bash
corepack pnpm sideload:windows
```

## User Install Shape

For non-store open-source distribution, the expected install flow is:

1. Install/configure the MCP server in the MCP client.
2. Run or auto-start the local add-in asset server.
3. Sideload the manifest once.
4. Open Excel; the add-in connects to the MCP-started backend.

Later releases can package these steps into a native installer, but Excel still requires user or admin trust approval for the add-in manifest.
