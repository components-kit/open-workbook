# Installation

Open Workbook is distributed as a local MCP runtime plus a sideloaded Excel Office.js add-in. It does not require Microsoft AppSource.

## Requirements

- Node.js `>=20.11`
- Corepack and pnpm
- Desktop Excel on macOS or Windows
- An MCP client such as OpenCode
- Permission to sideload or trust an Office add-in manifest

## Source Install

```bash
git clone https://github.com/open-workbook/open-workbook.git
cd open-workbook
corepack pnpm install
corepack pnpm build
node packages/cli/dist/index.js doctor
```

Run the MCP server:

```bash
node packages/cli/dist/index.js mcp
```

Run the add-in asset server:

```bash
node packages/cli/dist/index.js addin serve
```

Generate an OpenCode MCP config snippet:

```bash
node packages/cli/dist/index.js opencode config --id open-workbook --command "node packages/cli/dist/index.js"
```

## Installed CLI Shape

After package installation, use `owb` directly:

```bash
owb doctor
owb mcp
owb addin serve
owb opencode config --id open-workbook
```

The MCP server starts the local backend WebSocket automatically. The add-in asset server is a separate process because Excel loads the taskpane HTML from it.

## Sideload On macOS

```bash
owb sideload mac
```

This writes a generated manifest to:

```text
~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/open-workbook.xml
```

Restart Excel, then open the add-in from the ribbon or Insert > Add-ins, depending on your Excel version.

## Sideload On Windows

```bash
owb sideload windows --out open-workbook.xml
```

Then:

1. Create a folder such as `C:\open-workbook-addins`.
2. Share the folder in Windows.
3. Copy `open-workbook.xml` into the shared folder.
4. In Excel, open `File > Options > Trust Center > Trust Center Settings > Trusted Add-in Catalogs`.
5. Add the shared folder UNC path, such as `\\YOUR-PC\open-workbook-addins`.
6. Select `Show in Menu`.
7. Restart Excel and insert Open Workbook from Shared Folder.

## Custom Ports

```bash
OPEN_WORKBOOK_PORT=37855 \
OPEN_WORKBOOK_ADDIN_PORT=37856 \
owb sideload manifest --out open-workbook.xml
```

You can also pass URLs explicitly:

```bash
owb sideload manifest \
  --addin-url http://127.0.0.1:37846 \
  --backend-url ws://127.0.0.1:37845/addin \
  --out open-workbook.xml
```

## Runtime Processes

Use these together during local testing:

```bash
owb mcp
owb addin serve
```

Then open Excel, load the sideloaded add-in, and call:

```text
excel.runtime.get_status
excel.runtime.get_active_context
```

from your MCP client.

## Troubleshooting

- Run `owb doctor` to confirm packaged assets are available.
- Confirm the add-in server prints a manifest URL.
- Confirm the MCP process prints the backend WebSocket URL.
- If Excel cannot load the add-in, regenerate the manifest and repeat sideloading.
- If Windows Excel does not show the add-in, confirm the trusted catalog is a shared-folder UNC path and `Show in Menu` is enabled.
- If the add-in loads but does not connect, confirm the `backendUrl` query string in the generated manifest points at the running MCP backend.
