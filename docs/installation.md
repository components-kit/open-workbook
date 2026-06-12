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

Run the shared daemon:

```bash
node packages/cli/dist/index.js daemon start
```

Run the MCP adapter in another terminal:

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
owb daemon start
owb mcp
owb addin serve
owb opencode config --id open-workbook
```

The daemon starts the local backend WebSocket and owns shared workbook coordination. `owb mcp` attaches to that daemon when it is running. The add-in asset server is a separate process because Excel loads the taskpane HTML from it.

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

## HTTPS Add-in Serving

Loopback HTTP is the default for desktop sideload development. To serve the taskpane over HTTPS, provide a trusted local certificate and key:

```bash
owb addin serve --https \
  --tls-cert ./certs/open-workbook.local.pem \
  --tls-key ./certs/open-workbook.local-key.pem
```

Then generate a matching manifest:

```bash
OPEN_WORKBOOK_ADDIN_HTTPS=1 \
owb sideload manifest --out open-workbook.xml
```

The CLI also honors `OPEN_WORKBOOK_ADDIN_PROTOCOL=https`, `OPEN_WORKBOOK_ADDIN_TLS_CERT`, and `OPEN_WORKBOOK_ADDIN_TLS_KEY`. Certificate trust is intentionally left to the user or organization.

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
owb daemon start
owb mcp
owb addin serve
```

To generate an optional auto-start wrapper for the add-in asset server:

```bash
owb service manifest --target macos --service addin --out com.open-workbook.addin.plist
owb service manifest --target systemd --service addin --out com.open-workbook.addin.service
owb service manifest --target windows --service addin --out open-workbook-addin-task.ps1
```

See [Service Wrapper](service-wrapper.md) for install examples.

Then open Excel, load the sideloaded add-in, and call:

```text
excel.runtime.get_status
excel.runtime.get_active_context
excel.collab.get_status
```

from your MCP client.

## Troubleshooting

- Run `owb doctor` to confirm packaged assets are available.
- Run `owb paths` to see the daemon state directory and packaged entrypoints.
- Confirm the add-in server prints a manifest URL.
- Confirm `owb daemon status` returns JSON.
- Confirm the MCP process says it connected to the daemon, or use `owb mcp --standalone` for a single-process test.
- If Excel cannot load the add-in, regenerate the manifest and repeat sideloading.
- If Windows Excel does not show the add-in, confirm the trusted catalog is a shared-folder UNC path and `Show in Menu` is enabled.
- If the add-in loads but does not connect, confirm the `backendUrl` query string in the generated manifest points at the running MCP backend.
