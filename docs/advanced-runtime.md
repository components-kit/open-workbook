# Advanced Runtime

The default technical-user flow is:

```bash
npx -y @components-kit/open-workbook setup
```

and an MCP client command:

```bash
npx -y @components-kit/open-workbook@latest mcp
```

That command starts the MCP adapter, starts the local Excel add-in taskpane server when needed, and uses an embedded backend when no shared daemon is running.

## Shared Daemon

Use the daemon when multiple MCP clients or long-lived background coordination should share one workbook runtime:

```bash
owb daemon start
owb mcp
```

`owb mcp` attaches to the daemon when it is available. If no daemon is available, it falls back to embedded backend mode.

## Add-in Asset Server

The Excel manifest points at the local taskpane URL by default:

```text
http://localhost:37846/taskpane.html
```

`owb mcp` starts that server automatically for the simple flow. You can also run it directly:

```bash
owb addin serve
```

The taskpane server stays on local HTTP. The manifest uses hosted ComponentsKit HTTPS URLs only for static branding icons that Excel displays in the ribbon and Developer Add-ins gallery.

## File Bridge

The native file bridge is optional and used for host-level Save As and file-copy workflows:

```bash
owb file-bridge start
owb file-bridge smoke --workbook Book1.xlsx --target ./book-copy.xlsx
```

## Service Wrappers

Open Workbook can generate service wrapper files for teams that want login-time background processes:

```bash
owb service manifest --target macos --service daemon --out com.open-workbook.daemon.plist
owb service manifest --target macos --service addin --out com.open-workbook.addin.plist
owb service manifest --target windows --service daemon --out open-workbook-daemon-task.ps1
```

See [Service Wrapper](service-wrapper.md) for install examples.

## Custom URLs

```bash
OPEN_WORKBOOK_PORT=37855 \
OPEN_WORKBOOK_ADDIN_PORT=37856 \
owb sideload manifest --out open-workbook.xml
```

You can also pass URLs directly:

```bash
owb sideload manifest \
  --addin-url http://127.0.0.1:37846 \
  --backend-url ws://127.0.0.1:37845/addin \
  --out open-workbook.xml
```
