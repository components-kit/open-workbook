# Service Wrapper

Open Workbook does not install login services automatically. The CLI can generate native wrapper files that users or administrators can review and install.

## Generate Wrappers

Add-in asset server:

```bash
owb service manifest --target macos --service addin --out com.open-workbook.addin.plist
owb service manifest --target systemd --service addin --out com.open-workbook.addin.service
owb service manifest --target windows --service addin --out open-workbook-addin-task.ps1
```

Shared daemon:

```bash
owb service manifest --target macos --service daemon --out com.open-workbook.daemon.plist
owb service manifest --target systemd --service daemon --out com.open-workbook.daemon.service
owb service manifest --target windows --service daemon --out open-workbook-daemon-task.ps1
```

Native file bridge:

```bash
owb service manifest --target macos --service file-bridge --out com.open-workbook.file-bridge.plist
owb service manifest --target systemd --service file-bridge --out com.open-workbook.file-bridge.service
owb service manifest --target windows --service file-bridge --out open-workbook-file-bridge-task.ps1
```

Use `--command` when the installed command is not `owb`.

## Install Examples

macOS launchd:

```bash
mkdir -p ~/Library/LaunchAgents
cp com.open-workbook.addin.plist ~/Library/LaunchAgents/
cp com.open-workbook.file-bridge.plist ~/Library/LaunchAgents/
launchctl load ~/Library/LaunchAgents/com.open-workbook.addin.plist
launchctl load ~/Library/LaunchAgents/com.open-workbook.file-bridge.plist
```

Linux systemd user service:

```bash
mkdir -p ~/.config/systemd/user
cp com.open-workbook.addin.service ~/.config/systemd/user/
cp com.open-workbook.file-bridge.service ~/.config/systemd/user/
systemctl --user enable --now com.open-workbook.addin.service
systemctl --user enable --now com.open-workbook.file-bridge.service
```

Windows Task Scheduler:

```powershell
powershell -ExecutionPolicy Bypass -File .\open-workbook-addin-task.ps1
powershell -ExecutionPolicy Bypass -File .\open-workbook-file-bridge-task.ps1
```

The wrapper only starts the local process. Excel still requires the user or administrator to trust and sideload the add-in manifest. Configure the daemon with `OPEN_WORKBOOK_FILE_BRIDGE_URL=http://127.0.0.1:37847` when using the file bridge for native Save As.
