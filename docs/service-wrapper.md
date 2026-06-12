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

Use `--command` when the installed command is not `owb`.

## Install Examples

macOS launchd:

```bash
mkdir -p ~/Library/LaunchAgents
cp com.open-workbook.addin.plist ~/Library/LaunchAgents/
launchctl load ~/Library/LaunchAgents/com.open-workbook.addin.plist
```

Linux systemd user service:

```bash
mkdir -p ~/.config/systemd/user
cp com.open-workbook.addin.service ~/.config/systemd/user/
systemctl --user enable --now com.open-workbook.addin.service
```

Windows Task Scheduler:

```powershell
powershell -ExecutionPolicy Bypass -File .\open-workbook-addin-task.ps1
```

The wrapper only starts the local process. Excel still requires the user or administrator to trust and sideload the add-in manifest.
