import { homedir } from "node:os";
import { join } from "node:path";

export type ServiceTarget = "macos" | "systemd" | "windows";
export type ServiceName = "addin" | "daemon" | "file-bridge";

export function generateServiceManifest(options: { target: ServiceTarget; serviceName: ServiceName; command: string }): string {
  const args = options.serviceName === "addin" ? ["addin", "serve"] : options.serviceName === "daemon" ? ["daemon", "start"] : ["file-bridge", "start"];
  const label = `com.open-workbook.${options.serviceName}`;
  const description =
    options.serviceName === "addin"
      ? "Open Workbook Excel add-in asset server"
      : options.serviceName === "daemon"
        ? "Open Workbook shared daemon"
        : "Open Workbook native file bridge";
  const commandParts = [options.command, ...args];
  switch (options.target) {
    case "macos":
      return generateLaunchdPlist(label, commandParts);
    case "systemd":
      return generateSystemdUnit(label, description, commandParts);
    case "windows":
      return generateWindowsScheduledTask(label, description, commandParts);
  }
}

export function defaultServiceTarget(): ServiceTarget {
  if (process.platform === "darwin") {
    return "macos";
  }
  if (process.platform === "win32") {
    return "windows";
  }
  return "systemd";
}

export function defaultServiceCommand(): string {
  return process.env.OPEN_WORKBOOK_SERVICE_COMMAND ?? "owb";
}

export function normalizeServiceTarget(value: string): ServiceTarget {
  if (value === "mac" || value === "macos" || value === "launchd") {
    return "macos";
  }
  if (value === "win" || value === "windows" || value === "task-scheduler") {
    return "windows";
  }
  if (value === "linux" || value === "systemd") {
    return "systemd";
  }
  throw new Error(`Unknown service target: ${value}`);
}

export function normalizeServiceName(value: string): ServiceName {
  if (value === "addin" || value === "daemon" || value === "file-bridge") {
    return value;
  }
  throw new Error(`Unknown service name: ${value}`);
}

function generateLaunchdPlist(label: string, commandParts: string[]): string {
  return `<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "https://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
  <key>Label</key>
  <string>${escapeXml(label)}</string>
  <key>ProgramArguments</key>
  <array>
${commandParts.map((part) => `    <string>${escapeXml(part)}</string>`).join("\n")}
  </array>
  <key>RunAtLoad</key>
  <true/>
  <key>KeepAlive</key>
  <true/>
  <key>StandardOutPath</key>
  <string>${escapeXml(join(homedir(), "Library/Logs", `${label}.out.log`))}</string>
  <key>StandardErrorPath</key>
  <string>${escapeXml(join(homedir(), "Library/Logs", `${label}.err.log`))}</string>
</dict>
</plist>
`;
}

function generateSystemdUnit(label: string, description: string, commandParts: string[]): string {
  return `[Unit]
Description=${description}
After=network.target

[Service]
Type=simple
ExecStart=${commandParts.map(shellQuote).join(" ")}
Restart=on-failure
RestartSec=5

[Install]
WantedBy=default.target

# Save as ~/.config/systemd/user/${label}.service
# Enable with: systemctl --user enable --now ${label}.service
`;
}

function generateWindowsScheduledTask(label: string, description: string, commandParts: string[]): string {
  const executable = commandParts[0]!;
  const argumentsText = commandParts.slice(1).join(" ");
  return `$Action = New-ScheduledTaskAction -Execute ${powerShellQuote(executable)} -Argument ${powerShellQuote(argumentsText)}
$Trigger = New-ScheduledTaskTrigger -AtLogOn
$Principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel LeastPrivilege
$Settings = New-ScheduledTaskSettingsSet -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
Register-ScheduledTask -TaskName ${powerShellQuote(label)} -Description ${powerShellQuote(description)} -Action $Action -Trigger $Trigger -Principal $Principal -Settings $Settings -Force
`;
}

function escapeXml(value: string): string {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll("\"", "&quot;")
    .replaceAll("'", "&apos;");
}

function shellQuote(value: string): string {
  return `'${value.replaceAll("'", "'\\''")}'`;
}

function powerShellQuote(value: string): string {
  return `'${value.replaceAll("'", "''")}'`;
}
