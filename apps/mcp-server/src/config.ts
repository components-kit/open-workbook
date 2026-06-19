export interface McpServerConfig {
  host: string;
  port: number;
  addinPath: string;
  daemonUrl: string;
  agentName?: string | undefined;
  standalone: boolean;
  runtimeVersion: string;
}

export function readConfig(): McpServerConfig {
  const host = process.env.OPEN_WORKBOOK_HOST ?? "127.0.0.1";
  const port = Number(process.env.OPEN_WORKBOOK_PORT ?? 37845);
  const addinPath = process.env.OPEN_WORKBOOK_ADDIN_PATH ?? "/addin";
  return {
    host,
    port,
    addinPath,
    daemonUrl: trimTrailingSlash(readArg("--daemon-url") ?? process.env.OPEN_WORKBOOK_DAEMON_URL ?? `http://${host}:${port}`),
    agentName: readArg("--agent-name") ?? process.env.OPEN_WORKBOOK_AGENT_NAME,
    standalone: hasArg("--standalone") || process.env.OPEN_WORKBOOK_MCP_STANDALONE === "1",
    runtimeVersion: process.env.OPEN_WORKBOOK_VERSION ?? "0.1.14"
  };
}

function hasArg(name: string): boolean {
  return process.argv.includes(name);
}

function readArg(name: string): string | undefined {
  const index = process.argv.indexOf(name);
  if (index >= 0 && process.argv[index + 1]) {
    return process.argv[index + 1];
  }
  const prefixed = `${name}=`;
  const value = process.argv.find((arg) => arg.startsWith(prefixed));
  return value ? value.slice(prefixed.length) : undefined;
}

function trimTrailingSlash(value: string): string {
  return value.replace(/\/+$/, "");
}
