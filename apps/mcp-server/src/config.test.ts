import { OPEN_WORKBOOK_VERSION } from "@components-kit/open-workbook-protocol";
import { afterEach, describe, expect, it } from "vitest";
import { readConfig } from "./config.js";

const previousArgv = process.argv;
const previousEnv = {
  OPEN_WORKBOOK_VERSION: process.env.OPEN_WORKBOOK_VERSION,
  OPEN_WORKBOOK_DAEMON_URL: process.env.OPEN_WORKBOOK_DAEMON_URL,
  OPEN_WORKBOOK_AGENT_NAME: process.env.OPEN_WORKBOOK_AGENT_NAME
};

afterEach(() => {
  process.argv = previousArgv;
  restoreEnv("OPEN_WORKBOOK_VERSION", previousEnv.OPEN_WORKBOOK_VERSION);
  restoreEnv("OPEN_WORKBOOK_DAEMON_URL", previousEnv.OPEN_WORKBOOK_DAEMON_URL);
  restoreEnv("OPEN_WORKBOOK_AGENT_NAME", previousEnv.OPEN_WORKBOOK_AGENT_NAME);
});

describe("MCP config version", () => {
  it("defaults to the shared package version", () => {
    process.argv = ["node", "mcp-server"];
    delete process.env.OPEN_WORKBOOK_VERSION;
    delete process.env.OPEN_WORKBOOK_DAEMON_URL;

    expect(readConfig().runtimeVersion).toBe(OPEN_WORKBOOK_VERSION);
  });

  it("allows OPEN_WORKBOOK_VERSION override", () => {
    process.argv = ["node", "mcp-server"];
    process.env.OPEN_WORKBOOK_VERSION = "9.8.7";

    expect(readConfig().runtimeVersion).toBe("9.8.7");
  });
});

function restoreEnv(name: keyof typeof process.env, value: string | undefined): void {
  if (value === undefined) {
    delete process.env[name];
  } else {
    process.env[name] = value;
  }
}

