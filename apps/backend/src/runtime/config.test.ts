import { OPEN_WORKBOOK_VERSION } from "@components-kit/open-workbook-protocol";
import { afterEach, describe, expect, it } from "vitest";
import { runtimeVersion } from "./config.js";

const previousVersion = process.env.OPEN_WORKBOOK_VERSION;

afterEach(() => {
  if (previousVersion === undefined) {
    delete process.env.OPEN_WORKBOOK_VERSION;
  } else {
    process.env.OPEN_WORKBOOK_VERSION = previousVersion;
  }
});

describe("runtime config version", () => {
  it("defaults to the shared package version", () => {
    delete process.env.OPEN_WORKBOOK_VERSION;

    expect(runtimeVersion()).toBe(OPEN_WORKBOOK_VERSION);
  });

  it("allows OPEN_WORKBOOK_VERSION override", () => {
    process.env.OPEN_WORKBOOK_VERSION = "9.8.7";

    expect(runtimeVersion()).toBe("9.8.7");
  });
});

