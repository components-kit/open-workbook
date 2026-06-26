import { OPEN_WORKBOOK_VERSION } from "@components-kit/open-workbook-protocol";
import { afterEach, describe, expect, it, vi } from "vitest";
import { getRuntimeCapabilities } from "./executor-core.js";

const previousOffice = (globalThis as any).Office;

afterEach(() => {
  (globalThis as any).Office = previousOffice;
});

describe("add-in runtime version", () => {
  it("uses the shared package version in runtime capabilities", () => {
    (globalThis as any).Office = {
      context: {
        diagnostics: { version: "16.0" },
        document: {},
        host: "Excel",
        platform: "Mac",
        requirements: {
          isSetSupported: vi.fn(() => true)
        }
      },
      PlatformType: {
        Mac: "Mac",
        PC: "PC"
      }
    };

    expect(getRuntimeCapabilities().engine.version).toBe(OPEN_WORKBOOK_VERSION);
    expect(getRuntimeCapabilities().engine.taskpaneBundleVersion).toBe("20260626-3");
  });
});
