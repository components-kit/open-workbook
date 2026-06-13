import { describe, expect, it } from "vitest";
import type { WorkbookFileBridgeRequest, WorkbookId } from "@components-kit/open-workbook-protocol";
import { NativeFileBridge } from "./native-file-bridge.js";

describe("NativeFileBridge", () => {
  it("posts workbook operations to the configured bridge path", async () => {
    const urls: string[] = [];
    const request: WorkbookFileBridgeRequest = {
      operation: "workbook.export_copy",
      workbookId: "Book1.xlsx" as WorkbookId,
      targetPath: "/tmp/book-copy.xlsx"
    };
    const bridge = new NativeFileBridge({
      url: "http://127.0.0.1:37847/",
      path: "/custom/workbook-file",
      fetchImpl: (async (url) => {
        urls.push(String(url));
        return Response.json({
          ok: true,
          operation: request.operation,
          workbookId: request.workbookId,
          targetPath: request.targetPath,
          filePath: request.targetPath
        });
      }) as typeof fetch
    });

    expect(bridge.getStatus()).toMatchObject({
      available: true,
      url: "http://127.0.0.1:37847/",
      path: "/custom/workbook-file"
    });

    const result = await bridge.request(request);

    expect(result.ok).toBe(true);
    expect(urls).toEqual(["http://127.0.0.1:37847/custom/workbook-file"]);
  });

  it("probes configured bridge status", async () => {
    const urls: string[] = [];
    const bridge = new NativeFileBridge({
      url: "http://127.0.0.1:37847/",
      fetchImpl: (async (url) => {
        urls.push(String(url));
        return Response.json({
          ok: true,
          bridge: "open-workbook-native-file-bridge",
          route: "/v1/workbook-file",
          adapter: {
            platform: "darwin",
            operations: {
              "workbook.save_as": true,
              "workbook.export_copy": true
            }
          }
        });
      }) as typeof fetch
    });

    const status = await bridge.probeStatus();

    expect(status.reachable).toBe(true);
    expect(status.bridge).toBe("open-workbook-native-file-bridge");
    expect(status.route).toBe("/v1/workbook-file");
    expect(status.adapter?.platform).toBe("darwin");
    expect(urls).toEqual(["http://127.0.0.1:37847/status"]);
  });

  it("reports bridge probe failures without throwing", async () => {
    const bridge = new NativeFileBridge({
      url: "http://127.0.0.1:37847",
      fetchImpl: (async () => {
        throw new Error("connection refused");
      }) as typeof fetch
    });

    const status = await bridge.probeStatus();

    expect(status.available).toBe(true);
    expect(status.reachable).toBe(false);
    expect(status.error).toContain("connection refused");
  });
});
