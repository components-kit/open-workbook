import { describe, expect, it } from "vitest";
import type { WorkbookFileBridgeRequest, WorkbookId } from "@open-workbook/protocol";
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
});
