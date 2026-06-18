import { describe, expect, it } from "vitest";
import {
  NativeFileBridge,
  RuntimeService,
  appliedTransaction,
  existsSync,
  mkdirSync,
  mkdtempSync,
  operationOk,
  path,
  readFileSync,
  runtimeWithDynamicSnapshotHash,
  runtimeWithExecutingAddin,
  runtimeWithFormulaGraph,
  runtimeWithPersistentAddin,
  runtimeWithSnapshotHash,
  sleepForTest,
  snapshotResponse,
  tmpdir,
  writeFileSync,
  writeFormulaOperation,
  writeStyleOperation,
  writeValuesOperation
} from "./runtime-service.test-support.js";
import type { AgentId, OperationId, PlanId, RuntimeCapabilities, WorkbookId } from "./runtime-service.test-support.js";

describe("RuntimeService capabilities", () => {
  it("reports disconnected host capability fallback", () => {
    const runtime = new RuntimeService({ persistState: false });

    const capabilities = runtime.getCapabilities();

    expect(capabilities.activeHostCapabilities.engine.name).toBe("open-workbook-daemon");
    expect(capabilities.activeHostCapabilities.hostCapabilities?.some((capability) => capability.status === "unknown")).toBe(true);
    expect(capabilities.connectedHostCapabilities).toHaveLength(0);
    expect(capabilities.fileBridge.available).toBe(false);
    expect(runtime.getStatus().fileBridge.available).toBe(false);
  });

  it("reports configured native file bridge status", () => {
    const runtime = new RuntimeService({
      persistState: false,
      fileBridge: new NativeFileBridge({ url: "http://127.0.0.1:37999" })
    });

    expect(runtime.getStatus().fileBridge).toMatchObject({
      available: true,
      url: "http://127.0.0.1:37999",
      path: "/v1/workbook-file"
    });
    expect(runtime.getCapabilities().fileBridge.available).toBe(true);
  });

  it("marks old add-in sessions as stale instead of connected", () => {
    const runtime = new RuntimeService({ persistState: false });
    const session = runtime.sessions.createSession();
    session.lastSeenAt = new Date(Date.now() - 60_000).toISOString();

    const status = runtime.getStatus();

    expect(status.connectionState).toBe("stale");
    expect(status.activeAddinConnected).toBe(false);
    expect(status.activeAddinReachable).toBe(false);
    expect(status.sessions[0]).toMatchObject({ connectionId: session.connectionId, stale: true });
    expect(runtime.getCapabilities().connectedHostCapabilities).toHaveLength(0);
  });

  it("removes an active add-in session when the fast readiness probe fails", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async () => {
        throw new Error("Timed out waiting for add-in method: runtime.get_active_context");
      },
      close: () => undefined
    } as any);

    const readiness = await runtime.getConnectionReadiness();

    expect(readiness.ok).toBe(false);
    expect(readiness.connectionState).toBe("stale");
    expect(runtime.getStatus().connectionState).toBe("disconnected");
    expect(runtime.getStatus().sessions).toHaveLength(0);
  });

  it("reports ready after a fast readiness probe returns an active workbook", async () => {
    const runtime = new RuntimeService({ persistState: false });
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string) => {
        expect(method).toBe("runtime.get_active_context");
        return { workbookId: "workbook_ready", name: "Ready.xlsx", platform: "mac" };
      }
    } as any);

    const readiness = await runtime.getConnectionReadiness();

    expect(readiness.ok).toBe(true);
    expect(readiness.connectionState).toBe("ready");
    expect(runtime.getStatus()).toMatchObject({
      connectionState: "ready",
      activeAddinConnected: true,
      activeWorkbookAvailable: true
    });
  });

  it("can probe configured native file bridge status", async () => {
    const runtime = new RuntimeService({
      persistState: false,
      fileBridge: new NativeFileBridge({
        url: "http://127.0.0.1:37999",
        fetchImpl: (async () => Response.json({
          ok: true,
          bridge: "open-workbook-native-file-bridge",
          route: "/v1/workbook-file",
          adapter: { platform: "win32", saveAsSupported: true }
        })) as typeof fetch
      })
    });

    const status = await runtime.getStatusWithFileBridgeProbe();

    expect(status.fileBridge).toMatchObject({
      available: true,
      reachable: true,
      bridge: "open-workbook-native-file-bridge",
      route: "/v1/workbook-file",
      adapter: { platform: "win32", saveAsSupported: true }
    });
  });

  it("includes active add-in Office API set support when connected", () => {
    const runtime = new RuntimeService({ persistState: false });
    const session = runtime.sessions.createSession();
    const reportedCapabilities: RuntimeCapabilities = {
      engine: {
        name: "office-js-addin",
        version: "0.1.4",
        platform: "mac",
        host: "Excel",
        officeVersion: "16.99"
      },
      apiSets: [
        { set: "ExcelApi", version: "1.9", supported: true },
        { set: "ExcelApi", version: "1.17", supported: false }
      ],
      capabilities: [
        {
          name: "range.batch.read_write",
          supported: true,
          platforms: ["mac", "windows", "web"],
          requires: [{ set: "ExcelApi", version: "1.9" }]
        }
      ],
      hostCapabilities: [
        {
          name: "range-values-formulas-styles",
          supported: true,
          status: "supported",
          requires: [{ set: "ExcelApi", version: "1.9" }]
        }
      ]
    };
    runtime.sessions.update(session.connectionId, { capabilities: reportedCapabilities });

    const capabilities = runtime.getCapabilities();

    expect(capabilities.activeHostCapabilities.engine.platform).toBe("mac");
    expect(capabilities.activeHostCapabilities.apiSets?.find((apiSet) => apiSet.version === "1.9")?.supported).toBe(true);
    const connectedHost = capabilities.connectedHostCapabilities[0];
    expect(connectedHost?.connectionId).toBe(session.connectionId);
    expect(connectedHost?.capabilities?.hostCapabilities?.[0]?.status).toBe("supported");
  });
});

describe("RuntimeService selection", () => {
  it("passes through enriched selection metadata from the connected add-in", async () => {
    const workbookId = "workbook_runtime_selection" as WorkbookId;
    const runtime = new RuntimeService({ persistState: false });
    const session = runtime.sessions.createSession();
    runtime.attachAddinClient(session.connectionId, {
      request: async (method: string) => {
        expect(method).toBe("runtime.get_selection");
        return {
          workbook: {
            workbookId,
            name: "Selection.xlsx",
            platform: "mac"
          },
          selection: {
            workbookId,
            sheetName: "Sheet1",
            address: "B4:D10",
            startCell: {
              workbookId,
              sheetName: "Sheet1",
              address: "B4",
              row: 4,
              column: 2,
              rowIndex: 3,
              columnIndex: 1
            },
            endCell: {
              workbookId,
              sheetName: "Sheet1",
              address: "D10",
              row: 10,
              column: 4,
              rowIndex: 9,
              columnIndex: 3
            },
            rowCount: 7,
            columnCount: 3,
            cellCount: 21,
            isSingleCell: false
          }
        };
      }
    } as any);

    const result = await runtime.getSelection();

    expect(result.selection?.startCell).toMatchObject({ address: "B4", row: 4, column: 2 });
    expect(result.selection?.endCell).toMatchObject({ address: "D10", row: 10, column: 4 });
    expect(result.selection).toMatchObject({ rowCount: 7, columnCount: 3, cellCount: 21, isSingleCell: false });
  });
});
