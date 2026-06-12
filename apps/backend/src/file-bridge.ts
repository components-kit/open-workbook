import { startNativeFileBridgeServer } from "./native-file-bridge-server.js";

const host = process.env.OPEN_WORKBOOK_FILE_BRIDGE_HOST ?? "127.0.0.1";
const port = Number(process.env.OPEN_WORKBOOK_FILE_BRIDGE_PORT ?? 37847);
const route = process.env.OPEN_WORKBOOK_FILE_BRIDGE_PATH ?? "/v1/workbook-file";

try {
  await startNativeFileBridgeServer({ host, port, path: route });
  console.log(`open-workbook file bridge listening on http://${host}:${port}`);
  console.log(`Set OPEN_WORKBOOK_FILE_BRIDGE_URL=http://${host}:${port} for the backend daemon.`);
} catch (error) {
  const detail = error instanceof Error ? error.message : String(error);
  console.error(`Failed to start open-workbook file bridge on http://${host}:${port}: ${detail}`);
  process.exit(1);
}
