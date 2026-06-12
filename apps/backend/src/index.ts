import { RuntimeService } from "./runtime-service.js";
import { startBackendServer } from "./addin-websocket-server.js";

const runtime = new RuntimeService();

const host = process.env.OPEN_WORKBOOK_HOST ?? "127.0.0.1";
const port = Number(process.env.OPEN_WORKBOOK_PORT ?? 37845);
const addinPath = process.env.OPEN_WORKBOOK_ADDIN_PATH ?? "/addin";

await startBackendServer(runtime, { host, port, addinPath });

console.log(`open-workbook backend listening on http://${host}:${port}`);

export { runtime };
