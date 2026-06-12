import { AddinConnection } from "./connection.js";

const connection = new AddinConnection({
  backendUrl: "ws://127.0.0.1:37845/addin",
  heartbeatMs: 5_000
});

Office.onReady(() => {
  connection.connect();
});
