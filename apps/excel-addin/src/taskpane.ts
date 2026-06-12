import { AddinConnection } from "./connection.js";

const backendUrl = new URLSearchParams(window.location.search).get("backendUrl") ?? "ws://127.0.0.1:37845/addin";
document.getElementById("backend-url")?.replaceChildren(backendUrl);

const connection = new AddinConnection({
  backendUrl,
  heartbeatMs: 5_000
});

Office.onReady(() => {
  connection.connect();
});
