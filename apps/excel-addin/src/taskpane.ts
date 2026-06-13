const backendUrl = new URLSearchParams(window.location.search).get("backendUrl") ?? "ws://127.0.0.1:37845/addin";
document.getElementById("backend-url")?.replaceChildren(backendUrl);
const status = document.getElementById("status");

status?.replaceChildren("Taskpane script loaded. Waiting for Office.js...");

if (!("Office" in window)) {
  status?.replaceChildren("Office.js did not load. Check network access to appsforoffice.microsoft.com.");
} else {
  Office.onReady(async () => {
    status?.replaceChildren("Office.js ready. Connecting to local runtime...");
    try {
      const { AddinConnection } = await import("./connection.js");
      const connection = new AddinConnection({
        backendUrl,
        heartbeatMs: 5_000,
        reconnectMs: 2_000,
        onStatus: (message) => status?.replaceChildren(message)
      });
      connection.connect();
    } catch (error) {
      const detail = error instanceof Error ? error.message : String(error);
      status?.replaceChildren(`Could not load taskpane runtime: ${detail}`);
    }
  });
}
