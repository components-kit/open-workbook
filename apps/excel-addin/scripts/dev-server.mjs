import { createServer } from "node:http";
import { existsSync, readFileSync } from "node:fs";
import { extname, join, normalize } from "node:path";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("..", import.meta.url));
const publicDir = join(root, "public");
const port = Number(process.env.OPEN_WORKBOOK_ADDIN_PORT ?? 37846);
const host = process.env.OPEN_WORKBOOK_ADDIN_HOST ?? "127.0.0.1";

const contentTypes = new Map([
  [".html", "text/html; charset=utf-8"],
  [".css", "text/css; charset=utf-8"],
  [".js", "text/javascript; charset=utf-8"],
  [".png", "image/png"],
  [".xml", "application/xml; charset=utf-8"]
]);

const server = createServer((request, response) => {
  const url = new URL(request.url ?? "/", `http://${host}:${port}`);
  if (url.pathname.startsWith("/assets/icon-") && url.pathname.endsWith(".png")) {
    response.writeHead(200, {
      "content-type": "image/png",
      "cache-control": "no-store"
    });
    response.end(Buffer.from(ICON_PNG_BASE64, "base64"));
    return;
  }
  if (url.pathname === "/manifest.xml") {
    response.writeHead(200, {
      "content-type": "application/xml; charset=utf-8",
      "cache-control": "no-store"
    });
    response.end(generateManifest());
    return;
  }

  const pathname = url.pathname === "/" ? "/taskpane.html" : url.pathname;
  const filePath = pathname.startsWith("/dist/") ? join(root, pathname) : join(publicDir, pathname);

  const normalized = normalize(filePath);
  if (!isAllowedPath(normalized) || !existsSync(normalized)) {
    response.writeHead(404, { "content-type": "application/json" });
    response.end(JSON.stringify({ error: "not_found" }));
    return;
  }

  response.writeHead(200, {
    "content-type": contentTypes.get(extname(normalized)) ?? "application/octet-stream",
    "cache-control": "no-store"
  });
  response.end(readFileSync(normalized));
});

server.listen(port, host, () => {
  console.log(`Open Workbook add-in server listening on http://${host}:${port}`);
  console.log(`Manifest: http://${host}:${port}/manifest.xml`);
});

function isAllowedPath(filePath) {
  return filePath.startsWith(publicDir) || filePath.startsWith(join(root, "dist"));
}

function generateManifest() {
  const addinUrl = trimTrailingSlash(process.env.OPEN_WORKBOOK_ADDIN_URL ?? `http://${host}:${port}`);
  const backendUrl = process.env.OPEN_WORKBOOK_BACKEND_URL ?? `ws://${process.env.OPEN_WORKBOOK_HOST ?? "127.0.0.1"}:${process.env.OPEN_WORKBOOK_PORT ?? 37845}${process.env.OPEN_WORKBOOK_ADDIN_PATH ?? "/addin"}`;
  const taskpaneUrl = `${addinUrl}/taskpane.html?backendUrl=${encodeURIComponent(backendUrl)}`;
  return readFileSync(join(root, "manifest.xml"), "utf8")
    .replaceAll("http://localhost:37846/taskpane.html", taskpaneUrl)
    .replaceAll("http://localhost:37846", addinUrl);
}

function trimTrailingSlash(value) {
  return value.endsWith("/") ? value.slice(0, -1) : value;
}

const ICON_PNG_BASE64 =
  "iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAIAAAAlC+aJAAAAFUlEQVR4nO3BAQEAAACCIP+vbkhAAQAAAO8GEABAAAGl9n6SAAAAAElFTkSuQmCC";
