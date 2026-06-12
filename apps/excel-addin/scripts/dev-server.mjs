import { createServer } from "node:http";
import { createReadStream, existsSync } from "node:fs";
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
  const pathname = url.pathname === "/" ? "/taskpane.html" : url.pathname;
  const filePath =
    pathname === "/manifest.xml"
      ? join(root, "manifest.xml")
      : pathname.startsWith("/dist/")
        ? join(root, pathname)
        : join(publicDir, pathname);

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
  createReadStream(normalized).pipe(response);
});

server.listen(port, host, () => {
  console.log(`Open Workbook add-in server listening on http://${host}:${port}`);
  console.log(`Manifest: http://${host}:${port}/manifest.xml`);
});

function isAllowedPath(filePath) {
  return filePath.startsWith(publicDir) || filePath.startsWith(join(root, "dist")) || filePath === join(root, "manifest.xml");
}
