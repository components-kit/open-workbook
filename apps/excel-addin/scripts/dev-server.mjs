import { createServer as createHttpServer } from "node:http";
import { existsSync, readFileSync } from "node:fs";
import { extname, join, normalize } from "node:path";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("..", import.meta.url));
const repoRoot = join(root, "../..");
const packageScopeRoot = join(repoRoot, "..");
const publicDir = join(root, "public");
const workspaceModuleDirs = new Map([
  ["/workspace/excel-core/", firstExistingPath([
    join(repoRoot, "packages/excel-core/dist"),
    join(repoRoot, "../excel-core/dist"),
    join(repoRoot, "node_modules/@components-kit/open-workbook-excel-core/dist"),
    join(packageScopeRoot, "open-workbook-excel-core/dist")
  ])],
  ["/workspace/protocol/", firstExistingPath([
    join(repoRoot, "packages/protocol/dist"),
    join(repoRoot, "../protocol/dist"),
    join(repoRoot, "node_modules/@components-kit/open-workbook-protocol/dist"),
    join(packageScopeRoot, "open-workbook-protocol/dist")
  ])]
]);
const port = Number(process.env.OPEN_WORKBOOK_ADDIN_PORT ?? 37846);
const host = process.env.OPEN_WORKBOOK_ADDIN_HOST ?? "127.0.0.1";
const protocol = "http";
const runtimeVersion = process.env.OPEN_WORKBOOK_VERSION ?? readPackageVersion() ?? "0.1.1";

const contentTypes = new Map([
  [".html", "text/html; charset=utf-8"],
  [".css", "text/css; charset=utf-8"],
  [".js", "text/javascript; charset=utf-8"],
  [".png", "image/png"],
  [".svg", "image/svg+xml; charset=utf-8"],
  [".xml", "application/xml; charset=utf-8"]
]);

const requestHandler = (request, response) => {
  const url = new URL(request.url ?? "/", `${protocol}://${host}:${port}`);
  if (request.method === "GET" && url.pathname === "/status") {
    response.writeHead(200, {
      "content-type": "application/json",
      "cache-control": "no-store"
    });
    response.end(JSON.stringify(getStatus()));
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
  const workspaceModule = resolveWorkspaceModule(pathname);
  const filePath = workspaceModule ?? (pathname.startsWith("/dist/") ? join(root, pathname) : join(publicDir, pathname));

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
  const body = readFileSync(normalized);
  response.end(extname(normalized) === ".js" ? rewriteBrowserImports(body.toString("utf8")) : body);
};

const server = createHttpServer(requestHandler);

server.listen(port, host, () => {
  console.log(`Open Workbook add-in server listening on ${protocol}://${host}:${port}`);
  console.log(`Manifest: ${protocol}://${host}:${port}/manifest.xml`);
});

function isAllowedPath(filePath) {
  return (
    filePath.startsWith(publicDir) ||
    filePath.startsWith(join(root, "dist")) ||
    Array.from(workspaceModuleDirs.values()).some((moduleDir) => moduleDir && filePath.startsWith(moduleDir))
  );
}

function resolveWorkspaceModule(pathname) {
  for (const [prefix, moduleDir] of workspaceModuleDirs) {
    if (moduleDir && pathname.startsWith(prefix)) {
      return join(moduleDir, pathname.slice(prefix.length));
    }
  }
  return undefined;
}

function firstExistingPath(paths) {
  return paths.find((path) => existsSync(path));
}

function rewriteBrowserImports(source) {
  return source
    .replaceAll("\"@components-kit/open-workbook-excel-core\"", "\"/workspace/excel-core/index.js\"")
    .replaceAll("'@components-kit/open-workbook-excel-core'", "'/workspace/excel-core/index.js'")
    .replaceAll("\"@components-kit/open-workbook-protocol\"", "\"/workspace/protocol/index.js\"")
    .replaceAll("'@components-kit/open-workbook-protocol'", "'/workspace/protocol/index.js'");
}

function getStatus() {
  const addinUrl = trimTrailingSlash(process.env.OPEN_WORKBOOK_ADDIN_URL ?? `${protocol}://${host}:${port}`);
  const backendUrl = process.env.OPEN_WORKBOOK_BACKEND_URL ?? `ws://${process.env.OPEN_WORKBOOK_HOST ?? "127.0.0.1"}:${process.env.OPEN_WORKBOOK_PORT ?? 37845}${process.env.OPEN_WORKBOOK_ADDIN_PATH ?? "/addin"}`;
  return {
    ok: true,
    service: "open-workbook-addin-server",
    packageName: "@components-kit/open-workbook",
    version: runtimeVersion,
    pid: process.pid,
    taskpaneUrl: `${addinUrl}/taskpane.html?backendUrl=${encodeURIComponent(backendUrl)}`,
    backendUrl,
    workspaceModules: Object.fromEntries(
      Array.from(workspaceModuleDirs.entries()).map(([prefix, moduleDir]) => [
        prefix.replace(/^\/workspace\//, "").replace(/\/$/, ""),
        { available: Boolean(moduleDir) }
      ])
    )
  };
}

function generateManifest() {
  const addinUrl = trimTrailingSlash(process.env.OPEN_WORKBOOK_ADDIN_URL ?? `${protocol}://${host}:${port}`);
  const backendUrl = process.env.OPEN_WORKBOOK_BACKEND_URL ?? `ws://${process.env.OPEN_WORKBOOK_HOST ?? "127.0.0.1"}:${process.env.OPEN_WORKBOOK_PORT ?? 37845}${process.env.OPEN_WORKBOOK_ADDIN_PATH ?? "/addin"}`;
  const taskpaneUrl = `${addinUrl}/taskpane.html?backendUrl=${encodeURIComponent(backendUrl)}`;
  return readFileSync(join(root, "manifest.xml"), "utf8")
    .replaceAll("http://localhost:37846/taskpane.html", taskpaneUrl)
    .replaceAll("http://localhost:37846", addinUrl);
}

function trimTrailingSlash(value) {
  return value.endsWith("/") ? value.slice(0, -1) : value;
}

function readPackageVersion() {
  for (const packageJsonPath of [join(root, "package.json"), join(repoRoot, "package.json")]) {
    if (!existsSync(packageJsonPath)) {
      continue;
    }
    try {
      const parsed = JSON.parse(readFileSync(packageJsonPath, "utf8"));
      if (typeof parsed.version === "string") {
        return parsed.version;
      }
    } catch {
      return undefined;
    }
  }
  return undefined;
}
