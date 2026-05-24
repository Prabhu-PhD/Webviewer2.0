import { createServer as createHttpServer } from "node:http";
import { createServer as createHttpsServer } from "node:https";
import { access, readFile } from "node:fs/promises";
import { extname, join, normalize, resolve } from "node:path";

const projectRoot = resolve(process.cwd());
const webRoot = join(projectRoot, "web");
const host = process.env.HOST ?? "localhost";
const port = Number.parseInt(process.env.PORT ?? "3000", 10);

const mimeTypes = {
  ".css": "text/css; charset=utf-8",
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".json": "application/json; charset=utf-8",
  ".png": "image/png",
  ".svg": "image/svg+xml",
  ".txt": "text/plain; charset=utf-8"
};

async function buildTlsOptions() {
  const defaultPfxFile = join(projectRoot, "certs", "localhost.pfx");
  const defaultPassphraseFile = join(projectRoot, "certs", "localhost.passphrase.txt");
  const pfxFile = process.env.SSL_PFX_FILE ?? (await exists(defaultPfxFile) ? defaultPfxFile : null);
  const keyFile = process.env.SSL_KEY_FILE;
  const certFile = process.env.SSL_CERT_FILE;

  if (pfxFile) {
    const passphrase =
      process.env.SSL_PFX_PASSPHRASE ??
      (await exists(defaultPassphraseFile) ? (await readFile(defaultPassphraseFile, "utf8")).trim() : "");

    return {
      pfx: await readFile(resolve(projectRoot, pfxFile)),
      passphrase
    };
  }

  if (keyFile && certFile) {
    return {
      key: await readFile(resolve(projectRoot, keyFile)),
      cert: await readFile(resolve(projectRoot, certFile))
    };
  }

  return null;
}

async function exists(path) {
  try {
    await access(path);
    return true;
  } catch (error) {
    return false;
  }
}

function safePathname(urlPathname) {
  const normalizedPath = normalize(decodeURIComponent(urlPathname));
  const relativePath = normalizedPath === "\\" || normalizedPath === "/" ? "content.html" : normalizedPath.replace(/^[/\\]+/, "");
  const fullPath = resolve(webRoot, relativePath);

  if (!fullPath.startsWith(webRoot)) {
    return null;
  }

  return fullPath;
}

async function handleRequest(request, response) {
  try {
    const requestUrl = new URL(request.url ?? "/", `http://${host}:${port}`);
    const filePath = safePathname(requestUrl.pathname);

    if (!filePath) {
      response.writeHead(403, { "Content-Type": "text/plain; charset=utf-8" });
      response.end("Forbidden");
      return;
    }

    const buffer = await readFile(filePath);
    const extension = extname(filePath).toLowerCase();
    const contentType = mimeTypes[extension] ?? "application/octet-stream";

    response.writeHead(200, {
      "Cache-Control": "no-store",
      "Content-Type": contentType
    });
    response.end(buffer);
  } catch (error) {
    const isMissing = error && typeof error === "object" && "code" in error && error.code === "ENOENT";
    response.writeHead(isMissing ? 404 : 500, {
      "Content-Type": "text/plain; charset=utf-8"
    });
    response.end(isMissing ? "Not found" : "Internal server error");
  }
}

function createAppServer(tlsOptions) {
  return tlsOptions
    ? createHttpsServer(tlsOptions, handleRequest)
    : createHttpServer(handleRequest);
}

function resolveListenHosts(requestedHost) {
  return requestedHost === "localhost"
    ? ["127.0.0.1", "::1"]
    : [requestedHost];
}

function listen(server, listenHost) {
  return new Promise((resolvePromise, rejectPromise) => {
    server.once("error", rejectPromise);
    server.listen(port, listenHost, () => {
      server.removeListener("error", rejectPromise);
      resolvePromise();
    });
  });
}

const tlsOptions = await buildTlsOptions();
const listenHosts = resolveListenHosts(host);
const servers = listenHosts.map(() => createAppServer(tlsOptions));

await Promise.all(servers.map((server, index) => listen(server, listenHosts[index])));

const scheme = tlsOptions ? "https" : "http";
console.log(`[webviewer2] Serving ${webRoot}`);
console.log(`[webviewer2] ${scheme}://localhost:${port}/content.html`);
if (listenHosts.length > 1) {
  console.log(`[webviewer2] Bound loopback hosts: ${listenHosts.join(", ")}`);
}
if (!tlsOptions) {
  console.log("[webviewer2] HTTPS certificates were not provided, so this is browser-preview only.");
}
