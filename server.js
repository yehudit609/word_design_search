/*
 * WordFormatSearch – שרת HTTPS מקומי
 * מגיש רק את קבצי ה-Web של התוסף (src, assets). מאזין ל-loopback בלבד.
 * משתמש בתעודת הפיתוח של Office Add-ins (%USERPROFILE%\.office-addin-dev-certs).
 */
"use strict";

const https = require("https");
const fs = require("fs");
const path = require("path");
const os = require("os");

const PORT = Number(process.env.WFS_PORT) || 3000;
const ROOT = __dirname;
const ALLOWED_DIRS = ["src", "assets"];
const LOG_DIR = path.join(ROOT, "logs");
const LOG_FILE = path.join(LOG_DIR, "server.log");

const MIME = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".ico": "image/x-icon",
  ".svg": "image/svg+xml",
  ".json": "application/json; charset=utf-8"
};

function log(msg) {
  const line = new Date().toISOString() + " " + msg + os.EOL;
  try {
    fs.mkdirSync(LOG_DIR, { recursive: true });
    // שומרים את הלוג קטן
    if (fs.existsSync(LOG_FILE) && fs.statSync(LOG_FILE).size > 1024 * 1024) {
      fs.renameSync(LOG_FILE, LOG_FILE + ".old");
    }
    fs.appendFileSync(LOG_FILE, line);
  } catch (_) { /* ignore */ }
  process.stdout.write(line);
}

function loadCertificates() {
  const dir = process.env.WFS_CERT_DIR || path.join(os.homedir(), ".office-addin-dev-certs");
  const key = path.join(dir, "localhost.key");
  const cert = path.join(dir, "localhost.crt");
  const ca = path.join(dir, "ca.crt");
  if (!fs.existsSync(key) || !fs.existsSync(cert)) {
    throw new Error("Dev certificates not found in " + dir +
      ". Run: npx office-addin-dev-certs install");
  }
  const opts = { key: fs.readFileSync(key), cert: fs.readFileSync(cert) };
  if (fs.existsSync(ca)) opts.ca = fs.readFileSync(ca);
  return opts;
}

function send(res, status, body, type) {
  res.writeHead(status, {
    "Content-Type": type || "text/plain; charset=utf-8",
    "Cache-Control": "no-store",            // Word לא ישמור גרסה ישנה של הקבצים
    "X-Content-Type-Options": "nosniff"
  });
  res.end(body);
}

function handler(req, res) {
  if (req.method !== "GET" && req.method !== "HEAD") return send(res, 405, "Method Not Allowed");

  let urlPath;
  try {
    urlPath = decodeURIComponent(new URL(req.url, "https://localhost").pathname);
  } catch (_) {
    return send(res, 400, "Bad Request");
  }
  if (urlPath === "/" || urlPath === "") urlPath = "/src/taskpane.html";

  const filePath = path.normalize(path.join(ROOT, urlPath));
  const rel = path.relative(ROOT, filePath);
  const top = rel.split(path.sep)[0];

  // הגנה מפני path traversal – רק src/ ו-assets/
  if (rel.startsWith("..") || path.isAbsolute(rel) || !ALLOWED_DIRS.includes(top)) {
    return send(res, 404, "Not Found");
  }

  fs.stat(filePath, (err, st) => {
    if (err || !st.isFile()) return send(res, 404, "Not Found");
    const type = MIME[path.extname(filePath).toLowerCase()] || "application/octet-stream";
    if (req.method === "HEAD") return send(res, 200, "", type);
    fs.readFile(filePath, (e, data) => {
      if (e) return send(res, 500, "Internal Server Error");
      send(res, 200, data, type);
    });
  });
}

function listen(options, host) {
  return new Promise((resolve, reject) => {
    const server = https.createServer(options, handler);
    server.once("error", reject);
    server.listen(PORT, host, () => resolve(server));
  });
}

async function main() {
  const options = loadCertificates();

  try {
    await listen(options, "127.0.0.1");
  } catch (e) {
    if (e.code === "EADDRINUSE") {
      log("Port " + PORT + " already in use – assuming the server is already running. Exiting.");
      process.exit(0); // קוד 0 כדי שהמשימה המתוזמנת לא תיכנס ללולאת הפעלה מחדש
    }
    throw e;
  }
  log("WordFormatSearch server listening on https://127.0.0.1:" + PORT);

  // גם IPv6 loopback, כי "localhost" עשוי להתפרש כ-::1
  try {
    await listen(options, "::1");
    log("Also listening on https://[::1]:" + PORT);
  } catch (e) {
    log("IPv6 loopback not available (" + e.code + ") – continuing with IPv4 only.");
  }
}

process.on("uncaughtException", (e) => log("Uncaught exception: " + (e && e.stack || e)));
process.on("unhandledRejection", (e) => log("Unhandled rejection: " + (e && e.stack || e)));

main().catch((e) => {
  log("Failed to start: " + (e && e.message || e));
  process.exit(1);
});
