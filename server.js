#!/usr/bin/env node

// Minimal MCP-like server over STDIO + Socket.IO bridge per SPEC.md

import express from "express";
import https from "https";
import { Server as IOServer } from "socket.io";
import fs from "fs";
import path from "path";
import process from "process";
import { fileURLToPath } from "url";
// schemas and tool registration moved to tool.js

// MCP SDK
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";

// ---------- CLI args ----------
const argv = process.argv.slice(2);
const hasFlag = (name) => argv.includes(`--${name}`);
const getFlagVal = (name, def) => {
  const idx = argv.indexOf(`--${name}`);
  if (idx >= 0 && idx + 1 < argv.length && !argv[idx + 1].startsWith("--")) {
    return argv[idx + 1];
  }
  return def;
};
if (hasFlag("help")) {
  console.log(
    [
      "Usage: node server.js [--port 3000] [--key key.pem --cert cert.pem] [--pfx bundle.pfx --passphrase XXX] [--debug] [--simple]",
      "",
      "Options:",
      "  --port   HTTPS Socket.IO listen port (default 3000)",
      "  --key    TLS private key file (PEM)",
      "  --cert   TLS certificate file (PEM)",
      "  --pfx    TLS certificate bundle (PFX/P12)",
      "  --passphrase  TLS certificate passphrase (if required)",
      "  --debug  Enable debug logging (writes debug.log)",
    ].join("\n")
  );
  process.exit(0);
}
const PORT = parseInt(getFlagVal("port", "3000"), 10);
const DEBUG = hasFlag("debug");
const SIMPLE = hasFlag("simple");
const HTTPS_KEY = getFlagVal("key");
const HTTPS_CERT = getFlagVal("cert");
const HTTPS_PFX = getFlagVal("pfx");
const HTTPS_PASSPHRASE = getFlagVal("passphrase");

// ---------- logger ----------
const debugLogPath = path.join(process.cwd(), "debug.log");
function log(...args) {
  const line = `[${new Date().toISOString()}] ${args
    .map((a) => (typeof a === "string" ? a : JSON.stringify(a)))
    .join(" ")}`;
  // IMPORTANT: write logs to stderr to keep stdout clean for MCP stdio
  console.error(line);
}
function logErr(err, ctx = "error") {
  const msg =
    err && err.stack
      ? err.stack
      : typeof err === "string"
      ? err
      : JSON.stringify(err);
  console.error(`[${new Date().toISOString()}] ${ctx}: ${msg}`);
}

// ---------- HTTPS + Socket.IO ----------
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(express.json());

// Serve static Office Add-in files from public/
const publicDir = path.join(process.cwd(), "public");
app.use(express.static(publicDir));

// Health check endpoint
app.get("/healthz", (req, res) => {
  try {
    const clients = io.engine ? io.engine.clientsCount : 0;
    res.json({ ok: true, clients, port: PORT });
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e) });
  }
});

const tlsOpts = {};
try {
  if (HTTPS_PFX) {
    tlsOpts.pfx = fs.readFileSync(path.resolve(HTTPS_PFX));
    if (HTTPS_PASSPHRASE) tlsOpts.passphrase = HTTPS_PASSPHRASE;
  } else {
    if (!HTTPS_KEY || !HTTPS_CERT) {
      throw new Error("TLS required: provide --key and --cert (or use --pfx)");
    }
    tlsOpts.key = fs.readFileSync(path.resolve(HTTPS_KEY));
    tlsOpts.cert = fs.readFileSync(path.resolve(HTTPS_CERT));
  }
} catch (e) {
  console.error("Failed to read TLS materials:", e?.message || e);
  process.exit(1);
}
const server = https.createServer(tlsOpts, app);
const io = new IOServer(server, {
  cors: {
    origin: "*",
  },
});

io.on("connection", (socket) => {
  log("socket connected", { id: socket.id });
  // Log any events coming back from the add-in (status, edit-complete, edit-error, etc.)
  try {
    socket.onAny((event, payload) => {
      try {
        // Keep stderr visibility in debug, but avoid writing to debug.log
        if (DEBUG) console.error(`[socket:recv] ${event} ${JSON.stringify(payload || {})}`);
      } catch {}
    });
  } catch {}
  socket.on("disconnect", (reason) => {
    log("socket disconnected", { id: socket.id, reason });
  });
});

// ---------- MCP Server (STDIO) ----------
const mcp = new McpServer({
  name: "mcp-word",
  version: "1.0.0",
});

// Tools are loaded dynamically (tool.js by default, tool_simple.js with --simple)

// ---------- bootstrap ----------
async function main() {
  // Start HTTPS/Socket
  await new Promise((resolve) => {
    server.listen(PORT, () => {
      log(`HTTPS/Socket.IO listening on https://localhost:${PORT}`);
      resolve();
    });
  });

  // In debug mode, simply dump raw STDIN bytes to debug.log (no frame parsing)
  let debugLogStream = null;
  if (DEBUG) {
    try {
      debugLogStream = fs.createWriteStream(debugLogPath, { flags: "a" });
      process.stdin.on("data", (chunk) => {
        try {
          debugLogStream.write(chunk);
        } catch {}
      });

      // Tee STDOUT to debug.log without altering MCP output
      try {
        const originalWrite = process.stdout.write.bind(process.stdout);
        process.stdout.write = function (chunk, encoding, cb) {
          try {
            if (debugLogStream) {
              // Preserve encoding when provided for string writes
              const enc = typeof encoding === "string" ? encoding : undefined;
              debugLogStream.write(chunk, enc);
            }
          } catch {}
          return originalWrite(chunk, encoding, cb);
        };
      } catch {}
    } catch {}
  }

  // Connect MCP STDIO
  // Register tools BEFORE connecting transport (dynamic import based on --simple)
  try {
    const modulePath = SIMPLE ? "./tool_simple.js" : "./tool.js";
    const mod = await import(modulePath);
    if (!mod || typeof mod.registerTools !== "function") {
      throw new Error(`registerTools not found in ${modulePath}`);
    }
    mod.registerTools(mcp, io, log, logErr);
    log(`Tools registered from ${modulePath}`);
  } catch (e) {
    logErr(e, "tools");
    process.exit(1);
  }

  const transport = new StdioServerTransport();
  await mcp.connect(transport);
  log("MCP server connected via STDIO");

  // Graceful shutdown
  const shutdown = async (signal = "SIGTERM") => {
    try {
      log(`shutting down (${signal})...`);
      server.close(() => log("HTTPS server closed"));
      if (io && io.close) io.close();
      if (transport && transport.close) await transport.close();
      if (debugLogStream) {
        try { debugLogStream.end(); } catch {}
      }
    } catch (e) {
      logErr(e, "shutdown");
    } finally {
      process.exit(0);
    }
  };
  process.on("SIGINT", () => shutdown("SIGINT"));
  process.on("SIGTERM", () => shutdown("SIGTERM"));
}

main().catch((e) => {
  logErr(e, "bootstrap");
  process.exit(1);
});
