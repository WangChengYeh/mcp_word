#!/usr/bin/env node

// Minimal MCP-like server over STDIO + Socket.IO bridge per SPEC.md

import express from "express";
import http from "http";
import { Server as IOServer } from "socket.io";
import fs from "fs";
import path from "path";
import process from "process";
import { fileURLToPath } from "url";
import { z } from "zod";

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
      "Usage: node server.js [--port 3000] [--debug]",
      "",
      "Options:",
      "  --port   HTTP/Socket.IO 監聽的埠號（預設 3000）",
      "  --debug  啟用除錯日誌（寫入 debug.log）",
    ].join("\n")
  );
  process.exit(0);
}
const PORT = parseInt(getFlagVal("port", "3000"), 10);
const DEBUG = hasFlag("debug");

// ---------- logger ----------
const debugLogPath = path.join(process.cwd(), "debug.log");
function log(...args) {
  const line = `[${new Date().toISOString()}] ${args
    .map((a) => (typeof a === "string" ? a : JSON.stringify(a)))
    .join(" ")}`;
  // IMPORTANT: write logs to stderr to keep stdout clean for MCP stdio
  console.error(line);
  if (DEBUG) {
    fs.appendFile(debugLogPath, line + "\n", () => {});
  }
}
function logErr(err, ctx = "error") {
  const msg =
    err && err.stack
      ? err.stack
      : typeof err === "string"
      ? err
      : JSON.stringify(err);
  console.error(`[${new Date().toISOString()}] ${ctx}: ${msg}`);
  if (DEBUG) {
    fs.appendFile(
      debugLogPath,
      `[${new Date().toISOString()}] ${ctx}: ${msg}\n`,
      () => {}
    );
  }
}

// ---------- HTTP + Socket.IO ----------
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
app.use(express.json());

// 公開目錄：public/（符合 SPEC 的 Office Add-in 靜態檔）
const publicDir = path.join(process.cwd(), "public");
app.use(express.static(publicDir));

// 健康檢查端點
app.get("/healthz", (req, res) => {
  try {
    const clients = io.engine ? io.engine.clientsCount : 0;
    res.json({ ok: true, clients, port: PORT });
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e) });
  }
});

const server = http.createServer(app);
const io = new IOServer(server, {
  cors: {
    origin: "*",
  },
});

io.on("connection", (socket) => {
  log("socket connected", { id: socket.id });
  socket.on("disconnect", (reason) => {
    log("socket disconnected", { id: socket.id, reason });
  });
});

// ---------- MCP Server (STDIO) ----------
const mcp = new McpServer({
  name: "mcp-word",
  version: "1.0.0",
});

// 工具：editTask -> 透過 ai-cmd 廣播給 Add-in
mcp.registerTool(
  "editTask",
  {
    description:
      "Send an edit task to the Office Add-in via WebSocket (event: ai-cmd).",
    inputSchema: {
      content: z
        .string()
        .describe("Text to insert/replace in the Word document."),
      action: z
        .enum(["insert", "replace", "append"])
        .default("insert")
        .describe("insert | replace | append"),
      target: z
        .enum(["cursor", "selection", "document"]) 
        .default("selection")
        .describe("cursor | selection | document"),
      taskId: z.string().optional().describe("Optional client-correlated id."),
      meta: z.record(z.unknown()).optional().describe("Optional additional metadata."),
    },
  },
  async (args, _ctx) => {
    try {
      const payload = {
        type: "EditTask",
        content: args.content,
        action: args.action || "insert",
        target: args.target || "selection",
        taskId: args.taskId || null,
        meta: args.meta || {},
        ts: Date.now(),
      };
      io.emit("ai-cmd", payload);
      log("emitted ai-cmd", payload);
      return {
        content: [{ type: "text", text: "EditTask sent to client." }],
      };
    } catch (e) {
      logErr(e, "editTask");
      return {
        isError: true,
        content: [{ type: "text", text: `editTask failed: ${String(e)}` }],
      };
    }
  }
);

// 工具：ping -> 回傳 pong 或輸入訊息
mcp.registerTool(
  "ping",
  {
    description: "Health check tool.",
    inputSchema: {
      message: z.string().optional(),
    },
  },
  async (args, _ctx) => {
    const text = args?.message || "pong";
    return { content: [{ type: "text", text }] };
  }
);

// ---------- bootstrap ----------
async function main() {
  // 啟動 HTTP/Socket
  await new Promise((resolve) => {
    server.listen(PORT, () => {
      log(`HTTP/Socket.IO listening on http://localhost:${PORT}`);
      resolve();
    });
  });

  // 在 debug 模式下，監看 STDIN 原始資料以協助除錯 MCP 流
  if (DEBUG) {
    try {
      process.stdin.on("data", (chunk) => {
        // 只印前 200 字元避免洗版
        const s = chunk.toString("utf8");
        console.error(`[DEBUG stdin] ${s.slice(0, 200).replace(/\n/g, "\\n")}${
          s.length > 200 ? "..." : ""
        }`);
      });
    } catch {}
  }

  // 連線 MCP STDIO
  const transport = new StdioServerTransport();
  await mcp.connect(transport);
  log("MCP server connected via STDIO");

  // 優雅關閉
  const shutdown = async (signal = "SIGTERM") => {
    try {
      log(`shutting down (${signal})...`);
      server.close(() => log("HTTP server closed"));
      if (io && io.close) io.close();
      if (transport && transport.close) await transport.close();
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
