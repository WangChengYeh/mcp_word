#!/usr/bin/env bash
set -euo pipefail

PORT="${PORT:-3000}"

# Read STDIN (non-interactive) as JSONL input
STDIN_BUFFER=""
if [ ! -t 0 ]; then
  # Allow empty input without blocking
  if STDIN_BUFFER="$(cat)"; then
    :
  fi
fi

if [ ! -d node_modules ]; then
  echo "[test] installing dependencies..."
  npm install
fi

echo "[test] running e2e on port ${PORT}..."
TEST_INPUT="${STDIN_BUFFER}" PORT="${PORT}" node --input-type=module - <<'NODE'
import { spawn } from 'node:child_process';
import http from 'node:http';
import { io } from 'socket.io-client';

const port = Number(process.env.PORT || 3000);
const userInputRaw = process.env.TEST_INPUT || '';
const userLines = userInputRaw.trim()
  ? userInputRaw.split(/\r?\n/).map(l => l.trim()).filter(l => l.length)
  : [];

const server = spawn(process.execPath, ['server.js', '--port', String(port), '--debug'], {
  stdio: ['pipe', 'pipe', 'inherit']
});

const wait = ms => new Promise(r => setTimeout(r, ms));

async function waitHealthz(timeoutMs = 10000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      await new Promise((resolve, reject) => {
        const req = http.get({ hostname: '127.0.0.1', port, path: '/healthz', timeout: 800 }, (res) => {
          res.resume();
          (res.statusCode === 200) ? resolve() : reject(new Error('bad status'));
        });
        req.on('error', reject);
      });
      return;
    } catch {
      await wait(200);
    }
  }
  throw new Error('healthz not ready');
}

function sendMcp(obj) {
  const data = JSON.stringify(obj);
  const len = Buffer.byteLength(data, 'utf8');
  const frame = `Content-Length: ${len}\r\n\r\n${data}`;
  server.stdin.write(frame);
}

function gracefulExit(code = 0) {
  try { server.kill('SIGTERM'); } catch {}
  process.exit(code);
}

(async () => {
  try {
    await waitHealthz();

    const socket = io(`http://127.0.0.1:${port}`, { transports: ['websocket'] });

    await new Promise((resolve, reject) => {
      const t = setTimeout(() => reject(new Error('socket connect timeout')), 8000);
      socket.on('connect', () => { clearTimeout(t); console.log('SOCKET_READY'); resolve(); });
      socket.on('connect_error', () => {});
    });

    // Basic MCP handshake
    sendMcp({
      jsonrpc: '2.0',
      id: 1,
      method: 'initialize',
      params: {
        protocolVersion: '2024-11-05',
        clientInfo: { name: 'e2e-test', version: '0.0.1' }
      }
    });
    sendMcp({ jsonrpc: '2.0', method: 'notifications/initialized', params: {} });

    let expectedContent = null;

    if (userLines.length) {
      for (const line of userLines) {
        try {
          const obj = JSON.parse(line);
          sendMcp(obj);
          if (!expectedContent &&
              obj.method === 'tools/call' &&
              obj?.params?.name === 'editTask' &&
              obj?.params?.arguments?.content) {
            expectedContent = obj.params.arguments.content;
          }
        } catch (e) {
          console.error('[test] skip invalid JSON line:', line, e.message);
        }
      }
    } else {
      // If no user-provided JSONL: send default editTask
      const msg = 'HelloFromTest-' + Math.random().toString(36).slice(2);
      expectedContent = msg;
      sendMcp({
        jsonrpc: '2.0',
        id: 2,
        method: 'tools/call',
        params: {
          name: 'editTask',
            arguments: { content: msg, action: 'insert', target: 'selection' }
        }
      });
    }

    await new Promise((resolve, reject) => {
      const t = setTimeout(() => reject(new Error('ai-cmd not received')), 10000);
      socket.on('ai-cmd', (data) => {
        if (!expectedContent) {
          console.log('TEST_PASS');
          clearTimeout(t);
          resolve();
          return;
        }
        if (data && data.content === expectedContent) {
          console.log('TEST_PASS');
          clearTimeout(t);
          resolve();
        }
      });
    });

    socket.close();
    gracefulExit(0);
  } catch (e) {
    console.error('[test] error:', e && e.stack ? e.stack : e);
    gracefulExit(1);
  }
})();
NODE
