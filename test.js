#!/usr/bin/env node
// Integration test: MCP client + server.js + Socket.IO bridge
// - Spawns server.js via MCP StdioClientTransport
// - Lists tools and calls ping + editTask
// - Verifies ai-cmd over WebSocket contains expected content

import path from 'node:path';
import https from 'node:https';
import process from 'node:process';
import { fileURLToPath } from 'node:url';
import { io } from 'socket.io-client';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Accept self-signed certs for local HTTPS in tests
process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';

const PORT = Number(process.env.PORT || 3100);
const KEY = process.env.KEY_PATH || path.resolve(__dirname, 'key.pem');
const CERT = process.env.CERT_PATH || path.resolve(__dirname, 'cert.pem');

function wait(ms) {
  return new Promise((resolve) => setTimeout(resolve, ms));
}

async function waitHealthz(timeoutMs = 10000) {
  const deadline = Date.now() + timeoutMs;
  while (Date.now() < deadline) {
    try {
      await new Promise((resolve, reject) => {
        const req = https.get(
          {
            hostname: '127.0.0.1',
            port: PORT,
            path: '/healthz',
            timeout: 800,
            rejectUnauthorized: false,
          },
          (res) => {
            res.resume();
            return res.statusCode === 200 ? resolve() : reject(new Error('bad status'));
          }
        );
        req.on('error', reject);
      });
      return; // healthy
    } catch {
      await wait(200);
    }
  }
  throw new Error('healthz not ready');
}

async function main() {
  // Spawn server.js via MCP stdio transport
  const transport = new StdioClientTransport({
    command: process.execPath,
    args: [
      path.resolve(__dirname, 'server.js'),
      '--port', String(PORT),
      '--key', KEY,
      '--cert', CERT,
      '--debug',
    ],
    cwd: __dirname,
    stderr: 'inherit',
  });

  const client = new Client({ name: 'integration-test', version: '1.0.0' });

  try {
    // Connect MCP client (this starts the process)
    await client.connect(transport);

    // Wait for HTTPS/socket side to be ready
    await waitHealthz();

    // Connect socket.io client to capture ai-cmd
    const socket = io(`https://127.0.0.1:${PORT}`, {
      transports: ['websocket'],
      secure: true,
      transportOptions: { websocket: { rejectUnauthorized: false } },
    });

    await new Promise((resolve, reject) => {
      const t = setTimeout(() => reject(new Error('socket connect timeout')), 8000);
      socket.on('connect', () => { clearTimeout(t); resolve(); });
      socket.on('connect_error', () => {});
    });

    // Verify tools list contains expected tools
    const list = await client.listTools({});
    console.error('List tools', JSON.stringify(list,null, 2));
    const toolNames = new Set(list.tools.map((t) => t.name));
    if (!toolNames.has('ping') || !toolNames.has('editTask')) {
      throw new Error(`Expected tools ping and editTask, got: ${[...toolNames].join(', ')}`);
    }

    // Call ping tool and assert response
    const pingMsg = 'hello-from-integration';
    const pingRes = await client.callTool({ name: 'ping', arguments: { message: pingMsg } });
    const pingText = Array.isArray(pingRes.content) && pingRes.content[0]?.text;
    if (!pingText || !String(pingText).includes(pingMsg)) {
      throw new Error(`Unexpected ping result: ${JSON.stringify(pingRes)}`);
    }

    // Call editTask and expect ai-cmd with matching content
    const expected = 'FromIntegrationTest-' + Math.random().toString(36).slice(2);
    const aiCmdPromise = new Promise((resolve, reject) => {
      const t = setTimeout(() => reject(new Error('ai-cmd timeout')), 10000);
      socket.on('ai-cmd', (data) => {
        if (data && data.content === expected) {
          clearTimeout(t);
          resolve();
        }
      });
    });

    await client.callTool({
      name: 'editTask',
      arguments: { content: expected, action: 'insert', target: 'selection' },
    });

    await aiCmdPromise;

    socket.close();
    await client.close();
    console.log('TEST_PASS');
    process.exit(0);
  } catch (e) {
    try { await client.close(); } catch {}
    console.error('TEST_FAIL', e && e.stack ? e.stack : String(e));
    process.exit(1);
  }
}

main();
