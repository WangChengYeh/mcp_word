// Integration test per SPEC.md and tool.md
// - Spawns the MCP server (server.js) over stdio via @modelcontextprotocol/sdk Client
// - Connects a Socket.IO client to capture emitted events
// - Calls a subset of MCP tools and asserts corresponding socket events are observed

import path from 'path';
import { fileURLToPath } from 'url';
import process from 'process';
import { setTimeout as delay } from 'timers/promises';
import { Client } from '@modelcontextprotocol/sdk/client/index.js';
import { StdioClientTransport } from '@modelcontextprotocol/sdk/client/stdio.js';
import io from 'socket.io-client';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const PORT = Number(process.env.PORT || 3100);
const KEY_FILE = process.env.KEY_FILE || path.join(__dirname, 'key.pem');
const CERT_FILE = process.env.CERT_FILE || path.join(__dirname, 'cert.pem');

process.env.NODE_TLS_REJECT_UNAUTHORIZED = '0';

function log(...a) { console.log('[test]', ...a); }
function assert(cond, msg) { if (!cond) throw new Error(msg || 'assertion failed'); }

async function main() {
  log('starting integration test on port', PORT);

  // Socket client to observe events emitted by server
  const socket = io(`https://localhost:${PORT}`, { transports: ['websocket'] });

  const events = [];
  socket.onAny((event, payload) => {
    if (String(event).startsWith('word:')) {
      events.push({ event, payload });
    }
  });

  // MCP client spawning server.js via stdio
  const client = new Client({ name: 'mcp-word-test', version: '1.0.0' });
  const transport = new StdioClientTransport({
    command: 'node',
    args: [
      path.join(__dirname, 'server.js'),
      '--key', KEY_FILE,
      '--cert', CERT_FILE,
      '--port', String(PORT),
      '--debug',
    ],
    cwd: __dirname,
  });
  await client.connect(transport);
  log('mcp client connected via stdio');

  // list tools and basic presence checks
  const listed = await client.listTools();
  const toolNames = (listed?.tools || []).map(t => t.name);
  log('tools:', toolNames.join(', '));
  const required = [
    'insertText','getSelection','search','replace','insertPicture',
    'table_create','table_insertRows','table_insertColumns','table_deleteRows','table_deleteColumns','table_setCellText','table_mergeCells',
    'applyStyle','listStyles','ping'
  ];
  required.forEach(n => assert(toolNames.includes(n), `missing tool: ${n}`));

  // Call a couple of tools
  await client.callTool('insertText', { text: 'Hello from test', scope: 'document', location: 'end' });
  await client.callTool('search', { query: 'Hello', scope: 'document', matchWholeWord: false });
  await client.callTool('listStyles', { category: 'paragraph' });

  // Verify socket events observed
  const names = events.map(e => e.event);
  log('observed events:', names);
  assert(names.includes('word:insertText'), 'missing socket event word:insertText');
  assert(names.includes('word:search'), 'missing socket event word:search');
  assert(names.includes('word:listStyles'), 'missing socket event word:listStyles');

  // Clean up
  await client.close();
  try { socket.close(); } catch {}
  log('ok');
}

main().catch((e) => {
  console.error('[test] failed', e);
  process.exit(1);
});

