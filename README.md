# MCP Word — MCP Server + Office Add-in

**Version**: 0.1.0  
**Date**: 2025-08-13
**Language**: English only

AI-driven document editing for Microsoft Word using a local MCP server (stdio) bridged to an Office.js task pane via Socket.IO over HTTPS. This project builds a MCP server (`node server.js`) alongside an Office.js Word Add-in client to enable AI-driven document editing workflows.

## Architecture

```mermaid
flowchart LR
  Codex CLI -- (stdio) -- MCP Server -- (WebSocket) -- Office.js
```

## Components

### MCP Server (`server.js`)
1. Imports `McpServer` from `@modelcontextprotocol/sdk/server/mcp.js`
2. Imports `StdioServerTransport` from `@modelcontextprotocol/sdk/server/stdio.js`
3. Dynamically loads tools via `registerTools` from either `tool.js` (default) or `tool_simple.js` (when `--simple`)
4. Tool schemas follow `schema.json` with simple enums (no `anyOf`)
5. Forwards MCP tool payloads (JSON) to Socket.IO clients
6. Emits `io.emit(toolName, toolParams)` to the Office add‑in

### Office Add-in (`public/`)
- **`manifest.xml`**: Defines Add-in ID, version, provider, display name, description. Host: Document; Permissions: ReadWriteDocument
- **`taskpane.html`**: Loads Office.js and Socket.io client, includes `taskpane.js`, renders button or auto-start behavior
- **`taskpane.js`**: Uses `Office.onReady()` for Word host detection, establishes WebSocket with `io()`, listens for MCP tool events, calls Word functions via `Word.run()`, implements error handling
- **`taskpane.yaml`**: For Script Lab snippet import with libraries including Socket.io

### Snippet Upload
Upload to gist via package.json script:
```bash
npm run snippet  # Runs: cd public && gh gist edit 5b44e6ba1c99baae62ebc0783e1469da --add taskpane.yaml
```

## Workflow

1. **Install the MCP server**: `npm install`
2. **Start the server**: Use stdio from Codex CLI or fake master, use unix pipeline to provide input
3. **Sideload the Add-in**: Load manifest in Word 
4. **Send EditTask requests**: Via Codex CLI or another service, e.g. `{ content: '...' }`
5. **Real-time edits**: The Add-in client receives edits in real time and applies them to the document

## Prerequisites

- Node.js 18.17+
- Microsoft Word (desktop) with sideloading enabled
- Local HTTPS certificate trusted by your OS/Office (self-signed is fine for dev)

## Install

```bash
npm install
```

## Run Server (HTTPS required)

The server requires TLS. Start with PEM key/cert or a PFX bundle.

```bash
# PEM key/cert
node server.js --key path/to/key.pem --cert path/to/cert.pem --port 3000

# or PFX/P12
node server.js --pfx path/to/cert.pfx --passphrase "your-passphrase" --port 3000

# optional verbose logging
node server.js --key key.pem --cert cert.pem --port 3000 --debug

# simple mode (hooks tool_simple.js for flow testing)
node server.js --key key.pem --cert cert.pem --port 3000 --simple
```

Tips:
- Use self-signed certs for dev and mark them trusted so Office/Browser accept `https://localhost:3000`.
- Static files are served from `public/`.
- Endpoints:
  - Health: `GET https://localhost:3000/healthz`

### Create a Local Dev Certificate

```bash
openssl req -x509 -newkey rsa:2048 -sha256 -days 365 -nodes \
  -keyout key.pem -out cert.pem \
  -subj "/CN=localhost" \
  -addext "subjectAltName=DNS:localhost,IP:127.0.0.1"

# Optional PFX/P12 bundle
openssl pkcs12 -export -out cert.pfx -inkey key.pem -in cert.pem -passout pass:your-passphrase
```

Trust the cert in your OS keychain so Word and your browser accept it.

## Configure MCP Client

Point your MCP client to run `server.js` via Node.

Codex CLI (.codex/config.toml):
```toml
# Place in ./.codex/config.toml (project) or ~/.codex/config.toml (user)
[mcp_servers.mcp_word]
command = "node"
args = [
  "/absolute/path/to/server.js",
  "--key", "/abs/path/to/key.pem",
  "--cert", "/abs/path/to/cert.pem",
  "--port", "3000"
]
cwd = "/absolute/path/to/project"
# Optional: env vars
env = { NODE_ENV = "production" }
```

Notes:
- Use absolute paths for reliability.
- If using a PFX/P12 bundle: replace `--key/--cert` with `--pfx /abs/cert.pfx --passphrase "your-passphrase"`.
- Restart the Codex Client after saving the config; tools `ping` and `editTask` should appear.

Claude Desktop (settings excerpt):
```json
{
  "mcpServers": {
    "mcp-word": {
      "command": "node",
      "args": ["/absolute/path/to/server.js", "--key", "/abs/key.pem", "--cert", "/abs/cert.pem", "--port", "3000"],
      "cwd": "/absolute/path/to/project"
    }
  }
}
```

Claude CLI:
```bash
claude mcp add mcp-word -- node /absolute/path/to/server.js --key /abs/key.pem --cert /abs/cert.pem --port 3000
```

## Office Add-in

Two ways to connect Word to the server:

1) Sideload manifest
- Open Word → Add-ins → Sideload `public/manifest.xml`
- Manifest points to `public/taskpane.html`
- Ensure the same host/port and a trusted certificate

2) Script Lab (alternative)
- Option A (paste JS): Install Script Lab, create a new script and paste `public/taskpane.js`, and add library `https://cdn.socket.io/4.7.5/socket.io.min.js`.
- Option B (import YAML): In Script Lab, import from https://gist.github.com/WangChengYeh/5b44e6ba1c99baae62ebc0783e1469da

## Tools

### editTask
- **Purpose**: Send an edit instruction to the add-in via Socket.IO (event name `editTask`)
- **Schema**: Follows `schema.json` rules with simple enums (no `anyOf` patterns)
- **Args**:
  - `content` (string, required): text to insert/replace
  - `action` ("insert" | "replace" | "append", default "insert")  
  - `target` ("cursor" | "selection" | "document", default "selection")
  - `taskId` (string, optional)
  - `meta` (object, optional)

JSON-RPC example (MCP stdio frame):
```json
{
  "jsonrpc": "2.0",
  "id": 2,
  "method": "tools/call",
  "params": {
    "name": "editTask",
    "arguments": {
      "content": "Hello from MCP",
      "action": "insert",
      "target": "selection"
    }
  }
}
```

### ping  
- **Purpose**: Health check; echoes `message` or returns `"pong"`
- **Schema**: Simple tool registration without complex patterns

## Debugging

- `--debug`: Dumps detailed logs and errors to `debug.log`
- Records stdio stream input/output into `debug.log`
- Adds pipes after stdin and before stdout to record and forward if debug enabled
- Records JSON strings before Socket.IO send and after Socket.IO receive in `debug.log`
- Health endpoint: `GET https://localhost:3000/healthz` shows connected clients
- Stream logging in `debug.log`:
  - `[time][stdin] json string`: raw MCP stdio frames received
  - `[time][stdout] json string`: raw MCP stdio responses
  - `[time][socket:send] json string`: forwarded tool calls to Socket.IO
  - `[time][socket.recv] json string`: received back from the add-in (e.g., `edit-complete`)

## Testing

### Two Modes at a Glance

- Simple mode: `server.js --simple` hooks `tool_simple.js` (flow tests)
  - Unit: `npm run test:simple`
  - Integration: `npm run test:int:simple`
- Normal mode: `server.js` (no `--simple`) hooks `tool.js` (regular usage)
  - Unit: `npm test`
  - Integration: `npm run test:int`

### Unit Tests
- `test.sh`: Unit test with fake stdio for MCP client and socket connection for Office
- Uses shell pipeline to provide input
- Generates test JavaScript as socket client
- Prepare `package.json` before testing
- Default test port: 3100 (3000 reserved for normal use)
- `test_simple.sh`: Tests `server.js --simple` mode using `tool_simple.js`

Run via npm commands:
```bash
# Simple mode unit test
npm run test:simple

# Full unit test
npm test
```

### Integration Tests
- `test.js`: Integration test combining MCP client + MCP server (`server.js`)
- Uses `Client` from `@modelcontextprotocol/sdk/client/index.js`
- Lists tools and calls tools (e.g., `ping`, `editTask`)
- `test_simple.js`: Tests `server.js --simple` mode

Run via npm commands:
```bash
# Simple mode integration test
npm run test:int:simple

# Full integration test
npm run test:int
```

### End-to-End Testing
Use the provided script to verify end‑to‑end behavior. It auto-generates a dev cert, starts the server, opens a Socket.IO client, and sends MCP frames via stdio.

```bash
# Default smoke test
./test.sh

# Pipe custom MCP JSONL (one JSON object per line)
echo '{"jsonrpc":"2.0","id":10,"method":"tools/call","params":{"name":"editTask","arguments":{"content":"PipeMsg","action":"insert","target":"selection"}}}' | ./test.sh

# Multiple lines
cat <<'EOF' | ./test.sh
{"jsonrpc":"2.0","id":11,"method":"tools/call","params":{"name":"ping","arguments":{"message":"hello"}}}
{"jsonrpc":"2.0","id":12,"method":"tools/call","params":{"name":"editTask","arguments":{"content":"FromSTDIN","action":"insert","target":"selection"}}}
EOF
```

Notes:
- The proxy emits Socket.IO events named by the MCP tool. For `editTask`, the event is `editTask` and the payload is the tool arguments object.
- The script extracts expected content from the first `editTask` line.
- Accepts self-signed certs for test convenience.

## Extensibility

- Support additional `EditTask` types (tables, images, formatting)
- Add WebSocket authentication, logging, and error tracking

## Project Structure

```
mcp_word/
├── server.js          # MCP stdio + Socket.IO bridge (HTTPS)
├── tool.js            # MCP tool registration + Socket.IO forwarding
├── tool_simple.js     # Simplified tool set for flow tests (--simple)
├── public/
│   ├── manifest.xml   # Office add-in manifest
│   ├── taskpane.html  # Minimal task pane
│   ├── taskpane.js    # Applies edit commands via Office.js
│   └── taskpane.yaml  # Script Lab snippet for import
├── test.sh            # E2E test runner (JSONL over stdio)
├── test_simple.sh     # Unit test runner for --simple
├── test.js            # Integration test (normal mode)
├── test_simple.js     # Integration test (simple mode)
├── SPEC.md            # Refined spec
└── README.md
```

## Documentation Requirements

This README follows the SPEC requirements and includes (in order):
- Prerequisites: Node.js version, Microsoft Word, and trusted HTTPS cert note
- Install: `npm install`
- Run server: HTTPS launch examples (PEM and PFX), plus `--debug`
- Create a local dev certificate: OpenSSL commands to generate PEM and PFX
- Configure MCP client: Codex Client setup via `.codex/config.toml`
- Office add-in: Two options (Sideload manifest or Script Lab snippets)
- Tools: Document `editTask` and `ping` tools with args and examples
- Debugging: `--debug` and debug.log details
- Testing: How to run tests by npm command
- Project structure and License

## License

See LICENSE.
