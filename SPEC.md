---
# SPEC: Proxy Server & Office Add-in

# Version: 1.0.0  
# Date: 2025-08-07  
# Author: Your Name

## 1. Introduction
MCP_WORD build a MCP server (`node server.js`) alongside an Office.js Word Add-in client to enable AI-driven document editing workflows.

## 2. Architecture Overview
```mermaid
flowchart LR
  Codex CLI -- (stdio) -- MCP Server -- (WebSocket) -- Office.js
```


## 3. Components

### 3.1 MCP Server (`server.js`)
  1. import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js"
  2. import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js"
  3. server.registerTool to add Tool in MCP server (in file tool.js)
  4. forward the payload (MCP tool json string) to socket.io (in file tool.js)
  5. tool handling only in tool.js, server.js only forward json string or json object
  6. io.emit(tool name, tool params) to office add-in
  7. office add-in: socket.io

### 3.2 Office Add-in (public/)
#### `manifest.xml`
- Defines Add-in ID, version, provider name, display name, and description
- Host: Document; Permissions: ReadWriteDocument

#### `taskpane.html`
- Loads Office.js and the Socket.io client
- Includes `taskpane.js` and renders a button or auto-start behavior

#### `taskpane.js`
- Uses `Office.onReady()` to detect the Word host
- Establishes a WebSocket connection with `io()`
- Listens for 'MCP tool' events and call Word functions via `Word.run()`
- Implements basic error handling

#### `taskpane.yaml` (for Snippet import)
-- script: as taskpane.js
-- template: as taskpane.html
-- libraries: add socket.io library
### upload to gist
-- package.json: "snippet": "cd public ; gh gist edit 5b44e6ba1c99baae62ebc0783e1469da --add taskpane.yaml"
## 4. Workflow
1. install the MCP server: `npm install`
2. stdio from Codex CLI or fake master, use unix pipeline to provide input
3. Sideload the Add-in manifest in Word
4. Send `EditTask` requests via Codex CLI or another service, e.g. `{ content: '...' }`
5. The Add-in client receives edits in real time and applies them to the document

## 5. Extensibility
- Support additional `EditTask` types (tables, images, formatting)
- Add WebSocket authentication, logging, and error tracking

## 6. Debugging
- add argument --debug, dump detail and error in debug.log 
- Record stdio stream into debug.log
- Add a pipe after stdin and a pipe before stdout to record and forward if debug
- Record json string before send socket and after receive socket in debug.log

## 7. Test
- test.sh Unit test, Fake stdio for MCP client and socket connection for office
- stdio: use shell pipeline to provide input
- socket: generate a test javascript as a socket client
- before test, prepare package.json
- default test port: 3100 (3000 reserved for normal use)
## 8. Integration Test
- test.js Integration test: MCP client + MCP server (server.js)
- MCP client: import { Client } from "@modelcontextprotocol/sdk/client/index.js"
- MCP client: list tools and call tools (e.g., ping, editTask)
## 9. Documentation

### 9.1 README.md requirements (step-by-step)
The project README must include the following, in order:
- Prerequisites: Node.js version, Microsoft Word, and trusted HTTPS cert note.
- Install: `npm install`.
- Run server: HTTPS launch examples (PEM and PFX), plus `--debug`.
- Create a local dev certificate: OpenSSL commands to generate PEM and PFX.
- Configure MCP client: Codex Client setup via `.codex/config.toml` (see 9.2).
- Office add-in: two options:
  - Sideload `public/manifest.xml` in Word (points to `public/taskpane.xml`).
  - Script Lab alternative: import snippets from https://gist.github.com/WangChengYeh/5b44e6ba1c99baae62ebc0783e1469da
- Tools: document `editTask` (args and example frame) and `ping`.
- Debugging: `--debug` behavior and `GET /healthz`.
- Testing: how to run `./test.sh` and pipe custom MCP JSONL.
- Project structure and License.

### 9.2 Codex Client (Codex CLI) setup
Document how to point Codex to the MCP server using a TOML config. Provide both project-local and user-level options.

1) Config file location
- Preferred (project-local): `./.codex/config.toml`
- User-level (global): `~/.codex/config.toml`

2) Minimal configuration snippet
```toml
# ./.codex/config.toml or ~/.codex/config.toml
[mcp_servers.mcp_word]
command = "node"
args = [
  "/absolute/path/to/server.js",
  "--key", "/abs/path/to/key.pem",
  "--cert", "/abs/path/to/cert.pem",
  "--port", "3000"
]
cwd = "/absolute/path/to/project"
# Optional environment variables
env = { NODE_ENV = "production" }
```

Notes:
- Use absolute paths for reliability across shells.
- If using a PFX/P12 bundle instead of PEM:
  - Replace `--key/--cert` with `--pfx /abs/cert.pfx --passphrase "your-passphrase"` in `args`.
- Ensure the TLS certificate is trusted by your OS so Word and the browser accept `https://localhost:3000`.
- After saving the config, restart the Codex Client so it picks up changes. You should see tools `ping` and `editTask` available.

### 9.3 Language
- Code and documentation: English only.
