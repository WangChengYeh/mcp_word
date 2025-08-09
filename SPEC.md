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
  3. server.registerTool to add Tool in MCP server
  4. office add-in: socket.io

### 3.2 Office Add-in (public/)
#### `manifest.xml`
- Defines Add-in ID, version, provider name, display name, and description
- Host: Document; Permissions: ReadWriteDocument
- SourceLocation: `http://localhost:3000/taskpane.html`

#### `taskpane.html`
- Loads Office.js and the Socket.io client
- Includes `taskpane.js` and renders a button or auto-start behavior

#### `taskpane.js`
- Uses `Office.onReady()` to detect the Word host
- Establishes a WebSocket connection with `io()`
- Listens for `ai-cmd` events and inserts/edits text via `Word.run()`
- Implements basic error handling

## 4. Workflow
1. install the MCP server: `npm install`
2. STDIO from Codex CLI or fake master, use unix pipeline to provide input
3. Sideload the Add-in manifest in Word
4. Send `EditTask` requests via Codex CLI or another service, e.g. `{ content: '...' }`
5. The Add-in client receives edits in real time and applies them to the document

## 5. Extensibility
- Support additional `EditTask` types (tables, images, formatting)
- Add WebSocket authentication, logging, and error tracking

## 6. Debugging
- add argument --debug, dump detail and error in debug.log 
## 7. Test
- test.sh Unit test, Fake STDIO for MCP client and socket connection for office
- STDIO: use shell pipeline to provide input
- socket: generate a test javascript as a socket client
- before test, prepare package.json
- NEW: You can pipe custom MCP JSONL into test.sh, one JSON object per line (script auto-sends initialize + notifications/initialized)
  Example:
  echo '{"jsonrpc":"2.0","id":10,"method":"tools/call","params":{"name":"editTask","arguments":{"content":"PipeMsg","action":"insert","target":"selection"}}}' | ./test.sh
  Multiple lines:
  cat <<'EOF' | ./test.sh
  {"jsonrpc":"2.0","id":11,"method":"tools/call","params":{"name":"ping","arguments":{"message":"hello"}}}
  {"jsonrpc":"2.0","id":12,"method":"tools/call","params":{"name":"editTask","arguments":{"content":"FromSTDIN","action":"insert","target":"selection"}}}
  EOF
  The script extracts expected content from the first editTask; if none provided, the first ai-cmd event counts as success.
## 8. Doc
- README.md for use step-by-step stall and run
-- Codex MCP setting
-- Word add-in: Script Lab
--- Script: copy from task-pane.js, 
--- Libraries https://cdn.socket.io/4.7.5/socket.io.min.js

- Code / doc all english only
