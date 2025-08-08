---
# SPEC: Proxy Server & Office Add-in

# Version: 1.0.0  
# Date: 2025-08-07  
# Author: Your Name

## 1. Introduction
MCP_WORD build a MCP server (`server.js`) alongside an Office.js Word Add-in client to enable AI-driven document editing workflows.

## 2. Architecture Overview
```mermaid
flowchart LR
  CLI/AI --> MCP Server
  MCP Server --> Browser[Office.js Task Pane]
  Browser --> Word[Word Document]
```


## 3. Components

### 3.1 MCP Server (`server.js`)
- Stack:
  1. Claude MCP TypeScript SDK (@modelcontextprotocol/sdk/server/mcp.js stdio
  2. office add-in: socket.io

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
1. Start the proxy server: `npm install && npm start`
2. Sideload the Add-in manifest in Word
3. Send `EditTask` requests via Claude CLI or another service, e.g. `{ content: '...' }`
4. The Add-in client receives edits in real time and applies them to the document

## 5. Extensibility
- Support additional `EditTask` types (tables, images, formatting)
- Add WebSocket authentication, logging, and error tracking
