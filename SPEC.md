---
# SPEC: Proxy Server & Office Add-in

# Version: 1.0.0  
# Date: 2025-08-07  
# Author: Your Name

## 1. Introduction
The MCP Word Add-in uses a Node.js proxy server (`server.js`) alongside an Office.js Word Add-in client (`public/`) to enable AI-driven document editing workflows.

## 2. Architecture Overview
```mermaid
flowchart LR
  CLI/AI --> Proxy[Proxy Server (server.js)]
  Proxy --> Browser[Office.js Task Pane]
  Browser --> Word[Word Document]
```


## 3. Components

### 3.1 MCP Server (`server.js`)
- Stack: Node.js (ESM), Express (static hosting), Claude MCP SDK
- Responsibilities:
  1. Serve static resources (`manifest.xml`, `taskpane.html`, `taskpane.js`) via Express
  2. Initialize the MCP SDK and register an `EditTask` handler:
     - Receive edit requests from CLI or an AI agent and interact with the AI model via SDK
     - Return the edit results to the Office Add-in client
- Startup: `node server.js` (listens on port 3000 by default)

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
