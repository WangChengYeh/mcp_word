---
# SPEC: Proxy Server & Office Add-in

Version: 1.0.0  
Date: 2025-08-07  
Author: Your Name

## 1. Introduction
MCP Word Add-in 透過一個 Node.js 代理伺服器 (server.js) 與 Office.js Word Add-in 客戶端 (public/) 協作，實現 AI 驅動的 Word 文件編輯流程。

## 2. Architecture Overview
```mermaid
flowchart LR
  CLI/AI --> Proxy[Proxy Server (server.js)]
  Proxy --> Browser[Office.js Task Pane]
  Browser --> Word[Word Document]
```

## 3. Components

### 3.1 Proxy Server (server.js)
- 技術棧：Node.js (ESM)、Express、Socket.io、archiver、body-parser  
- 功能：
  1. 靜態託管 public 資源 (`manifest.xml`, `taskpane.html`, `taskpane.js`)  
  2. POST `/mcp`：接收 CLI 或 AI Agent 發送的 JSON，並通過 WebSocket 廣播 `ai-cmd` 事件  
  3. GET `/download`：動態打包專案成 ZIP 供下載 (排除 `node_modules`、`.git`)  
- 啟動：`node server.js`，預設監聽 3000 埠  

### 3.2 Office Add-in (public/)
#### manifest.xml
- 定義 Add-in ID、版本、ProviderName、DisplayName、Description  
- Host: Document；Permissions: ReadWriteDocument  
- SourceLocation: `http://localhost:3000/taskpane.html`

#### taskpane.html
- 引入 Office.js、Socket.io 客戶端  
- 加載 `taskpane.js`，提供一個按鈕或自動啟動

#### taskpane.js
- `Office.onReady()` 檢測 Word Host  
- 使用 `io()` 建立 WebSocket 連線  
- 監聽 `ai-cmd` 事件，於 `Word.run()` 中插入或處理文字  
- 基本錯誤處理

## 4. 使用流程
1. 啟動 Proxy Server：`npm install && npm start`  
2. 在 Word 中側載 Add-in manifest  
3. CLI 或其他服務 POST `/mcp` 並帶入 `{ content: '...' }`  
4. Add-in 實時接收並插入 Word 文件

## 5. Extensibility
- 支援更多 EditTask 類型 (表格、圖片、格式)  
- 增加 WebSocket 認證、日誌與錯誤追蹤
