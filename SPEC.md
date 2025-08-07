---
# SPEC: MCP Server for AI-Driven Word Document Processing

Version: 1.0.0  
Date: 2025-08-07  
Author: **Your Name**

## 1. Introduction
MCP Server 接收 AI Agent 的自然語言提示，使用 `@microsoft/mcp-sdk` 驗證並轉譯為結構化的編輯任務，然後將任務傳遞到 Office.js Word Add-in，於 Word 客戶端中實際執行文件操作。

## 2. Architecture Overview
```mermaid
flowchart LR
  A[AI Agent] --> B[MCP Server]
  B --> C[Office.js Word Add-in]
  C --> D[Word Document]
```

## 3. Components & Interfaces

### 3.1 MCP Server (src/server/mcpServer.ts)
**Class**: `MCPServer`  
**Dependencies**:
- `@microsoft/mcp-sdk`（訊息協定、驗證）  
- `express`（HTTP 伺服器）

#### Endpoints

1. **POST** `/api/process`  
   - **Request**:
     ```json
     { "prompt": string }
     ```
   - **Behavior**:
     1. 使用 MCP SDK 驗證輸入。  
     2. 呼叫 AI Agent 產生 `EditTask[]`。  
   - **Response**:
     ```json
     { "tasks": EditTask[] }
     ```

2. **POST** `/api/execute`  
   - **Request**:
     ```json
     { "tasks": EditTask[] }
     ```
   - **Behavior**:
     1. 使用 MCP SDK 驗證 `tasks`。  
     2. 回傳原樣 `tasks` 給客戶端由 Add-in 執行。  
   - **Response**:
     ```json
     { "tasks": EditTask[] }
     ```

#### Error Handling
- 回傳 HTTP 4xx/5xx 並附：
  ```json
  { "error": string; "statusCode": number }
  ```
- 短暫性錯誤自動重試最多 3 次。

### 3.2 Data Schemas

#### EditTask
```ts
interface EditTask {
  action: 'insertText' | 'deleteText' | 'findAndReplace' | 'format' | 'addTable' | 'insertPicture';
  target?: string;    // 如搜尋字串或表格索引
  content?: string;   // 插入或替換用文字
  options?: any;      // 各動作專屬參數
}
```

#### AIResponse
```ts
interface AIResponse {
  tasks: EditTask[];
}
```

#### ErrorResponse
```ts
interface ErrorResponse {
  error: string;
  statusCode: number;
}
```

## 4. Office.js Word Add-in (office/taskpane.ts)

**Function**: `async function applyEdits(prompt: string)`

1. POST `/api/process` with `{ prompt }`  
2. 取得 `{ tasks }`  
3. POST `/api/execute` with `{ tasks }`  
4. 在 Word 客戶端透過 Office.js 逐條套用編輯：
   ```ts
   await Word.run(async (context) => {
     const body = context.document.body;
     for (const task of tasks) {
       switch (task.action) {
         case 'insertText':
           body.insertText(task.content, Word.InsertLocation.end);
           break;
         case 'deleteText':
           context.document.getSelection().clear();
           break;
         case 'findAndReplace':
           const results = body.search(task.target, { matchCase: task.options?.matchCase });
           results.load();
           await context.sync();
           results.items.forEach(item => {
             item.insertText(task.content, Word.InsertLocation.replace);
           });
           break;
         case 'format':
           const sel = context.document.getSelection();
           sel.load('font');
           await context.sync();
           if (task.options.bold) sel.font.bold = true;
           if (task.options.italic) sel.font.italic = true;
           await context.sync();
           break;
         case 'addTable':
           body.insertTable(task.options.rows, task.options.cols, Word.InsertLocation.end);
           break;
         case 'insertPicture':
           body.insertInlinePictureFromBase64(task.options.base64, Word.InsertLocation.end);
           break;
         // ...其他動作...
       }
     }
     await context.sync();
   });
   ```

#### Error Handling (Client)
- 使用 `try/catch` 包裹 `fetch`，並以 Office UI 顯示錯誤通知。

## 5. Extensibility
- 擴充 `EditTask.action` 以支援 Excel/PowerPoint，並在 Add-in 中實作對應處理。  
- 新增 WebSocket 或 Webhook 支援即時協作。
