---
 # SPEC: MCP Server for AI-Driven Office Task Processing
---

Version: 1.0.0  
Date: 2025-08-07  
Author: **Your Name**

## 1. Introduction
The MCP Server receives structured instructions from an AI agent and executes Office tasks—such as editing Word documents—via a CLI-based Office interop tool. This specification details the server’s architecture, HTTP endpoints, data schemas, and error handling.

## 2. Architecture Overview
```mermaid
flowchart LR
  A[AI Agent] --> B[MCP Server]
  B --> C[Office Interop CLI]
  C --> D[Office Document]
```

**Note:** This specification covers only the MCP Server component.

## 3. Components & Interfaces

### 3.1 MCP Server (src/server/mcpServer.ts)
**Class**: `MCPServer`  
**Dependencies**:
- `@microsoft/mcp-sdk` for protocol and validation  
- Node.js `child_process` for invoking the interop CLI

#### Endpoints
- **POST** `/api/process`
  - **Request**:
    ```json
    { "prompt": string }
    ```
  - **Behavior**:
    1. Invoke AI agent to generate `AIResponse` tasks.
    2. Return the tasks JSON.
  - **Response**:
    ```json
    { "tasks": AIResponse }
    ```

- **POST** `/api/execute`
  - **Request**:
    ```json
    { "tasks": AIResponse }
    ```
  - **Behavior**:
    1. Serialize `tasks` to `input.json`.
    2. Execute:
       ```bash
       mcp-msoffice-interop-office --input input.json --output output.docx
       ```
    3. Return path to `output.docx`.
  - **Response**:
    ```json
    { "docPath": string }
    ```


### 3.2 CLI Integration
Use the `mcp-msoffice-interop-office` CLI tool to perform the actual Office automation tasks.

#### Public Methods
- `getWordApplication(): Promise<WordApplication>`
- `getActiveDocument(): Promise<WordDocument>`
- `createDocument(): Promise<WordDocument>`
- `openDocument(path: string): Promise<WordDocument>`
- `saveDocument(): Promise<void>`
- `saveDocumentAs(path: string, format?: number): Promise<void>`
- `closeDocument(doc: WordDocument, saveChanges?: boolean): Promise<void>`
- `insertText(text: string): Promise<void>`
- `deleteText(unit: number, count: number): Promise<void>`
- `findAndReplace(findText: string, replaceText: string, options?: ReplaceOptions): Promise<boolean>`
- `toggleFormat(type: 'bold'|'italic'|'underline'): Promise<void>`
- `setParagraphStyle(options: ParagraphOptions): Promise<void>`
- `addTable(rows: number, cols: number): Promise<Word.Table>`
- `setTableCellText(tableIndex: number, row: number, col: number, text: string): Promise<void>`
- `insertPicture(path: string, options?: PictureOptions): Promise<void>`
...additional methods as required...

#### Data Types
```ts
interface ReplaceOptions { matchCase?: boolean; matchWholeWord?: boolean; replaceAll?: boolean; }
interface ParagraphOptions { alignment?: number; indent?: { left?: number; right?: number; firstLine?: number }; spacing?: { before?: number; after?: number; line?: number }; }
interface PictureOptions { link?: boolean; embed?: boolean; width?: number; height?: number; }
```

### 3.3 Office.js Word Add-in (office/taskpane.js)
The Add-in listens for edit instructions pushed by the MCP Server.

#### Endpoint: `/taskpane/receiveEdits`
- Registers a message handler to apply edits:
  ```ts
  Office.actions.associate('receiveEdits', async (args) => {
    const edits = args.data.edits as Array<{ action: string; target: string; content?: string }>;
    // Apply edits to the document using Office.js APIs
    for (const edit of edits) {
      // e.g., insertText, deleteText, replace logic
    }
    return { status: 'success' };
  });
  ```
Users do not directly invoke UI flows; MCP Server pushes edits via Office runtime messages.

## 4. Data Schemas

### 4.1 AIResponse
```ts
interface AIResponse {
  edits: Array<{ action: string; target: string; content?: string }>;
}
```

### 4.2 ErrorResponse
```ts
interface ErrorResponse {
  error: string;
  statusCode: number;
}
```

### 4.3 HTTP Error Response
```ts
interface ErrorResponse { error: string; statusCode: number; }
```

## 5. Error Handling
- All endpoints must return HTTP 4xx/5xx with an `ErrorResponse`.
- Retry transient failures (network, COM timeouts) up to 3 times.
- Log errors for diagnostics.

## 6. Extensibility
- Extend to other Office hosts (Excel, PowerPoint) by updating CLI wrapper and task schema.
- Support webhook subscriptions or WebSocket streams for real-time collaboration.
---

## 1. Pipeline Overview


```
Codex CLI → MCPServer → Office.js Word Add-in → Word Document
```

---

## 2. Components & Interfaces
### 2.1 MCPServer (src/server/mcpServer.ts)

**Class**: `MCPServer`

**Dependencies**:

- `@microsoft/mcp-sdk`: MCP TypeScript SDK for message handling, validation, and protocol support.
**Endpoints**:

1. **POST** `/api/process`  
   - **Request Body**:  
     ```json
     {
       "edits": AIResponse.edits
     }
     ```
   - **Response**:
     ```json
     {
       "processedData": any // enriched with document context if needed
     }
     ```

2. **POST** `/api/apply`  
   - **Request Body**:
     ```json
     {
       "processedData": any
     }
     ```
   - **Behavior**: Invokes `mcp-msoffice-interop-word` CLI tool under the hood:
     ```bash
     mcp-msoffice-interop-word --input processedData.json --output edited.docx
     ```
   - **Response**:
     ```json
     {
       "docPath": string; // path to edited document
     }
     ```

- **Error Handling**: Returns HTTP 4xx/5xx with `{ error: string }` on failures.

---

## 2.2 WordService API (mcp-msoffice-interop-word/src/word/word-service.ts)

透過 `WordService` 類別，可在 Node.js 環境中直接呼叫 mcp-msoffice-interop-word 工具的主要方法：

```ts
interface WordService {
  getWordApplication(): Promise<WordApplication>;
  getActiveDocument(): Promise<WordDocument>;
  createDocument(): Promise<WordDocument>;
  openDocument(filePath: string): Promise<WordDocument>;
  saveActiveDocument(): Promise<void>;
  saveActiveDocumentAs(filePath: string, fileFormat?: any): Promise<void>;
  closeDocument(doc: WordDocument, saveChanges?: any): Promise<void>;
  quitWord(): Promise<void>;
  insertText(text: string): Promise<void>;
  deleteText(count?: number, unit?: number): Promise<void>;
  findAndReplace(
    findText: string,
    replaceText: string,
    matchCase?: boolean,
    matchWholeWord?: boolean,
    replaceAll?: boolean
  ): Promise<boolean>;
  toggleBold(): Promise<void>;
  toggleItalic(): Promise<void>;
  toggleUnderline(underlineStyle?: number): Promise<void>;
  setParagraphAlignment(alignment: number): Promise<void>;
  setParagraphLeftIndent(indentPoints: number): Promise<void>;
  setParagraphRightIndent(indentPoints: number): Promise<void>;
  setParagraphFirstLineIndent(indentPoints: number): Promise<void>;
  setParagraphSpaceBefore(spacePoints: number): Promise<void>;
  setParagraphSpaceAfter(spacePoints: number): Promise<void>;
  setParagraphLineSpacing(lineSpacingRule: number, lineSpacingValue?: number): Promise<void>;
  addTable(
    numRows: number,
    numCols: number,
    defaultTableBehavior?: number,
    autoFitBehavior?: number
  ): Promise<any>;
  getTableCell(tableIndex: number, rowIndex: number, colIndex: number): Promise<any>;
  setTableCellText(tableIndex: number, rowIndex: number, colIndex: number, text: string): Promise<void>;
  insertTableRow(tableIndex: number, beforeRowIndex?: number): Promise<any>;
  insertTableColumn(tableIndex: number, beforeColIndex?: number): Promise<any>;
  applyTableAutoFormat(
    tableIndex: number,
    formatName: string | number,
    applyFormatting?: number
  ): Promise<void>;
  insertPicture(
    filePath: string,
    linkToFile?: boolean,
    saveWithDocument?: boolean
  ): Promise<any>;
  setInlinePictureSize(
    shapeIndex: number,
    heightPoints: number,
    widthPoints: number,
    lockAspectRatio?: boolean
  ): Promise<void>;
  getHeaderFooter(
    sectionIndex: number,
    headerFooterType: number,
    isHeader: boolean
  ): Promise<any>;
  setHeaderFooterText(
    sectionIndex: number,
    headerFooterType: number,
    isHeader: boolean,
    text: string
  ): Promise<void>;
  setPageMargins(
    topPoints: number,
    bottomPoints: number,
    leftPoints: number,
    rightPoints: number
  ): Promise<void>;
  setPageOrientation(orientation: number): Promise<void>;
  setPaperSize(paperSize: number): Promise<void>;
  moveCursorToStart(): Promise<void>;
  moveCursorToEnd(): Promise<void>;
  moveCursor(unit: number, count: number, extend?: boolean): Promise<void>;
  selectAll(): Promise<void>;
  selectParagraph(paragraphIndex: number): Promise<void>;
  collapseSelection(toStart?: boolean): Promise<void>;
  getSelectionText(): Promise<string>;
  getSelectionInfo(): Promise<{ text: string; start: number; end: number; isActive: boolean; type: number }>;
}
```

#### Method Descriptions

- **getWordApplication()**: Retrieves or creates the Word.Application instance and ensures it is visible for subsequent operations.
- **getActiveDocument()**: Returns the currently active Word document; throws an error if no document is open.
- **createDocument()**: Creates and returns a new blank Word document.
- **openDocument(filePath: string)**: Opens the specified file path in Word and returns the document instance; throws on failure.
- **saveActiveDocument()**: Saves the currently active document; throws an error if save fails.
- **saveActiveDocumentAs(filePath: string, fileFormat?: any)**: Saves the active document to the provided path and format.
- **closeDocument(doc: WordDocument, saveChanges?: any)**: Closes the given document, optionally saving changes.
- **quitWord()**: Quits the Word application and releases the COM object reference.
- **insertText(text: string)**: Inserts the given text at the current selection point.
- **deleteText(count?: number, unit?: number)**: Deletes text based on count and unit (character, word, etc.) at the selection.
- **findAndReplace(findText: string, replaceText: string, matchCase?: boolean, matchWholeWord?: boolean, replaceAll?: boolean)**: Finds and replaces text in the document; returns a boolean indicating success.
- **toggleBold()**: Toggles bold formatting on the current selection.
- **toggleItalic()**: Toggles italic formatting on the current selection.
- **toggleUnderline(underlineStyle?: number)**: Toggles underline style on the current selection.
- **setParagraphAlignment(alignment: number)**: Sets paragraph alignment (left, center, right, justify).
- **setParagraphLeftIndent(indentPoints: number)**: Sets left indent for the selected paragraphs.
- **setParagraphRightIndent(indentPoints: number)**: Sets right indent for the selected paragraphs.
- **setParagraphFirstLineIndent(indentPoints: number)**: Sets first-line indent or hanging indent for paragraphs.
- **setParagraphSpaceBefore(spacePoints: number)**: Sets spacing before paragraphs.
- **setParagraphSpaceAfter(spacePoints: number)**: Sets spacing after paragraphs.
- **setParagraphLineSpacing(lineSpacingRule: number, lineSpacingValue?: number)**: Sets line spacing rule and value for paragraphs.
- **addTable(numRows: number, numCols: number, defaultTableBehavior?: number, autoFitBehavior?: number)**: Inserts a table at the selection.
- **getTableCell(tableIndex: number, rowIndex: number, colIndex: number)**: Retrieves the specified cell from a table.
- **setTableCellText(tableIndex: number, rowIndex: number, colIndex: number, text: string)**: Sets text in the specified table cell.
- **insertTableRow(tableIndex: number, beforeRowIndex?: number)**: Inserts a new row into the table.
- **insertTableColumn(tableIndex: number, beforeColIndex?: number)**: Inserts a new column into the table.
- **applyTableAutoFormat(tableIndex: number, formatName: string | number, applyFormatting?: number)**: Applies an auto-format style to the table.
- **insertPicture(filePath: string, linkToFile?: boolean, saveWithDocument?: boolean)**: Inserts an inline picture at the selection.
- **setInlinePictureSize(shapeIndex: number, heightPoints: number, widthPoints: number, lockAspectRatio?: boolean)**: Resizes an inline shape with optional aspect ratio lock.
- **getHeaderFooter(sectionIndex: number, headerFooterType: number, isHeader: boolean)**: Retrieves a header or footer object for a section.
- **setHeaderFooterText(sectionIndex: number, headerFooterType: number, isHeader: boolean, text: string)**: Sets text in the specified header or footer.
- **setPageMargins(topPoints: number, bottomPoints: number, leftPoints: number, rightPoints: number)**: Sets page margins.
- **setPageOrientation(orientation: number)**: Sets page orientation (portrait or landscape).
- **setPaperSize(paperSize: number)**: Sets paper size (e.g., Letter, A4).
- **moveCursorToStart()**: Moves the cursor to the start of the document.
- **moveCursorToEnd()**: Moves the cursor to the end of the document.
- **moveCursor(unit: number, count: number, extend?: boolean)**: Moves the cursor by the specified unit and count, optionally extending the selection.
- **selectAll()**: Selects the entire document.
- **selectParagraph(paragraphIndex: number)**: Selects the paragraph at the given index.
- **collapseSelection(toStart?: boolean)**: Collapses the selection to its start or end.
- **getSelectionText()**: Returns the text of the current selection.
- **getSelectionInfo()**: Returns selection metadata including text, start/end positions, active state, and type.

---

### 2.4 Office.js Word Add-in (office/taskpane.js)

**Function**: `async function applyEdits(prompt: string)`

- **UI Flow**:
  1. User enters instruction in task pane.
  2. `applyEdits` calls Codex CLI via MCP Server:
     ```ts
     const aiResponse = await fetch('/api/process', { method: 'POST', body: JSON.stringify({ prompt }) });
     const { processedData } = await aiResponse.json();
     ```
  3. Sends `processedData` to `/api/apply`:
     ```ts
     const applyRes = await fetch('/api/apply', { method: 'POST', body: JSON.stringify({ processedData }) });
     const { docPath } = await applyRes.json();
     Office.context.document.openAsync(docPath);
     ```

- **Error Handling**: Displays errors in UI toast on fetch or processing failures.

---

## 3. Data Schemas

### 3.1 Prompt Format
- Plain text natural language in user’s language.

### 3.2 AIResponse
```ts
interface AIResponse {
  edits: Array<{ action: string; location: string; content?: string }>;
}
```

### 3.3 ProcessedData
- Schema extends `AIResponse` with document context (e.g., full paragraphs, tables).

---

## 4. Error Handling

- **CodexCLI**: Throws on CLI errors or invalid JSON.
- **MCPServer**: Returns appropriate HTTP status codes with error message.
- **Office.js**: Shows toast notifications for network or processing errors.

---

## 5. Extensibility

- Support additional file formats (Excel, PowerPoint) by extending `mcp-msoffice-interop` tool and update endpoints.
- Add webhook/event-based flows for collaborative editing.
