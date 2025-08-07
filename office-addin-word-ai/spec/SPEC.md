# SPEC: AI-Driven Word Editing Pipeline

This document defines the prompt, tools, and interfaces for the pipeline that connects:

1. **Codex CLI**
2. **MCP Server**
3. **mcp-msoffice-interop-word** CLI tool
4. **Office.js Word Add-in**

---

## 1. Pipeline Overview

User invokes a prompt in the Office Add-in, which is forwarded to the Codex CLI. The AI-generated JSON travels via the MCP Server to the mcp-msoffice-interop-word CLI tool, which applies edits to the Word document. The modified document is then returned to the Office Add-in.

```
Office.js UI → CodexCLI → MCPServer → mcp-msoffice-interop-word → Word Document
```  

---

## 2. Components & Interfaces

### 2.1 CodexCLI (src/ai/codexCLI.ts)

**Class**: `CodexCLI`

**Method**: `execute(prompt: string, options?: object): Promise<AIResponse>`

- **Input**:
  - `prompt`: Natural language instruction for document editing
  - `options`: (optional) parameters such as temperature, max_tokens
- **Output**:
  ```ts
  interface AIResponse {
    edits: Array<{ action: string; location: string; content?: string }>;
  }
  ```

- **Behavior**:
  1. Generates an `AIResponse` by invoking the AI agent.
  2. Sends the resulting `AIResponse.edits` to the MCP server for processing:
     ```ts
     await fetch(`${MCP_SERVER_URL}/api/process`, {
       method: 'POST',
       headers: { 'Content-Type': 'application/json' },
       body: JSON.stringify({ edits: aiResponse.edits })
     });
     ```

- **Error**: Throws an `Error` if CLI invocation fails, output is invalid JSON, or the MCP server request fails.

---

### 2.2 MCPServer (src/server/mcpServer.ts)

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

### 2.3 mcp-msoffice-interop-word CLI Tool

**Repository**: https://github.com/mario-andreschak/mcp-msoffice-interop-word

**Invocation**:
```bash
mcp-msoffice-interop-word \
  --input <path/to/processedData.json> \
  --output <path/to/edited.docx>
```

**Input Format**: JSON file matching the `processedData` schema.

**Output**: A Word `.docx` file with applied edits.

---

### 2.3.1 WordService API (mcp-msoffice-interop-word/src/word/word-service.ts)

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
