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
    edits: Array<{
      action: string;      // e.g., "insert", "delete", "replace"
      location: string;    // e.g., "Paragraph[2]", "Table[0].Cell[1]"
      content?: string;    // for insert/replace actions
    }>;
  }
  ```

- **Error**: Throws an `Error` if CLI invocation fails or output is invalid JSON.

---

### 2.2 MCPServer (src/server/mcpServer.ts)

**Class**: `MCPServer`

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
