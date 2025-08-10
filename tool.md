# MCP Word Tool Specification

Defines the MCP server tool interface and the canonical payloads for common Microsoft Word operations, along with explicit mappings to Office.js (Word JavaScript API). Covered operations: insert text, get selection, search, replace, insert picture, table operations, and style application.


## MCP Tool Definition

- Tool name: `mcp_word__editTask` (event: `ai-cmd`)
- Purpose: carry a structured Word operation via `meta` while keeping `action/content/target` for simple fallbacks.

Parameters:
- `action` (optional): `insert | replace | append` — legacy/simple text helper. For full control use `meta`.
- `content` (required): short human-readable summary for UI/logging.
- `meta` (optional but recommended): JSON string as defined in Common Envelope below.
- `target` (optional): `cursor | selection | document` — for simple text ops; otherwise specify scope in `meta.args`.
- `taskId` (optional): caller-provided correlation id.

Response (suggested): JSON with `{ ok: boolean, op: string, data?: object, diagnostics?: { level: 'info'|'warn'|'error', msg: string }[] }`.


## Common Envelope (meta)

- Event: `ai-cmd`
- Payload (embedded in `meta` as a JSON string):
  ```json
  {
    "type": "word.op",
    "op": "<operation>",
    "args": { /* operation-specific */ },
    "version": "1.0"
  }
  ```
- Suggested response (returned by the add-in):
  ```json
  {
    "ok": true,
    "op": "<operation>",
    "data": { /* result-specific */ },
    "diagnostics": [ { "level": "info|warn|error", "msg": "..." } ]
  }
  ```

Conventions and types:
- Encoding: UTF‑8 text.
- Colors: `#RRGGBB` or `rgba(r,g,b,a)`.
- Size unit: points (`pt`) unless otherwise stated; bare numbers imply `pt`.
- Positioning & scope:
  - `scope`: `document | selection | rangeId:<id>`.
  - `location`: `start | end | before | after | replace` (directly maps to `Word.InsertLocation`).
  - Aliases (accepted for compatibility, mapped internally): `append -> end`, `prepend -> start`, `insert -> replace` (when selection is collapsed, provider may map to `start`).
  - `rangeId`: opaque reusable id managed by the add-in via `context.trackedObjects`.
- Suggested error codes: `E_INVALID_ARG`, `E_NOT_FOUND`, `E_UNSUPPORTED`, `E_PERMISSION`, `E_RUNTIME`, `E_TIMEOUT`.

Range identity management (add-in guidance): when returning a `rangeId`, call `range.track()` and `context.trackedObjects.add(range)`, store it in an internal map keyed by a generated id, and return that id. On subsequent calls with `rangeId`, look up the tracked object; periodically untrack when no longer needed.


## Meta JSON Quick Reference (for AI authors)

Use these rules and templates to author `meta` JSON confidently.

- Required keys: `type` = `"word.op"`, `op` = one of the operations, `version` (e.g., `"1.0"`), and `args` object.
- Scope defaults: omit `scope` to use `selection`. Use `"document"` for whole doc; use `"rangeId:<id>"` to target a saved range.
- Location defaults: prefer `location`. If not set, providers treat text insert as `replace` at selection; aliases: `append -> end`, `prepend -> start`.
- Keep it minimal: only include arguments you actually need; defaults cover the rest.

Canonical template:
```json
{
  "type": "word.op",
  "op": "<insertText|getSelection|search|replace|insertPicture|table.*|applyStyle>",
  "version": "1.0",
  "args": { /* see per-op sections */ }
}
```

Common field aliases the provider accepts:
- `where` -> `location`
- `insert` (when selection collapsed) -> `location: "replace"` (provider MAY map)
- `append` -> `location: "end"`, `prepend` -> `location: "start"`
- `tableId:<id>` or `rangeId:<id>` can be used in `tableRef`

Validation guidance (what to expect on errors):
- Missing required field: `{ ok:false, code:"E_INVALID_ARG", diagnostics:[{"msg":"<field> required"}] }`
- Unsupported features (e.g., true regex search, picture wrap types): `{ ok:false, code:"E_UNSUPPORTED" }`
- Not found (bad `rangeId`/`tableRef`): `{ ok:false, code:"E_NOT_FOUND" }`



## Operation Index

- insertText: insert or replace text
- getSelection: get current selection text and range
- search: find text/style matches
- replace: replace text (first/all)
- insertPicture: insert image from URL or base64
- table.*: create and modify tables
- applyStyle: apply character/paragraph styles
- listStyles: list available named styles (paragraph/character/table)


## insertText

Purpose: insert or replace text at cursor, selection, document endpoints, or a saved range.

Args:
- `text` (string, required)
- `scope` (`document | selection | rangeId:<id>`, default `selection`)
- `location` (`start | end | before | after | replace`, default `replace`)
- `newParagraph` (boolean, default `false`)
- `keepFormatting` (boolean, default `true`)

Returns: `{ rangeId: string, length: number }`

Example `meta`:
```json
{
  "type": "word.op",
  "op": "insertText",
  "version": "1.0",
  "args": {
    "text": "Hello, Word!",
    "scope": "selection",
    "location": "replace",
    "newParagraph": false
  }
}
```

Office.js mapping:
- Resolve `targetRange` by `scope`:
  - `selection` -> `const target = context.document.getSelection();`
  - `document` -> `const target = context.document.body;` (use `insertText` on `Body` with `location`)
  - `rangeId:<id>` -> look up tracked `Word.Range` by id.
- Insert:
  - If `newParagraph` and `location` in `start|end|replace`, use `insertParagraph(text, location)`; else `insertText(text, location)`.
  - `location` maps 1:1 to `Word.InsertLocation` enum (`Start|End|Before|After|Replace`).
- Example:
  ```ts
  await Word.run(async (context) => {
    const sel = context.document.getSelection();
    sel.insertText("Hello, Word!", Word.InsertLocation.replace);
    await context.sync();
  });
  ```

Authoring tips (NL → meta):
- Say: "Insert 'Hello' at the start of the document" →
  {"type":"word.op","op":"insertText","version":"1.0","args":{"text":"Hello","scope":"document","location":"start"}}
- Say: "Replace current selection with 'OK'" →
  {"type":"word.op","op":"insertText","version":"1.0","args":{"text":"OK","location":"replace"}}
- Say: "Append a new paragraph 'Thanks'" →
  {"type":"word.op","op":"insertText","version":"1.0","args":{"text":"Thanks","scope":"document","location":"end","newParagraph":true}}


## getSelection

Purpose: return the current selection’s plain text and range information.

Returns:
- `text` (string)
- `rangeId` (string)
- `start` (number, optional document-relative)
- `end` (number, optional)

Example `meta`:
```json
{ "type": "word.op", "op": "getSelection", "version": "1.0", "args": {} }
```

Office.js mapping:
- `const sel = context.document.getSelection(); sel.load(["text", "start", "end"]); await context.sync();`
- Return `{ text: sel.text, rangeId: <tracked id>, start?, end? }`. Track the range and return an id.

Authoring tip:
- Say: "Get current selection" → {"type":"word.op","op":"getSelection","version":"1.0","args":{}}


## search

Purpose: search within a scope for text or by style-like constraints.

Args:
- `query` (string; note: Office.js supports wildcards, not general regex)
- `scope` (`document | selection | rangeId:<id>`, default `document`)
- `useRegex` (boolean, default `false`)
- `matchCase` (boolean, default `false`)
- `matchWholeWord` (boolean, default `false`)
- `matchPrefix` (boolean, default `false`)
- `maxResults` (number, default `100`)

Returns:
- `results`: `[ { rangeId, text, context?: string, start?: number, end?: number } ]`

Example:
```json
{
  "type": "word.op",
  "op": "search",
  "version": "1.0",
  "args": {
    "query": "Invoice",
    "scope": "document",
    "matchWholeWord": true,
    "maxResults": 50
  }
}
```

Office.js mapping and notes:
- Use `context.document.body.search(query, { matchCase, matchWholeWord, matchPrefix, matchWildcards })`.
- `useRegex`: not natively supported. Providers MAY translate simple regexes to Word wildcards or return `{ ok: false, code: "E_UNSUPPORTED" }`.
- Load results: `results.load(["text"]); await context.sync();` Then track each `Range` and return `rangeId`s.

Authoring tips (NL → meta):
- Say: "Find all 'Invoice' (whole word)" →
  {"type":"word.op","op":"search","version":"1.0","args":{"query":"Invoice","matchWholeWord":true}}
- Say: "Case-sensitive search for 'Total' in selection" →
  {"type":"word.op","op":"search","version":"1.0","args":{"query":"Total","scope":"selection","matchCase":true}}


## replace

Purpose: conditional replacement based on a query or an explicit target range.

Args:
- `target` (`document | selection | rangeId:<id> | searchQuery`)
- `query` (string; required when `target=searchQuery`)
- `useRegex`, `matchCase`, `matchWholeWord`, `matchPrefix` (same as search)
- `replaceWith` (string, required)
- `mode` (`replaceFirst | replaceAll`, default `replaceAll`)

Returns: `{ replaced: number }`

Example:
```json
{
  "type": "word.op",
  "op": "replace",
  "version": "1.0",
  "args": {
    "target": "searchQuery",
    "query": "2024",
    "matchWholeWord": true,
    "replaceWith": "2025",
    "mode": "replaceAll"
  }
}
```

Office.js mapping:
- If `target=searchQuery`, first run search (see above) and collect ranges; otherwise resolve `targetRange` by `scope`.
- For each range: `range.insertText(replaceWith, Word.InsertLocation.replace)`.
- Count successful replacements; return `{ replaced }`.

Authoring tips (NL → meta):
- Say: "Replace all 2024 with 2025 in the document" →
  {"type":"word.op","op":"replace","version":"1.0","args":{"target":"searchQuery","query":"2024","replaceWith":"2025","mode":"replaceAll"}}
- Say: "Replace the first match of 'foo' with 'bar' in the selection" →
  {"type":"word.op","op":"replace","version":"1.0","args":{"target":"searchQuery","query":"foo","replaceWith":"bar","mode":"replaceFirst","scope":"selection"}}


## insertPicture

Purpose: insert an image from URL or base64.

Args:
- `source` (`url | base64`)
- `data` (string; URL or base64)
- `scope` (`document | selection | rangeId:<id>`, default `selection`)
- `location` (`start | end | before | after | replace`, default `replace`)
- `width` (number, pt; optional)
- `height` (number, pt; optional)
- `lockAspectRatio` (boolean, default `true`)
- `altText` (string; optional)
- `wrapType` (`inline | square | tight | behind | inFront`, default `inline`)

Returns: `{ shapeId?: string, rangeId: string }`

Example:
```json
{
  "type": "word.op",
  "op": "insertPicture",
  "version": "1.0",
  "args": {
    "source": "url",
    "data": "https://example.com/logo.png",
    "location": "replace",
    "width": 120,
    "lockAspectRatio": true,
    "altText": "Company Logo"
  }
}
```

Office.js mapping and notes:
- Word API inserts inline pictures from base64 only. Providers MUST convert `url` to base64 prior to insertion or return an error if network is unavailable.
- At selection/range: `range.insertInlinePictureFromBase64(base64, location)`; at document: `context.document.body.insertInlinePictureFromBase64(base64, location)`.
- Sizing: after insertion, set `pic.width`/`pic.height` if provided; if one dimension provided, maintain aspect ratio when `lockAspectRatio=true`.
- Wrapping: Office.js inline pictures do not support floating wrap via this API; treat `wrapType` other than `inline` as `{ ok:false, code:"E_UNSUPPORTED" }` unless provider supports shapes.

Authoring tips (NL → meta):
- Say: "Insert logo from URL at cursor, width 120pt" →
  {"type":"word.op","op":"insertPicture","version":"1.0","args":{"source":"url","data":"https://example.com/logo.png","location":"replace","width":120}}
- Say: "Insert base64 image at end of document" →
  {"type":"word.op","op":"insertPicture","version":"1.0","args":{"source":"base64","data":"<BASE64>","scope":"document","location":"end"}}


## table.* (table operations)

Use `op: "table.<subop>"`. Common sub-operations:

1) `table.create`
- Args:
  - `rows` (number, required)
  - `cols` (number, required)
  - `scope` / `where` (same as insertText; default insert at selection)
  - `data` (string[][], optional initial cell values)
  - `header` (boolean; treat first row as header)
- Returns: `{ tableId: string, rangeId: string }`

2) `table.insertRows`
- Args: `{ tableRef: "tableId:<id> | rangeId:<id>", at: number, count: number }`

3) `table.insertColumns`
- Args: `{ tableRef, at: number, count: number }`

4) `table.deleteRows` / `table.deleteColumns`
- Args: `{ tableRef, indexes: number[] }`

5) `table.setCellText`
- Args: `{ tableRef, row: number, col: number, text: string }`

6) `table.mergeCells`
- Args: `{ tableRef, startRow, startCol, rowSpan, colSpan }`

7) `table.applyStyle`
- Args:
  - `tableRef`
  - `style`: `BuiltinName | { bandedRows?: boolean, bandedColumns?: boolean, firstRow?: boolean, lastRow?: boolean, firstColumn?: boolean, lastColumn?: boolean }`

Example (create 3x3 and fill one cell):
```json
{
  "type": "word.op",
  "op": "table.create",
  "version": "1.0",
  "args": { "rows": 3, "cols": 3, "header": true }
}
```
```json
{
  "type": "word.op",
  "op": "table.setCellText",
  "version": "1.0",
  "args": { "tableRef": "tableId:...", "row": 0, "col": 0, "text": "Title" }
}
```

Office.js mapping and notes:
- Create: `const table = context.document.body.insertTable(rows, cols, Word.InsertLocation.start|end|before|after, data?);` or `range.insertTable(...)` when inserting relative to a range. `table.load("id"); await context.sync();` Track `table` and return `tableId`.
- Header and banding: set `table.headerRow = true/false`; also `table.bandedRows`, `table.bandedColumns`, `table.firstColumn`, `table.lastColumn`, `table.totalRow` as needed.
- Insert rows/columns: `table.addRows(Word.InsertLocation.after|before, count)` / `table.addColumns(...)` relative to `at` index; for index-based, use row/column collections: `const row = table.rows.getItemAt(at); row.insertRows(...)` where supported. Otherwise rebuild via copy when API is limited.
- Delete rows/columns: `table.rows.getItemAt(i).delete()` / `table.columns.getItemAt(i).delete()`.
- Set cell text: `table.getCell(row, col).insertText(text, Word.InsertLocation.replace)`.
- Merge cells: `table.getCell(startRow, startCol).merge(table.getCell(startRow+rowSpan-1, startCol+colSpan-1));` Note: merging support depends on requirement set.
- Style: apply built-in names via `table.style = "TableGridLight"` (or other names). Banding flags as above.

Authoring tips (NL → meta):
- Say: "Create a 3x3 table with header at cursor" →
  {"type":"word.op","op":"table.create","version":"1.0","args":{"rows":3,"cols":3,"header":true}}
- Say: "Set row 0 col 0 to 'Title' for tableId:123" →
  {"type":"word.op","op":"table.setCellText","version":"1.0","args":{"tableRef":"tableId:123","row":0,"col":0,"text":"Title"}}
- Say: "Insert 2 rows after row 1 for tableId:123" →
  {"type":"word.op","op":"table.insertRows","version":"1.0","args":{"tableRef":"tableId:123","at":1,"count":2}}
- Say: "Merge a 2x2 area from (1,1) for tableId:123" →
  {"type":"word.op","op":"table.mergeCells","version":"1.0","args":{"tableRef":"tableId:123","startRow":1,"startCol":1,"rowSpan":2,"colSpan":2}}


## applyStyle

Purpose: apply character and paragraph styles to a target range.

Args:
- `scope` (`selection | document | rangeId:<id>`, default `selection`)
- `char`:
  - `bold?` (boolean)
  - `italic?` (boolean)
  - `underline?` (`none | single | double`)
  - `fontName?` (string)
  - `fontSize?` (number, pt)
  - `color?` (string text color)
  - `highlight?` (string background color)
- `para`:
  - `alignment?` (`left | center | right | justify`)
  - `lineSpacing?` (number, e.g., 1.15)
  - `spaceBefore?` (number, pt)
  - `spaceAfter?` (number, pt)
  - `list?` (`none | bullet | number`)
- `namedStyle` (string; e.g., `Normal`, `Heading 1`, `Title`)
- `clearOtherStyles` (boolean, default `false`)

Returns: `{ rangeId: string }`

Example:
```json
{
  "type": "word.op",
  "op": "applyStyle",
  "version": "1.0",
  "args": {
    "scope": "selection",
    "char": { "bold": true, "fontSize": 12, "color": "#333333" },
    "para": { "alignment": "justify", "lineSpacing": 1.15 },
    "namedStyle": "Normal"
  }
}
```

Office.js mapping:
- Resolve target range; for `document` use `context.document.body.getRange()`.
- Named style: `range.style = "Heading 1"` (where supported) or apply paragraph style via `range.paragraphs.items.forEach(p => p.style = ...)`.
- Character formatting: `range.font.bold/italic/underline/fontSize/color/highlightColor`.
- Paragraph formatting: `range.paragraphFormat.alignment/lineSpacing/spaceBefore/spaceAfter`.
- Lists:
  - `bullet`: `range.paragraphs.load(); await context.sync(); range.paragraphs.items.forEach(p => p.startNewList());` then set as bullet style if available.
  - `number`: same as above but set numbering. If not supported, return `E_UNSUPPORTED`.
- Clearing: if `clearOtherStyles=true`, remove direct formatting by reapplying `Normal` then re-apply specified overrides.

Authoring tips (NL → meta):
- Say: "Make selection bold, 12pt, dark gray, justified" →
  {"type":"word.op","op":"applyStyle","version":"1.0","args":{"char":{"bold":true,"fontSize":12,"color":"#333333"},"para":{"alignment":"justify"}}}
- Say: "Apply Heading 1 to selection" →
  {"type":"word.op","op":"applyStyle","version":"1.0","args":{"namedStyle":"Heading 1"}}


## Mapping to `mcp_word__editTask`

- `action`: `insert | replace | append` for basic text; providers MAY ignore when `meta` is supplied.
- `content`: short description (e.g., "Replace selection with 'Hello'").
- `meta`: JSON string per this spec (authoritative for operation and args).
- `target`: kept for simple scenarios; for full fidelity, specify `scope` and `location` in `meta.args`.


## Best Practices

- Small steps: split complex flows into multiple `ai-cmd` calls; pass back `rangeId` / `tableId` for chaining.
- Stable targeting: prefer `rangeId` to avoid cursor movement races.
- Graceful empty results: `search` with no hits should return `{ ok: true, data: { results: [] } }`.
- External images: if a URL cannot be fetched, return `{ ok: false, code: "E_RUNTIME" }` with diagnostics.
- Versioning: include `version` in payloads; add-in may use it for compatibility.


## listStyles

Purpose: provide a list of style names that can be used with `applyStyle` and `table.applyStyle`. Use this to power style pickers or validate style names.

Args:
- `category` (`paragraph | character | table | all`, default `all`)
- `query` (string; optional name substring filter)
- `builtInOnly` (boolean; default `true`)
- `includeLocalized` (boolean; default `true`) — when available, include localized display names in addition to canonical names.
- `max` (number; optional cap on returned items)

Returns:
```json
{
  "paragraphStyles": [ { "name": "Normal", "builtIn": true }, ... ],
  "characterStyles": [ { "name": "Emphasis", "builtIn": true }, ... ],
  "tableStyles": [ { "name": "Table Grid", "builtIn": true }, ... ]
}
```

Example `meta`:
```json
{ "type": "word.op", "op": "listStyles", "version": "1.0", "args": { "category": "paragraph" } }
```

Office.js mapping and notes:
- Word JavaScript API does not currently expose a direct enumeration of all styles. Providers SHOULD:
  - Return a curated list of common built-ins (see `task.md`).
  - Optionally verify availability by attempting to apply the style to a temporary paragraph or table created at `End`, then removing it.
  - Optionally augment with tenant/template-specific styles from configuration.
- For localized clients, providers MAY return both canonical English `name` and a `localizedName` field when known.
- If `query` is present, filter case-insensitively on `name` and `localizedName`.

Suggested response shape (extended):
```json
{
  "ok": true,
  "op": "listStyles",
  "data": {
    "paragraphStyles": [ { "name": "Normal", "localizedName": "Normal", "builtIn": true } ],
    "characterStyles": [ { "name": "Emphasis", "builtIn": true } ],
    "tableStyles": [ { "name": "Table Grid Light", "builtIn": true } ]
  }
}
```


## Office.js API Used

- Core batching
  - `Word.run(handler)`
  - `context.sync()`

- Document, body, selection
  - `context.document.getSelection()`
  - `context.document.body.getRange()`
  - `context.document.body.insertText(text, Word.InsertLocation)`
  - `context.document.body.insertParagraph(text, Word.InsertLocation)`
  - `context.document.body.search(query, options)`
  - `context.document.body.insertInlinePictureFromBase64(base64, Word.InsertLocation)`
  - `context.document.body.insertTable(rows, cols, Word.InsertLocation, data?)`
  - `context.document.body.getOoxml()` (optional: advanced providers may inspect OOXML)

- Range
  - `range.insertText(text, Word.InsertLocation)`
  - `range.insertParagraph(text, Word.InsertLocation)`
  - `range.search(query, options)`
  - `range.insertInlinePictureFromBase64(base64, Word.InsertLocation)`
  - `range.insertTable(rows, cols, Word.InsertLocation, data?)`
  - `range.getOoxml()` (optional)
  - `range.load(props)`
  - `range.track()` / `range.untrack()`
  - `range.style` (property)
  - `range.font.bold | italic | underline | fontSize | color | highlightColor` (properties)
  - `range.paragraphFormat.alignment | lineSpacing | spaceBefore | spaceAfter` (properties)
  - `range.paragraphs` (collection)

- Tables
  - `table.addRows(Word.InsertLocation, count)`
  - `table.addColumns(Word.InsertLocation, count)`
  - `table.rows.getItemAt(index)` / `table.columns.getItemAt(index)`
  - `table.rows.getItemAt(i).delete()` / `table.columns.getItemAt(i).delete()`
  - `table.getCell(row, col)`
  - `tableCell.insertText(text, Word.InsertLocation)`
  - `tableCell.merge(targetCell)`
  - `table.headerRow` (property)
  - `table.bandedRows | bandedColumns | firstColumn | lastColumn | totalRow` (properties)
  - `table.style` (property)

- Enums and options
  - `Word.InsertLocation` (`Start | End | Before | After | Replace`)
  - `Word.SearchOptions` fields: `matchCase`, `matchWholeWord`, `matchPrefix`, `matchWildcards`


## Types (suggested TypeScript for validation)

```ts
export type Scope = "document" | "selection" | `rangeId:${string}`;
export type Location = "start" | "end" | "before" | "after" | "replace";

export interface Envelope<T = unknown> {
  type: "word.op";
  op: "insertText" | "getSelection" | "search" | "replace" | "insertPicture" | `table.${string}` | "applyStyle";
  args: T;
  version: "1.0" | string;
}

export interface InsertTextArgs {
  text: string;
  scope?: Scope;
  location?: Location;
  newParagraph?: boolean;
  keepFormatting?: boolean;
}

export interface SearchArgs {
  query: string;
  scope?: Scope;
  useRegex?: boolean;
  matchCase?: boolean;
  matchWholeWord?: boolean;
  matchPrefix?: boolean;
  maxResults?: number;
}

export interface ReplaceArgs extends Omit<SearchArgs, "scope"> {
  target: Scope | "searchQuery";
  replaceWith: string;
  mode?: "replaceFirst" | "replaceAll";
}

export interface InsertPictureArgs {
  source: "url" | "base64";
  data: string;
  scope?: Scope;
  location?: Location;
  width?: number;
  height?: number;
  lockAspectRatio?: boolean;
  altText?: string;
  wrapType?: "inline" | "square" | "tight" | "behind" | "inFront";
}

export interface TableCreateArgs {
  rows: number; cols: number; scope?: Scope; location?: Location; data?: string[][]; header?: boolean;
}

export interface ListStylesArgs {
  category?: "paragraph" | "character" | "table" | "all";
  query?: string;
  builtInOnly?: boolean;
  includeLocalized?: boolean;
  max?: number;
}
```


## Minimal Integration Example (caller side)

- `content`: human text like "Replace current selection with 'Hello, Word!'".
- `meta`:
```json
{
  "type": "word.op",
  "op": "insertText",
  "version": "1.0",
  "args": { "text": "Hello, Word!", "scope": "selection", "location": "replace" }
}
```
- `target`: `selection`

Add-in handling outline:
1) Parse the `meta` JSON.
2) Dispatch by `op` to the appropriate handler.
3) Execute Word API (Office.js) calls.
4) Return a normalized response with `ok`, `data`, and optional `diagnostics`.
