# MCP Word Tool Specification

Defines the MCP server tool interface and the canonical payloads for common Microsoft Word operations, along with explicit mappings to Office.js (Word JavaScript API). Covered operations: insert text, get selection, search, replace, insert picture, table operations, and style application.


## Quick Map

- MCP tools → Socket.IO events (emit)
  - insertText → `word:insertText`
  - getSelection → `word:getSelection`
  - search → `word:search`
  - replace → `word:replace`
  - insertPicture → `word:insertPicture`
  - table.create → `word:table.create`
  - table.insertRows → `word:table.insertRows`
  - table.insertColumns → `word:table.insertColumns`
  - table.deleteRows → `word:table.deleteRows`
  - table.deleteColumns → `word:table.deleteColumns`
  - table.setCellText → `word:table.setCellText`
  - table.mergeCells → `word:table.mergeCells`
  - table.applyStyle → `word:table.applyStyle`
  - applyStyle → `word:applyStyle`
  - listStyles → `word:listStyles`
  - paragraphs.index → `word:paragraphs.index`


## Conventions

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



## MCP Tool Index

- insertText
- getSelection
- search
- replace
- insertPicture
- table.create
- table.insertRows
- table.insertColumns
- table.deleteRows
- table.deleteColumns
- table.setCellText
- table.mergeCells
- table.applyStyle
- applyStyle
- listStyles
- paragraphs.index

## MCP Tools (detailed)


<a id="op-insertText"></a>
## insertText

Purpose: insert or replace text at cursor, selection, document endpoints, or a saved range.

Socket.IO event: `word:insertText`

Args:
- `text` (string, required)
- `scope` (`document | selection | rangeId:<id>`, default `selection`)
- `location` (`start | end | before | after | replace`, default `replace`)
- `newParagraph` (boolean, default `false`)
- `keepFormatting` (boolean, default `true`) — provider MAY ignore; Word `insertText` inherits formatting by design.

Returns: `{ rangeId: string, length: number }`

Example args (MCP tool):
```json
{
  "text": "Hello, Word!",
  "scope": "selection",
  "location": "replace",
  "newParagraph": false
}
```

Office.js mapping:
- Resolve `targetRange` by `scope`:
  - `selection` -> `const target = context.document.getSelection();`
  - `document` -> `const target = context.document.body;` (Body supports `Start`/`End`; `Replace/Before/After` are not applicable)
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

Authoring tips (NL → args):
- Say: "Insert 'Hello' at the start of the document" →
  {"text":"Hello","scope":"document","location":"start"}
- Say: "Replace current selection with 'OK'" →
  {"text":"OK","location":"replace"}
- Say: "Append a new paragraph 'Thanks'" →
  {"text":"Thanks","scope":"document","location":"end","newParagraph":true}


<a id="op-getSelection"></a>
## getSelection

Purpose: return the current selection’s plain text and range information.

Socket.IO event: `word:getSelection`

Returns:
- `text` (string)
- `rangeId` (string)
- `start` (number, optional document-relative)
- `end` (number, optional)

Example args (MCP tool): `{}`

Office.js mapping:
- `const sel = context.document.getSelection(); sel.load(["text", "start", "end"]); await context.sync();`
- Return `{ text: sel.text, rangeId: <tracked id>, start?, end? }`. Track the range and return an id.

Authoring tip:
- Say: "Get current selection" → `{}`


<a id="op-search"></a>
## search

Purpose: search within a scope for text or by style-like constraints.

Socket.IO event: `word:search`

Args:
- `query` (string; note: Office.js supports wildcards, not general regex)
- `scope` (`document | selection | rangeId:<id>`, default `document`)
- `useRegex` (boolean, default `false`)
- `matchCase` (boolean, default `false`)
- `matchWholeWord` (boolean, default `false`)
- `matchPrefix` (boolean, default `false`)
- `matchSuffix` (boolean, default `false`)
- `ignoreSpace` (boolean, default `false`)
- `ignorePunct` (boolean, default `false`)
- `maxResults` (number, default `100`)

Returns:
- `results`: `[ { rangeId, text, context?: string, start?: number, end?: number } ]`

Example args (MCP tool):
```json
{ "query": "Invoice", "scope": "document", "matchWholeWord": true, "maxResults": 50 }
```

Office.js mapping and notes:
- Use `context.document.body.search(query, { matchCase, matchWholeWord, matchPrefix, matchWildcards })`.
- `useRegex`: not natively supported. Providers MAY translate simple regexes to Word wildcards or return `{ ok: false, code: "E_UNSUPPORTED" }`.
- Load results: `results.load(["text"]); await context.sync();` Then track each `Range` and return `rangeId`s.
 - Word.SearchOptions also supports `matchSuffix`, `ignoreSpace`, `ignorePunct`.

Authoring tips (NL → args):
- Say: "Find all 'Invoice' (whole word)" →
  {"query":"Invoice","matchWholeWord":true}
- Say: "Case-sensitive search for 'Total' in selection" →
  {"query":"Total","scope":"selection","matchCase":true}


<a id="op-replace"></a>
## replace

Purpose: conditional replacement based on a query or an explicit target range.

Socket.IO event: `word:replace`

Args:
- `target` (`document | selection | rangeId:<id> | searchQuery`)
- `query` (string; required when `target=searchQuery`)
- `useRegex`, `matchCase`, `matchWholeWord`, `matchPrefix` (same as search)
- `replaceWith` (string, required)
- `mode` (`replaceFirst | replaceAll`, default `replaceAll`)

Returns: `{ replaced: number }`

Example args (MCP tool):
```json
{ "target": "searchQuery", "query": "2024", "matchWholeWord": true, "replaceWith": "2025", "mode": "replaceAll" }
```

Office.js mapping:
- If `target=searchQuery`, first run search (see above) and collect ranges; otherwise resolve `targetRange` by `scope`.
- For each range: `range.insertText(replaceWith, Word.InsertLocation.replace)`.
- Count successful replacements; return `{ replaced }`.

Authoring tips (NL → args):
- Say: "Replace all 2024 with 2025 in the document" →
  {"target":"searchQuery","query":"2024","replaceWith":"2025","mode":"replaceAll"}
- Say: "Replace the first match of 'foo' with 'bar' in the selection" →
  {"target":"searchQuery","query":"foo","replaceWith":"bar","mode":"replaceFirst","scope":"selection"}


<a id="op-insertPicture"></a>
## insertPicture

Purpose: insert an image from URL or base64.

Socket.IO event: `word:insertPicture`

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

Example args (MCP tool):
```json
{ "source": "url", "data": "https://example.com/logo.png", "location": "replace", "width": 120, "lockAspectRatio": true, "altText": "Company Logo" }
```

Office.js mapping and notes:
- Word API inserts inline pictures from base64 only. Providers MUST convert `url` to base64 prior to insertion or return an error if network is unavailable.
- At selection/range: `range.insertInlinePictureFromBase64(base64, location)`; at document: `context.document.body.insertInlinePictureFromBase64(base64, location)`.
- Sizing: after insertion, set `pic.width`/`pic.height` if provided; if one dimension provided, maintain aspect ratio when `lockAspectRatio=true`.
- Wrapping: Office.js inline pictures do not support floating wrap via this API; treat `wrapType` other than `inline` as `{ ok:false, code:"E_UNSUPPORTED" }` unless provider supports shapes.
 - Alt text: map `altText` to `inlinePicture.altTextDescription` (and optionally `altTextTitle`).

Authoring tips (NL → args):
- Say: "Insert logo from URL at cursor, width 120pt" →
  {"source":"url","data":"https://example.com/logo.png","location":"replace","width":120}
- Say: "Insert base64 image at end of document" →
  {"source":"base64","data":"<BASE64>","scope":"document","location":"end"}


<a id="op-table"></a>
## table.* (table operations)

Use `op: "table.<subop>"`. Common sub-operations:

Socket.IO events:
- `word:table.create`
- `word:table.insertRows`
- `word:table.insertColumns`
- `word:table.deleteRows`
- `word:table.deleteColumns`
- `word:table.setCellText`
- `word:table.mergeCells`
- `word:table.applyStyle`

1) `table.create`
- Args:
  - `rows` (number, required)
  - `cols` (number, required)
  - `scope` / `location` (same as insertText; default insert at selection)
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

Example args (create 3x3 and fill one cell):
```json
{ "rows": 3, "cols": 3, "header": true }
```
```json
{ "tableRef": "tableId:...", "row": 0, "col": 0, "text": "Title" }
```

Office.js mapping and notes:
- Create: `const table = context.document.body.insertTable(rows, cols, Word.InsertLocation.start|end|before|after, data?);` or `range.insertTable(...)` when inserting relative to a range. `table.load("id"); await context.sync();` Track `table` and return `tableId`.
- Header and banding: set `table.headerRow = true/false`; also `table.bandedRows`, `table.bandedColumns`, `table.firstColumn`, `table.lastColumn`, `table.totalRow` as needed.
- Insert rows/columns: `table.addRows(Word.InsertLocation.after|before, count)` / `table.addColumns(...)` relative to `at` index; for index-based, use row/column collections: `const row = table.rows.getItemAt(at); row.insertRows(...)` where supported. Otherwise rebuild via copy when API is limited.
- Delete rows/columns: `table.rows.getItemAt(i).delete()` / `table.columns.getItemAt(i).delete()`.
- Set cell text: `table.getCell(row, col).insertText(text, Word.InsertLocation.replace)`.
- Merge cells: `table.getCell(startRow, startCol).merge(table.getCell(startRow+rowSpan-1, startCol+colSpan-1));` Note: merging support depends on requirement set.
- Style: apply built-in names via `table.style = "TableGridLight"` (or other names). Banding flags as above.

Authoring tips (NL → args):
- Say: "Create a 3x3 table with header at cursor" →
  {"rows":3,"cols":3,"header":true}
- Say: "Set row 0 col 0 to 'Title' for tableId:123" →
  {"tableRef":"tableId:123","row":0,"col":0,"text":"Title"}
- Say: "Insert 2 rows after row 1 for tableId:123" →
  {"tableRef":"tableId:123","at":1,"count":2}
- Say: "Merge a 2x2 area from (1,1) for tableId:123" →
  {"tableRef":"tableId:123","startRow":1,"startCol":1,"rowSpan":2,"colSpan":2}


<a id="op-applyStyle"></a>
## applyStyle

Purpose: apply named styles (Heading, Quote, etc.) and/or direct formatting (font size, bold, alignment) to a target range.

Socket.IO event: `word:applyStyle`

Key principles
- Named styles set baseline formatting on paragraphs or character runs. Direct formatting then overrides the baseline.
- Recommended order: apply namedStyle first, then apply paragraph (`para`) and character (`char`) overrides.

Args:
- `scope` (`selection | document | rangeId:<id>`, default `selection`)
- `namedStyle` (string; e.g., `Normal`, `Heading 1`, `Title`) — paragraph or character style; provider chooses the appropriate target based on style kind when known.
- `char` (direct character formatting overrides):
  - `bold?` (boolean)
  - `italic?` (boolean)
  - `underline?` (`none | single | double`)
  - `strikeThrough?` (boolean)
  - `doubleStrikeThrough?` (boolean)
  - `allCaps?` (boolean)
  - `smallCaps?` (boolean)
  - `superscript?` (boolean)
  - `subscript?` (boolean)
  - `fontName?` (string)
  - `fontSize?` (number, pt)
  - `color?` (string; text color)
  - `highlight?` (string; text highlight color)
- `para` (direct paragraph formatting overrides):
  - `alignment?` (`left | center | right | justify`)
  - `lineSpacing?` (number, e.g., 1.15)
  - `spaceBefore?` (number, pt)
  - `spaceAfter?` (number, pt)
  - `leftIndent?` (number, pt)
  - `rightIndent?` (number, pt)
  - `firstLineIndent?` (number, pt)
  - `list?` (`none | bullet | number`)
- `precedence` (`styleThenOverrides | overridesThenStyle`, default `styleThenOverrides`) — controls application order if both are provided.
- `resetDirectFormatting` (boolean, default `false`) — when true, provider should clear existing direct formatting before applying settings (e.g., by reapplying `Normal` and then the desired `namedStyle`/overrides).

Returns: `{ rangeId: string }`

Example args (MCP tool):
```json
{
  "scope": "selection",
  "namedStyle": "Heading 1",
  "para": { "alignment": "justify", "lineSpacing": 1.15, "firstLineIndent": 18 },
  "char": { "bold": true, "fontSize": 14, "color": "#333333" },
  "precedence": "styleThenOverrides"
}
```

Office.js mapping:
- Resolve target range; for `document` use `context.document.body.getRange()`.
- Named style: set `range.style = "Heading 1"` (applies to paragraphs intersecting the range). For character styles, apply to the selection/range; provider may need to ensure the selection is text-only.
- Character formatting: map to `range.font` properties: `bold`, `italic`, `underline`, `strikeThrough`, `doubleStrikeThrough`, `allCaps`, `smallCaps`, `superscript`, `subscript`, `name`, `size`, `color`, `highlightColor`.
- Paragraph formatting: map to `range.paragraphFormat` properties: `alignment`, `lineSpacing`, `spaceBefore`, `spaceAfter`, `leftIndent`, `rightIndent`, `firstLineIndent`.
- Lists:
  - `bullet`: where supported, use `Paragraph.startNewList()` and set bullet type; otherwise apply named style "List Paragraph" as a fallback.
  - `number`: similar approach for numbered lists. If API not available, return `E_UNSUPPORTED`.
- Clearing: if `clearOtherStyles=true`, remove direct formatting by reapplying `Normal` then re-apply specified overrides.

Authoring tips (NL → args):
- Say: "Make selection bold, 12pt, dark gray, justified" →
  {"char":{"bold":true,"fontSize":12,"color":"#333333"},"para":{"alignment":"justify"}}
- Say: "Apply Heading 1 to selection" →
  {"namedStyle":"Heading 1"}


## Compatibility: Aggregate Tool (optional)

Some providers may prefer a single aggregate MCP tool (e.g., `editTask`) that emits one Socket.IO event carrying `{ op, args }` or a `meta` string with the same structure. When adopting this pattern:
- Event name: choose a single event like `editTask`, or forward the `op` directly as event name `word:<op>`.
- Envelope shape (JSON): `{ type: "word.op", op: "<name>", args: { ... }, version: "1.0" }`.
- Map each `<name>` to the per-tool sections above and execute the same Office.js calls.


## Best Practices

- Small steps: split complex flows into multiple Socket.IO emits (e.g., `word:insertText`, `word:applyStyle`); pass back `rangeId` / `tableId` for chaining.
- Stable targeting: prefer `rangeId` to avoid cursor movement races.
- Graceful empty results: `search` with no hits should return `{ ok: true, data: { results: [] } }`.
- External images: if a URL cannot be fetched, return `{ ok: false, code: "E_RUNTIME" }` with diagnostics.
- Versioning: include `version` in payloads; add-in may use it for compatibility.


<a id="op-listStyles"></a>
## listStyles

Purpose: provide a list of style names that can be used with `applyStyle` and `table.applyStyle`. Use this to power style pickers or validate style names.

Socket.IO event: `word:listStyles`

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

Example args (MCP tool): `{ "category": "paragraph" }`

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


<a id="op-paragraphs.index"></a>
## paragraphs.index

Purpose: enumerate paragraphs within a scope; optionally ensure each paragraph is wrapped in a content control and return stable identifiers for later targeting.

Socket.IO event: `word:paragraphs.index`

Args:
- scope (`document | selection | rangeId:<id>`, default `document`)
- ensureContentControls (boolean, default `true`) — when true, create a content control per paragraph if missing.
- tagPrefix (string, default `"para"`) — applied to each created/normalized content control tag, e.g., `para:0`, `para:1`, ...
- includeText (boolean, default `true`) — include plain text of each paragraph in response.
- includeEmpty (boolean, default `false`) — include empty paragraphs.
- max (number, optional) — cap the number of returned items.

Returns:
- `items`: `[ { index: number, rangeId: string, contentControlId?: number, tag?: string, text?: string } ]`
- `total`: number — total paragraphs in scope (pre-cap)

Example args (MCP tool):
```json
{ "scope": "document", "ensureContentControls": true, "tagPrefix": "para", "includeText": true }
```

Example response shape:
```json
{
  "ok": true,
  "op": "paragraphs.index",
  "data": {
    "total": 42,
    "items": [
      { "index": 0, "rangeId": "range-abc", "contentControlId": 101, "tag": "para:0", "text": "Title" },
      { "index": 1, "rangeId": "range-def", "contentControlId": 102, "tag": "para:1", "text": "Introduction ..." }
    ]
  }
}
```

Office.js mapping and notes:
- Resolve target range by `scope`:
  - `document` → `const target = context.document.body;`
  - `selection` → `const target = context.document.getSelection();`
  - `rangeId:<id>` → look up tracked `Word.Range` by id.
- Enumerate paragraphs: `const paras = target.paragraphs; paras.load(["text", "items"]); await context.sync();`
- For each paragraph `p` (respecting `includeEmpty` and `max`):
  - Get its range: `const r = p.getRange();`
  - Determine existing content control:
    - Preferred: `const parent = r.parentContentControl` (or `p.parentContentControl` where available).
    - Fallback: `const ccColl = context.document.contentControls.getByRange(r);`
  - If `ensureContentControls=true` and none found:
    - `const cc = p.insertContentControl();`
    - Set metadata: `cc.tag = \`${tagPrefix}:${index}\`; cc.title = \`Paragraph ${index}\`;` (optionally set `cc.appearance = "boundingBox"`).
  - Track the range to return a `rangeId` (use the same tracking guidance as elsewhere in this spec: `range.track()` and store an opaque id).
  - Collect:
    - `index` — paragraph ordinal within the enumerated scope
    - `rangeId` — tracked id
    - `contentControlId` — `cc.id` when present
    - `tag` — `cc.tag` when present
    - `text` — when `includeText=true`, from `p.text`
- Notes:
  - `contentControl.id` is unique per document session; `tag` provides a human-readable, stable label. Providers may prefer `tag` for persistence.
  - Protected documents may block insertion of content controls; return `{ ok:false, code:"E_PERMISSION" }`.
  - Paragraph collections may include empty paragraphs; control via `includeEmpty`.

Authoring tips (NL → args):
- Say: "List all paragraphs and assign ids" →
  `{"ensureContentControls": true, "includeText": true}`
- Say: "Index paragraphs in selection only (no text)" →
  `{"scope":"selection","includeText":false}`


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
  - `paragraphs` (collection)
  - `paragraph.getRange()`
  - `paragraph.insertContentControl()`
  - `range.parentContentControl` (or `contentControls.getByRange(range)`)
  - `contentControl.id | tag | title | appearance`

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
  - `Word.SearchOptions` fields: `matchCase`, `matchWholeWord`, `matchPrefix`, `matchSuffix`, `ignoreSpace`, `ignorePunct`, `matchWildcards`


## Types (suggested TypeScript for validation)

```ts
export type Scope = "document" | "selection" | `rangeId:${string}`;
export type Location = "start" | "end" | "before" | "after" | "replace";

// Optional (only if you implement the aggregate tool pattern)
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

export interface ApplyStyleArgs {
  scope?: Scope;
  namedStyle?: string;
  char?: {
    bold?: boolean;
    italic?: boolean;
    underline?: "none" | "single" | "double";
    strikeThrough?: boolean;
    doubleStrikeThrough?: boolean;
    allCaps?: boolean;
    smallCaps?: boolean;
    superscript?: boolean;
    subscript?: boolean;
    fontName?: string;
    fontSize?: number;
    color?: string;
    highlight?: string;
  };
  para?: {
    alignment?: "left" | "center" | "right" | "justify";
    lineSpacing?: number;
    spaceBefore?: number;
    spaceAfter?: number;
    leftIndent?: number;
    rightIndent?: number;
    firstLineIndent?: number;
    list?: "none" | "bullet" | "number";
  };
  precedence?: "styleThenOverrides" | "overridesThenStyle";
  resetDirectFormatting?: boolean;
}

export interface ParagraphsIndexArgs {
  scope?: Scope;
  ensureContentControls?: boolean;
  tagPrefix?: string;
  includeText?: boolean;
  includeEmpty?: boolean;
  max?: number;
}

export interface ParagraphsIndexItem {
  index: number;
  rangeId: string;
  contentControlId?: number;
  tag?: string;
  text?: string;
}
```


## Minimal Integration Example (Socket.IO + MCP)

- MCP tool call: `insertText` with args `{ "text": "Hello, Word!", "scope": "selection", "location": "replace" }`.
- Socket.IO emit: server/provider emits `word:insertText` with the same args.
- Add-in handling outline:
  1) Receive event `word:insertText` with args.
  2) Execute Office.js: `context.document.getSelection().insertText(args.text, Word.InsertLocation.replace); await context.sync();`
  3) Return `{ ok: true, data: { rangeId, length } }`.

Compatibility (optional, aggregate tool): implement an additional MCP tool that forwards a `meta` envelope and emits a single event carrying `{ op, args }` if you prefer a single entry point.
