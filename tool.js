// MCP tool registration mapping to Socket.IO events (see tool.md)
// Exports registerTools(mcp, io)

import { z } from "zod";

export function registerTools(mcp, io, log = () => {}, logErr = () => {}) {
  // Core enums/patterns with descriptions
  const Location = z
    .enum(["start", "end", "before", "after", "replace"])
    .describe("Insert position relative to target range.");

  const RangeId = z
    .string()
    .regex(/^rangeId:.+/, {
      message: "Must be 'rangeId:<id>'",
    })
    .describe("A tracked range reference: 'rangeId:<id>'.");

  const Scope = z
    .union([z.literal("document"), z.literal("selection"), RangeId])
    .describe("Where to target: document, selection, or a saved rangeId.");

  const TableRef = z
    .string()
    .regex(/^(tableId|rangeId):.+/, {
      message: "Must be 'tableId:<id>' or 'rangeId:<id>'",
    })
    .describe("Table handle: 'tableId:<id>' or 'rangeId:<id>'.");

  const emitTool = (event, args) => {
    try { log(`[emit] ${event} ${JSON.stringify(args || {})}`); } catch {}
    io.emit(event, args || {});
  };

  const reg = (name_in, schema, event, desc) => {
    let name = name_in.replace('.','_')
    mcp.registerTool(
      name,
      {
        description: desc || `${name} → Socket.IO '${event}'`,
        inputSchema: schema || {},
      },
      async (args, _ctx) => {
        try {
          emitTool(event, args || {});
          return { content: [{ type: "text", text: `${name} emitted (${event}).` }] };
        } catch (e) {
          logErr(e, name);
          return { isError: true, content: [{ type: "text", text: `${name} failed: ${String(e)}` }] };
        }
      }
    );
  };

  // ---------- Text ----------
  reg(
    "insertText",
    {
      text: z.string().describe("Text content to insert."),
      scope: Scope.optional().default("selection").describe("Where to insert (default: selection)."),
      location: Location.optional().default("replace").describe("Position relative to scope (default: replace)."),
      newParagraph: z.boolean().optional().describe("Insert as a new paragraph when true."),
      keepFormatting: z
        .boolean()
        .optional()
        .describe("Preserve surrounding formatting if possible."),
    },
    "word:insertText",
    "Insert or replace text at a target range."
  );

  reg(
    "getSelection",
    {},
    "word:getSelection",
    "Return current selection text and a reusable rangeId."
  );

  reg(
    "search",
    {
      query: z.string().describe("Text to find (wildcards supported)."),
      scope: Scope.optional().default("document").describe("Where to search (default: document)."),
      useRegex: z.boolean().optional().describe("Treat query as regex (may be unsupported)."),
      matchCase: z.boolean().optional().describe("Match case exactly."),
      matchWholeWord: z.boolean().optional().describe("Match whole words only."),
      matchPrefix: z.boolean().optional().describe("Match at word starts."),
      matchSuffix: z.boolean().optional().describe("Match at word ends."),
      ignoreSpace: z.boolean().optional().describe("Ignore whitespace differences."),
      ignorePunct: z.boolean().optional().describe("Ignore punctuation differences."),
      maxResults: z.number().optional().describe("Maximum results to return."),
    },
    "word:search",
    "Search within a scope for text with options."
  );

  reg(
    "replace",
    {
      target: z
        .union([Scope, z.literal("searchQuery")])
        .describe("What to replace: a scope, rangeId, or 'searchQuery'."),
      query: z
        .string()
        .optional()
        .describe("Search text when target is 'searchQuery'."),
      useRegex: z.boolean().optional().describe("Treat query as regex (may be unsupported)."),
      matchCase: z.boolean().optional().describe("Case-sensitive matching."),
      matchWholeWord: z.boolean().optional().describe("Whole-word matching."),
      matchPrefix: z.boolean().optional().describe("Prefix matching."),
      matchSuffix: z.boolean().optional().describe("Suffix matching."),
      ignoreSpace: z.boolean().optional().describe("Ignore whitespace differences."),
      ignorePunct: z.boolean().optional().describe("Ignore punctuation differences."),
      replaceWith: z.string().describe("Replacement text."),
      mode: z
        .enum(["replaceFirst", "replaceAll"])        .optional()
        .describe("Replace first match or all matches."),
    },
    "word:replace",
    "Replace text by scope, range, or search results."
  );

  // ---------- Pictures ----------
  reg(
    "insertPicture",
    {
      source: z.enum(["url", "base64"]).describe("Image input type."),
      data: z.string().describe("Image URL or base64 string."),
      scope: Scope.optional().default("selection").describe("Where to insert the image."),
      location: Location.optional().default("replace").describe("Insert position relative to scope."),
      width: z.number().optional().describe("Image width in points."),
      height: z.number().optional().describe("Image height in points."),
      lockAspectRatio: z.boolean().optional().describe("Keep width/height proportional."),
      altText: z.string().optional().describe("Accessibility description."),
      wrapType: z
        .enum(["inline", "square", "tight", "behind", "inFront"])        .optional()
        .describe("Text wrapping preference."),
    },
    "word:insertPicture",
    "Insert an image from URL or base64."
  );

  // ---------- Tables ----------
  reg(
    "table.create",
    {
      rows: z.number().describe("Number of table rows."),
      cols: z.number().describe("Number of table columns."),
      scope: Scope.optional().describe("Where to insert the table."),
      location: Location.optional().describe("Insert position relative to scope."),
      data: z
        .array(z.array(z.string()))
        .optional()
        .describe("Initial cell values by row/column."),
      header: z.boolean().optional().describe("Treat first row as a header."),
    },
    "word:table.create",
    "Create a table at the target location."
  );
  reg(
    "table.insertRows",
    {
      tableRef: TableRef,
      at: z.number().describe("Zero-based row index to insert relative to."),
      count: z.number().describe("Number of rows to insert."),
    },
    "word:table.insertRows",
    "Insert rows into a table."
  );
  reg(
    "table.insertColumns",
    {
      tableRef: TableRef,
      at: z.number().describe("Zero-based column index to insert relative to."),
      count: z.number().describe("Number of columns to insert."),
    },
    "word:table.insertColumns",
    "Insert columns into a table."
  );
  reg(
    "table.deleteRows",
    { tableRef: TableRef, indexes: z.array(z.number()).describe("Row indexes to delete.") },
    "word:table.deleteRows",
    "Delete rows by index."
  );
  reg(
    "table.deleteColumns",
    { tableRef: TableRef, indexes: z.array(z.number()).describe("Column indexes to delete.") },
    "word:table.deleteColumns",
    "Delete columns by index."
  );
  reg(
    "table.setCellText",
    {
      tableRef: TableRef,
      row: z.number().describe("Zero-based row index."),
      col: z.number().describe("Zero-based column index."),
      text: z.string().describe("Cell text content."),
    },
    "word:table.setCellText",
    "Set text for a specific cell."
  );
  reg(
    "table.mergeCells",
    {
      tableRef: TableRef,
      startRow: z.number().describe("Starting row index (zero-based)."),
      startCol: z.number().describe("Starting column index (zero-based)."),
      rowSpan: z.number().describe("Number of rows to span."),
      colSpan: z.number().describe("Number of columns to span."),
    },
    "word:table.mergeCells",
    "Merge a rectangular cell region."
  );

  // ---------- Styles ----------
  reg(
    "applyStyle",
    {
      scope: Scope.optional().describe("What range to format (default: selection)."),
      namedStyle: z.string().optional().describe("Word style name to apply (e.g., 'Heading 1')."),
      precedence: z
        .enum(["styleThenOverrides", "overridesThenStyle"])        .optional()
        .describe("Order of applying style vs overrides."),
      resetDirectFormatting: z
        .boolean()
        .optional()
        .describe("Clear existing direct formatting before applying."),
      char: z
        .object({
          bold: z.boolean().optional().describe("Make text bold."),
          italic: z.boolean().optional().describe("Italicize text."),
          underline: z
            .enum(["none", "single", "double"])            .optional()
            .describe("Underline style."),
          strikeThrough: z.boolean().optional().describe("Apply single strikethrough."),
          doubleStrikeThrough: z.boolean().optional().describe("Apply double strikethrough."),
          allCaps: z.boolean().optional().describe("Render letters as uppercase."),
          smallCaps: z.boolean().optional().describe("Render letters as small caps."),
          superscript: z.boolean().optional().describe("Raise text above baseline."),
          subscript: z.boolean().optional().describe("Lower text below baseline."),
          fontName: z.string().optional().describe("Font family name."),
          fontSize: z.number().optional().describe("Font size in points."),
          color: z.string().optional().describe("Text color (e.g., '#333333')."),
          highlight: z.string().optional().describe("Text highlight color."),
        })
        .partial()
        .optional(),
      para: z
        .object({
          alignment: z
            .enum(["left", "center", "right", "justify"])            .optional()
            .describe("Paragraph alignment."),
          lineSpacing: z.number().optional().describe("Line spacing multiplier (e.g., 1.15)."),
          spaceBefore: z.number().optional().describe("Space before paragraph (pt)."),
          spaceAfter: z.number().optional().describe("Space after paragraph (pt)."),
          leftIndent: z.number().optional().describe("Left indent (pt)."),
          rightIndent: z.number().optional().describe("Right indent (pt)."),
          firstLineIndent: z.number().optional().describe("First-line indent (pt)."),
          list: z
            .enum(["none", "bullet", "number"])            .optional()
            .describe("List formatting type."),
        })
        .partial()
        .optional(),
    },
    "word:applyStyle",
    "Apply named styles and direct formatting to a range."
  );

  reg(
    "listStyles",
    {
      category: z
        .enum(["paragraph", "character", "table", "all"])        .optional()
        .describe("Filter by style category."),
      query: z.string().optional().describe("Filter by name substring."),
      builtInOnly: z.boolean().optional().describe("Include only built-in styles."),
      includeLocalized: z
        .boolean()
        .optional()
        .describe("Include localized display names."),
      max: z.number().optional().describe("Maximum styles to return."),
    },
    "word:listStyles",
    "List available Word styles with filtering options."
  );

  // ping tool (utility)
  mcp.registerTool(
    "ping",
    {
      description: "Health check tool.",
      inputSchema: {
        message: z.string().optional().describe("Optional message to echo."),
      },
    },
    async (args, _ctx) => ({ content: [{ type: "text", text: args?.message || "pong" }] })
  );
}
