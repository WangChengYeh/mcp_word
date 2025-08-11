// MCP tool registration mapping to Socket.IO events (see tool.md)
// Exports registerTools(mcp, io)

import { z } from "zod";

export function registerTools(mcp, io, log = () => {}, logErr = () => {}) {
  // Allow fixed values or a rangeId:... reference
  const Scope = z.enum(["document", "selection"]);
  const Location = z.enum(["start", "end", "before", "after", "replace"]);

  const emitTool = (event, args) => {
    try { log(`[emit] ${event} ${JSON.stringify(args || {})}`); } catch {}
    io.emit(event, args || {});
  };

  const reg = (name, schema, event, desc) => {
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
      text: z.string(),
      scope: Scope.optional().default("selection"),
      location: Location.optional().default("replace"),
      newParagraph: z.boolean().optional(),
      keepFormatting: z.boolean().optional(),
    },
    "word:insertText"
  );

  reg("getSelection", {}, "word:getSelection");

  reg(
    "search",
    {
      query: z.string(),
      scope: Scope.optional().default("document"),
      useRegex: z.boolean().optional(),
      matchCase: z.boolean().optional(),
      matchWholeWord: z.boolean().optional(),
      matchPrefix: z.boolean().optional(),
      matchSuffix: z.boolean().optional(),
      ignoreSpace: z.boolean().optional(),
      ignorePunct: z.boolean().optional(),
      maxResults: z.number().optional(),
    },
    "word:search"
  );

  reg(
    "replace",
    {
      target: Scope,
      query: z.string().optional(),
      useRegex: z.boolean().optional(),
      matchCase: z.boolean().optional(),
      matchWholeWord: z.boolean().optional(),
      matchPrefix: z.boolean().optional(),
      replaceWith: z.string(),
      mode: z.enum(["replaceFirst", "replaceAll"]).optional(),
    },
    "word:replace"
  );

  // ---------- Pictures ----------
  reg(
    "insertPicture",
    {
      source: z.enum(["url", "base64"]),
      data: z.string(),
      scope: Scope.optional().default("selection"),
      location: Location.optional().default("replace"),
      width: z.number().optional(),
      height: z.number().optional(),
      lockAspectRatio: z.boolean().optional(),
      altText: z.string().optional(),
      wrapType: z.enum(["inline", "square", "tight", "behind", "inFront"]).optional(),
    },
    "word:insertPicture"
  );

  // ---------- Tables ----------
  const tableRef = z.enum(["tableId", "rangeId"]);
  reg(
    "table_create",
    { rows: z.number(), cols: z.number(), scope: Scope.optional(), location: Location.optional(), data: z.array(z.array(z.string())).optional(), header: z.boolean().optional() },
    "word:table.create"
  );
  reg("table_insertRows", { tableRef: tableRef, at: z.number(), count: z.number() }, "word:table.insertRows");
  reg("table_insertColumns", { tableRef: tableRef, at: z.number(), count: z.number() }, "word:table.insertColumns");
  reg("table_deleteRows", { tableRef: tableRef, indexes: z.array(z.number()) }, "word:table.deleteRows");
  reg("table_deleteColumns", { tableRef: tableRef, indexes: z.array(z.number()) }, "word:table.deleteColumns");
  reg("table_setCellText", { tableRef: tableRef, row: z.number(), col: z.number(), text: z.string() }, "word:table.setCellText");
  reg(
    "table_mergeCells",
    { tableRef: tableRef, startRow: z.number(), startCol: z.number(), rowSpan: z.number(), colSpan: z.number() },
    "word:table.mergeCells"
  );

  // ---------- Styles ----------
  reg(
    "applyStyle",
    {
      scope: Scope.optional(),
      namedStyle: z.string().optional(),
      precedence: z.enum(["styleThenOverrides", "overridesThenStyle"]).optional(),
      resetDirectFormatting: z.boolean().optional(),
      char: z
        .object({
          bold: z.boolean().optional(),
          italic: z.boolean().optional(),
          underline: z.enum(["none", "single", "double"]).optional(),
          strikeThrough: z.boolean().optional(),
          doubleStrikeThrough: z.boolean().optional(),
          allCaps: z.boolean().optional(),
          smallCaps: z.boolean().optional(),
          superscript: z.boolean().optional(),
          subscript: z.boolean().optional(),
          fontName: z.string().optional(),
          fontSize: z.number().optional(),
          color: z.string().optional(),
          highlight: z.string().optional(),
        })
        .partial()
        .optional(),
      para: z
        .object({
          alignment: z.enum(["left", "center", "right", "justify"]).optional(),
          lineSpacing: z.number().optional(),
          spaceBefore: z.number().optional(),
          spaceAfter: z.number().optional(),
          leftIndent: z.number().optional(),
          rightIndent: z.number().optional(),
          firstLineIndent: z.number().optional(),
          list: z.enum(["none", "bullet", "number"]).optional(),
        })
        .partial()
        .optional(),
    },
    "word:applyStyle"
  );

  reg(
    "listStyles",
    {
      category: z.enum(["paragraph", "character", "table", "all"]).optional(),
      query: z.string().optional(),
      builtInOnly: z.boolean().optional(),
      includeLocalized: z.boolean().optional(),
      max: z.number().optional(),
    },
    "word:listStyles"
  );

  // ping tool (utility)
  mcp.registerTool(
    "ping",
    { description: "Health check tool.", inputSchema: { message: z.string().optional() } },
    async (args, _ctx) => ({ content: [{ type: "text", text: args?.message || "pong" }] })
  );
}
