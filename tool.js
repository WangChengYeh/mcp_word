// MCP tool registration mapping to Socket.IO events (see tool.md)
// Exports registerTools(mcp, io)

import { z } from "zod";

export function registerTools(mcp, io, log = () => {}, logErr = () => {}) {
  const Scope = z.union([z.literal("document"), z.literal("selection"), z.string().regex(/^rangeId:.+/)]);
  const Location = z.union([
    z.literal("start"),
    z.literal("end"),
    z.literal("before"),
    z.literal("after"),
    z.literal("replace"),
  ]);

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
      target: z.union([Scope, z.literal("searchQuery")]),
      query: z.string().optional(),
      useRegex: z.boolean().optional(),
      matchCase: z.boolean().optional(),
      matchWholeWord: z.boolean().optional(),
      matchPrefix: z.boolean().optional(),
      replaceWith: z.string(),
      mode: z.union([z.literal("replaceFirst"), z.literal("replaceAll")]).optional(),
      scope: Scope.optional(),
    },
    "word:replace"
  );

  // ---------- Pictures ----------
  reg(
    "insertPicture",
    {
      source: z.union([z.literal("url"), z.literal("base64")]),
      data: z.string(),
      scope: Scope.optional().default("selection"),
      location: Location.optional().default("replace"),
      width: z.number().optional(),
      height: z.number().optional(),
      lockAspectRatio: z.boolean().optional(),
      altText: z.string().optional(),
      wrapType: z
        .union([
          z.literal("inline"),
          z.literal("square"),
          z.literal("tight"),
          z.literal("behind"),
          z.literal("inFront"),
        ])
        .optional(),
    },
    "word:insertPicture"
  );

  // ---------- Tables ----------
  const tableRef = z.union([z.string().regex(/^tableId:.+/), z.string().regex(/^rangeId:.+/)]);
  reg(
    "table.create",
    { rows: z.number(), cols: z.number(), scope: Scope.optional(), location: Location.optional(), data: z.array(z.array(z.string())).optional(), header: z.boolean().optional() },
    "word:table.create"
  );
  reg("table.insertRows", { tableRef: tableRef, at: z.number(), count: z.number() }, "word:table.insertRows");
  reg("table.insertColumns", { tableRef: tableRef, at: z.number(), count: z.number() }, "word:table.insertColumns");
  reg("table.deleteRows", { tableRef: tableRef, indexes: z.array(z.number()) }, "word:table.deleteRows");
  reg("table.deleteColumns", { tableRef: tableRef, indexes: z.array(z.number()) }, "word:table.deleteColumns");
  reg("table.setCellText", { tableRef: tableRef, row: z.number(), col: z.number(), text: z.string() }, "word:table.setCellText");
  reg(
    "table.mergeCells",
    { tableRef: tableRef, startRow: z.number(), startCol: z.number(), rowSpan: z.number(), colSpan: z.number() },
    "word:table.mergeCells"
  );

  // ---------- Styles ----------
  reg(
    "applyStyle",
    {
      scope: Scope.optional(),
      namedStyle: z.string().optional(),
      precedence: z.union([z.literal("styleThenOverrides"), z.literal("overridesThenStyle")]).optional(),
      resetDirectFormatting: z.boolean().optional(),
      char: z
        .object({
          bold: z.boolean().optional(),
          italic: z.boolean().optional(),
          underline: z.union([z.literal("none"), z.literal("single"), z.literal("double")]).optional(),
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
          alignment: z.union([z.literal("left"), z.literal("center"), z.literal("right"), z.literal("justify")]).optional(),
          lineSpacing: z.number().optional(),
          spaceBefore: z.number().optional(),
          spaceAfter: z.number().optional(),
          leftIndent: z.number().optional(),
          rightIndent: z.number().optional(),
          firstLineIndent: z.number().optional(),
          list: z.union([z.literal("none"), z.literal("bullet"), z.literal("number")]).optional(),
        })
        .partial()
        .optional(),
    },
    "word:applyStyle"
  );

  reg(
    "listStyles",
    {
      category: z.union([z.literal("paragraph"), z.literal("character"), z.literal("table"), z.literal("all")]).optional(),
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
