// Tool registration and forwarding logic per SPEC.md
// Exports registerTools() which registers MCP tools and forwards editTask args verbatim to Socket.IO

import { z } from "zod";

export function registerTools(mcp, io, log = () => {}, logErr = () => {}) {
  // editTask: forward the provided arguments object as-is to the add-in via ai-cmd
  mcp.registerTool(
    "editTask",
    {
      description:
        "Send an edit task to the Office Add-in via WebSocket (event: ai-cmd).",
      inputSchema: {
        content: z
          .string()
          .describe("Text to insert/replace in the Word document."),
        action: z
          .enum(["insert", "replace", "append"]) 
          .default("insert")
          .describe("insert | replace | append"),
        target: z
          .enum(["cursor", "selection", "document"]) 
          .default("selection")
          .describe("cursor | selection | document"),
        taskId: z.string().optional().describe("Optional client-correlated id."),
        meta: z.string().optional().describe("Optional additional metadata."),
      },
    },
    async (args, _ctx) => {
      try {
        // Log standardized DEBUG line for Socket.IO send
        const payload = args || {};
        try {
          log(
            `[DEBUG socket:send] ${JSON.stringify({ event: "editTask", payload })}`
          );
        } catch {}
        // Forward exactly the provided args on event named by tool
        io.emit("editTask", payload);
        return { content: [{ type: "text", text: "EditTask forwarded." }] };
      } catch (e) {
        logErr(e, "editTask");
        return {
          isError: true,
          content: [{ type: "text", text: `editTask failed: ${String(e)}` }],
        };
      }
    }
  );

  // ping tool
  mcp.registerTool(
    "ping",
    {
      description: "Health check tool.",
      inputSchema: { message: z.string().optional() },
    },
    async (args, _ctx) => {
      const text = args?.message || "pong";
      return { content: [{ type: "text", text }] };
    }
  );
}
