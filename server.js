#!/usr/bin/env node

import { Server } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioTransport } from "@modelcontextprotocol/sdk/shared/transport.js";
import { Server as SocketIOServer } from "socket.io";
import { createServer } from "http";
import express from "express";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Create Express app and HTTP server for serving static files and Socket.IO
const app = express();
const httpServer = createServer(app);
const io = new SocketIOServer(httpServer, {
  cors: {
    origin: "*",
    methods: ["GET", "POST"]
  }
});

// Serve static files for the Office Add-in
app.use(express.static(path.join(__dirname, 'public')));

// Store connected Office Add-in clients
const connectedClients = new Set();

// Socket.IO connection handling
io.on('connection', (socket) => {
  console.error(`Office Add-in client connected: ${socket.id}`);
  connectedClients.add(socket);
  
  socket.on('disconnect', () => {
    console.error(`Office Add-in client disconnected: ${socket.id}`);
    connectedClients.delete(socket);
  });
  
  socket.on('error', (error) => {
    console.error(`Socket error from ${socket.id}:`, error);
  });
});

// Create MCP server
const server = new Server(
  {
    name: "mcp-word-server",
    version: "1.0.0",
  },
  {
    capabilities: {
      tools: {},
    },
  }
);

// Register the EditTask tool
server.tool("EditTask", "Send edit commands to connected Word document", {
  content: {
    type: "string",
    description: "Content to insert or edit in the Word document"
  },
  action: {
    type: "string",
    description: "Action to perform (insert, replace, append)",
    enum: ["insert", "replace", "append"],
    default: "insert"
  },
  position: {
    type: "string",
    description: "Position to perform action (start, end, cursor)",
    enum: ["start", "end", "cursor"],
    default: "cursor"
  }
}, async ({ content, action = "insert", position = "cursor" }) => {
  try {
    if (connectedClients.size === 0) {
      return {
        isError: true,
        content: [{
          type: "text",
          text: "No Word Add-in clients connected. Please open Word and load the add-in."
        }]
      };
    }

    // Send command to all connected Office Add-in clients
    const command = {
      action,
      content,
      position,
      timestamp: Date.now()
    };

    connectedClients.forEach(client => {
      client.emit('ai-cmd', command);
    });

    console.error(`Sent EditTask to ${connectedClients.size} client(s):`, command);

    return {
      content: [{
        type: "text",
        text: `EditTask sent successfully to ${connectedClients.size} Word client(s). Action: ${action}, Position: ${position}, Content length: ${content.length} characters.`
      }]
    };
  } catch (error) {
    console.error("Error in EditTask:", error);
    return {
      isError: true,
      content: [{
        type: "text",
        text: `Error executing EditTask: ${error.message}`
      }]
    };
  }
});

// Start HTTP server for Socket.IO and static files
const PORT = process.env.PORT || 3000;
httpServer.listen(PORT, () => {
  console.error(`MCP Word Server running on http://localhost:${PORT}`);
  console.error("Waiting for Office Add-in connections...");
});

// Start MCP server with stdio transport
const transport = new StdioTransport();
server.connect(transport).catch((error) => {
  console.error("Failed to start MCP server:", error);
  process.exit(1);
});

// Handle graceful shutdown
process.on('SIGINT', async () => {
  console.error("\nShutting down MCP Word Server...");
  await server.close();
  httpServer.close();
  process.exit(0);
});
