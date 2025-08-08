import express from 'express';
import { Server } from 'socket.io';
import { createServer } from 'http';
import path from 'path';
import { fileURLToPath } from 'url';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const httpServer = createServer(app);
const io = new Server(httpServer, {
  cors: {
    origin: "*",
    methods: ["GET", "POST"]
  }
});

const PORT = process.env.PORT || 3000;

// Serve static files from public directory
app.use(express.static(path.join(__dirname, 'public')));

// Serve manifest.xml with correct MIME type
app.get('/manifest.xml', (req, res) => {
  res.type('application/xml');
  res.sendFile(path.join(__dirname, 'public', 'manifest.xml'));
});

// Store connected Office Add-in clients
const connectedClients = new Set();

// Socket.io connection handling
io.on('connection', (socket) => {
  console.log('Office Add-in client connected:', socket.id);
  connectedClients.add(socket);

  socket.on('disconnect', () => {
    console.log('Office Add-in client disconnected:', socket.id);
    connectedClients.delete(socket);
  });

  socket.on('edit-result', (data) => {
    console.log('Edit result received:', data);
    // Store result for MCP response if needed
  });
});

// MCP Server setup
const mcpServer = new McpServer({
  name: "mcp-word-server",
  version: "1.0.0"
});

// Register EditTask tool
mcpServer.registerTool({
  name: "EditTask",
  description: "Edit Word document content through the Office Add-in",
  inputSchema: {
    type: "object",
    properties: {
      content: {
        type: "string",
        description: "The text content to insert or edit in the document"
      },
      operation: {
        type: "string",
        enum: ["insert", "replace", "append"],
        description: "The type of edit operation to perform",
        default: "insert"
      },
      position: {
        type: "string",
        enum: ["cursor", "start", "end"],
        description: "Where to perform the operation",
        default: "cursor"
      }
    },
    required: ["content"]
  }
}, async (args) => {
  const { content, operation = "insert", position = "cursor" } = args;

  if (connectedClients.size === 0) {
    throw new Error("No Office Add-in clients connected");
  }

  // Send edit command to all connected clients
  const editCommand = {
    content,
    operation,
    position,
    timestamp: new Date().toISOString()
  };

  // Broadcast to all connected clients
  io.emit('ai-cmd', editCommand);

  return {
    success: true,
    message: `Edit command sent to ${connectedClients.size} client(s)`,
    command: editCommand
  };
});

// Start MCP server with stdio transport
const transport = new StdioServerTransport();
mcpServer.connect(transport);

console.log('MCP Word server started');

// Store pending edits for tracking
const pendingEdits = new Map();

// Enhanced socket.io connection handling
io.on('connection', (socket) => {
  console.log(`Office Add-in client connected: ${socket.id}`);
  
  socket.on('disconnect', () => {
    console.log(`Office Add-in client disconnected: ${socket.id}`);
  });
  
  // Handle client status updates
  socket.on('status', (data) => {
    console.log(`Client status: ${JSON.stringify(data)}`);
  });
  
  // Return the edit results from the Office Add-in
  socket.on('edit-complete', (data) => {
    const { editId, success, message, error } = data;
    console.log(`Edit completed: ${editId}, Success: ${success}`);
    
    const pendingEdit = pendingEdits.get(editId);
    if (pendingEdit) {
      clearTimeout(pendingEdit.timeout);
      pendingEdits.delete(editId);
      
      if (success) {
        pendingEdit.resolve({ message: message || 'Edit applied successfully' });
      } else {
        pendingEdit.reject(new Error(error || 'Edit failed in Office Add-in'));
      }
    }
  });
  
  // Handle Office Add-in errors
  socket.on('edit-error', (data) => {
    const { editId, error } = data;
    console.log(`Edit error: ${editId}, Error: ${error}`);
    
    const pendingEdit = pendingEdits.get(editId);
    if (pendingEdit) {
      clearTimeout(pendingEdit.timeout);
      pendingEdits.delete(editId);
      pendingEdit.reject(new Error(error));
    }
  });
});

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ 
    status: 'ok', 
    timestamp: new Date().toISOString(),
    connectedClients: io.engine.clientsCount,
    pendingEdits: pendingEdits.size
  });
});

// Startup: node server.js (listens on port 3000 by default)
httpServer.listen(PORT, () => {
  console.log(`MCP Word Proxy Server running on http://localhost:${PORT}`);
  console.log(`Static resources served from: ${path.join(__dirname, 'public')}`);
  console.log('Ready to serve manifest.xml, taskpane.html, taskpane.js');
});

// Handle graceful shutdown
process.on('SIGINT', async () => {
  console.log('\nShutting down gracefully...');
  
  // Clean up pending edits
  for (const [editId, pendingEdit] of pendingEdits) {
    clearTimeout(pendingEdit.timeout);
    pendingEdit.reject(new Error('Server shutting down'));
  }
  pendingEdits.clear();
  
  httpServer.close();
  await mcpServer.close();
  process.exit(0);
});
