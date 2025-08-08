import express from 'express';
import { Server } from 'socket.io';
import { createServer } from 'http';
import path from 'path';
import { fileURLToPath } from 'url';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Parse command line arguments
const args = process.argv.slice(2);
const DEBUG_MODE = args.includes('--debug');

// Debug logging function
function debugLog(type, message, data = null) {
  if (DEBUG_MODE) {
    const timestamp = new Date().toISOString();
    console.log(`[DEBUG ${timestamp}] ${type}: ${message}`);
    if (data) {
      console.log(JSON.stringify(data, null, 2));
    }
  }
}

if (DEBUG_MODE) {
  console.log('[DEBUG MODE ENABLED] MCP requests and responses will be logged');
}

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

// Office Add-in endpoint - serve taskpane.html
app.get('/office', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'taskpane.html'));
});

// MCP endpoint for HTTP transport
app.use('/mcp', express.json());
app.post('/mcp', async (req, res) => {
  try {
    debugLog('HTTP MCP REQUEST', 'Received HTTP request', req.body);
    // Handle MCP over HTTP if needed
    res.json({ status: 'MCP endpoint ready', transport: 'stdio' });
  } catch (error) {
    debugLog('HTTP MCP ERROR', error.message, error);
    res.status(500).json({ error: error.message });
  }
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

// Add debug logging for MCP server events
if (DEBUG_MODE) {
  mcpServer.on('request', (request) => {
    debugLog('MCP REQUEST', `Method: ${request.method}`, request);
  });

  mcpServer.on('response', (response) => {
    debugLog('MCP RESPONSE', 'Response sent', response);
  });

  mcpServer.on('error', (error) => {
    debugLog('MCP ERROR', error.message, error);
  });
}

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
  debugLog('TOOL EXECUTION', 'EditTask called', args);

  const { content, operation = "insert", position = "cursor" } = args;

  if (connectedClients.size === 0) {
    const error = new Error("No Office Add-in clients connected");
    debugLog('TOOL ERROR', error.message);
    throw error;
  }

  // Generate unique edit ID for tracking
  const editId = `edit_${Date.now()}_${Math.random().toString(36).substr(2, 9)}`;
  
  // Create edit command with ID
  const editCommand = {
    editId,
    content,
    operation,
    position,
    timestamp: new Date().toISOString()
  };

  debugLog('WEBSOCKET EMIT', 'Sending ai-cmd to clients', editCommand);

  // Create promise for tracking edit completion
  return new Promise((resolve, reject) => {
    // Set timeout for edit operation
    const timeout = setTimeout(() => {
      pendingEdits.delete(editId);
      reject(new Error('Edit operation timed out'));
    }, 30000); // 30 second timeout

    // Store pending edit
    pendingEdits.set(editId, { resolve, reject, timeout });

    // Broadcast to all connected clients
    io.emit('ai-cmd', editCommand);
  });
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
  connectedClients.add(socket);
  debugLog('WEBSOCKET', `Client connected: ${socket.id}`);
  
  socket.on('disconnect', () => {
    console.log(`Office Add-in client disconnected: ${socket.id}`);
    connectedClients.delete(socket);
    debugLog('WEBSOCKET', `Client disconnected: ${socket.id}`);
  });
  
  // Handle client status updates
  socket.on('status', (data) => {
    console.log(`Client status: ${JSON.stringify(data)}`);
    debugLog('WEBSOCKET STATUS', `Client ${socket.id} status`, data);
  });
  
  // Return the edit results from the Office Add-in
  socket.on('edit-complete', (data) => {
    const { editId, success, message, error } = data;
    console.log(`Edit completed: ${editId}, Success: ${success}`);
    debugLog('WEBSOCKET EDIT', `Edit completed: ${editId}`, data);
    
    const pendingEdit = pendingEdits.get(editId);
    if (pendingEdit) {
      clearTimeout(pendingEdit.timeout);
      pendingEdits.delete(editId);
      
      if (success) {
        pendingEdit.resolve({ 
          success: true, 
          message: message || 'Edit applied successfully',
          editId 
        });
      } else {
        pendingEdit.reject(new Error(error || 'Edit failed in Office Add-in'));
      }
    }
  });
  
  // Handle Office Add-in errors
  socket.on('edit-error', (data) => {
    const { editId, error } = data;
    console.log(`Edit error: ${editId}, Error: ${error}`);
    debugLog('WEBSOCKET ERROR', `Edit error: ${editId}`, data);
    
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
  if (DEBUG_MODE) {
    console.log('[DEBUG MODE] Use --debug flag to see detailed MCP communication logs');
  }
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
