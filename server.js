import express from 'express';
import { Server } from 'socket.io';
import { createServer } from 'http';
import path from 'path';
import { fileURLToPath } from 'url';
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { SSEServerTransport } from '@modelcontextprotocol/sdk/server/sse.js';

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

// Store connected Office Add-in clients
const connectedClients = new Set();

// Store pending edits for tracking
const pendingEdits = new Map();

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

// EditTask handler function
async function handleEditTask(args) {
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
}, handleEditTask);

// MCP router using SSE (Server-Sent Events)
const mcpRouter = express.Router();
mcpRouter.use(express.json());

// Create SSE transport for MCP
const sseTransport = new SSEServerTransport('/mcp/sse', httpServer);

// Connect MCP server to SSE transport
mcpServer.connect(sseTransport);

// Mount MCP router
app.use('/mcp', mcpRouter);

debugLog('MCP SERVER', 'MCP server connected with SSE transport');

console.log('MCP Word server started (SSE transport)');

// Socket.io connection handling for Office Add-in
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
    pendingEdits: pendingEdits.size,
    transport: 'SSE'
  });
});

// Startup: node server.js (listens on port 3000 by default)
httpServer.listen(PORT, () => {
  console.log(`MCP Word Proxy Server running on http://localhost:${PORT}`);
  console.log(`MCP endpoint (SSE): http://localhost:${PORT}/mcp/sse`);
  console.log(`Office Add-in endpoint: http://localhost:${PORT}/office`);
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
        enum: ["cursor", "start", "end"],
        description: "Where to perform the operation",
        default: "cursor"
      }
    },
    required: ["content"]
  }
}, handleEditTask);

// Don't start stdio transport - only HTTP endpoint will be used
// const transport = new StdioServerTransport();
// mcpServer.connect(transport);

console.log('MCP Word server started (HTTP only)');

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
