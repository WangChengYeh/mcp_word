import express from 'express';
import { createServer } from 'http';
import { Server as SocketIOServer } from 'socket.io';
import { Server } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import path from 'path';
import { fileURLToPath } from 'url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Express app setup
const app = express();
const httpServer = createServer(app);
const io = new SocketIOServer(httpServer, {
  cors: {
    origin: "*",
    methods: ["GET", "POST"]
  }
});

const PORT = process.env.PORT || 3000;

// Serve static files from public directory
app.use(express.static(path.join(__dirname, 'public')));

// MCP Server setup
const mcpServer = new Server(
  {
    name: "mcp-word-proxy",
    version: "1.0.0"
  },
  {
    capabilities: {
      tools: {}
    }
  }
);

// Register EditTask tool
mcpServer.setRequestHandler('tools/list', async () => {
  return {
    tools: [
      {
        name: "EditTask",
        description: "Handle AI-driven document editing tasks",
        inputSchema: {
          type: "object",
          properties: {
            content: {
              type: "string",
              description: "Content to be inserted or edited"
            },
            action: {
              type: "string",
              enum: ["insert", "replace", "append"],
              description: "Type of edit action to perform"
            }
          },
          required: ["content"]
        }
      }
    ]
  };
});

mcpServer.setRequestHandler('tools/call', async (request) => {
  const { name, arguments: args } = request.params;
  
  if (name === "EditTask") {
    const { content, action = "insert" } = args;
    
    // Broadcast edit command to connected Office Add-in clients
    io.emit('ai-cmd', {
      type: 'edit',
      content: content,
      action: action,
      timestamp: new Date().toISOString()
    });
    
    return {
      content: [
        {
          type: "text",
          text: `Edit task sent successfully. Action: ${action}, Content length: ${content.length} characters`
        }
      ]
    };
  }
  
  throw new Error(`Unknown tool: ${name}`);
});

// WebSocket connection handling
io.on('connection', (socket) => {
  console.log(`Office Add-in client connected: ${socket.id}`);
  
  socket.on('disconnect', () => {
    console.log(`Office Add-in client disconnected: ${socket.id}`);
  });
  
  // Handle client status updates
  socket.on('status', (data) => {
    console.log(`Client status: ${JSON.stringify(data)}`);
  });
  
  // Handle edit completion confirmation
  socket.on('edit-complete', (data) => {
    console.log(`Edit completed: ${JSON.stringify(data)}`);
  });
});

// Health check endpoint
app.get('/health', (req, res) => {
  res.json({ 
    status: 'ok', 
    timestamp: new Date().toISOString(),
    connectedClients: io.engine.clientsCount
  });
});

// Start HTTP server
httpServer.listen(PORT, () => {
  console.log(`MCP Word Proxy Server running on http://localhost:${PORT}`);
  console.log(`Static files served from: ${path.join(__dirname, 'public')}`);
});

// Initialize MCP server with stdio transport
async function initializeMCPServer() {
  const transport = new StdioServerTransport();
  await mcpServer.connect(transport);
  console.log('MCP Server initialized and connected via stdio');
}

// Handle graceful shutdown
process.on('SIGINT', async () => {
  console.log('\nShutting down gracefully...');
  httpServer.close();
  await mcpServer.close();
  process.exit(0);
});

// Initialize MCP server
initializeMCPServer().catch(console.error);
