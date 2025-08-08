#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { Server as SocketIOServer } from 'socket.io';
import http from 'http';
import fs from 'fs';

const DEBUG = process.argv.includes('--debug');
const PORT = 3000;

// Debug logging function
function debugLog(message, data = null) {
  if (DEBUG) {
    const timestamp = new Date().toISOString();
    const logEntry = `[${timestamp}] ${message}${data ? '\n' + JSON.stringify(data, null, 2) : ''}\n`;
    fs.appendFileSync('debug.log', logEntry);
    console.error(`DEBUG: ${message}`, data || '');
  }
}

// Create HTTP server for Socket.io
const httpServer = http.createServer((req, res) => {
  if (req.url === '/taskpane.html') {
    // Serve taskpane.html for Office Add-in
    try {
      const html = fs.readFileSync('./public/taskpane.html', 'utf8');
      res.writeHead(200, { 'Content-Type': 'text/html' });
      res.end(html);
    } catch (error) {
      debugLog('Error serving taskpane.html', { error: error.message });
      res.writeHead(404);
      res.end('File not found');
    }
  } else if (req.url === '/taskpane.js') {
    // Serve taskpane.js
    try {
      const js = fs.readFileSync('./public/taskpane.js', 'utf8');
      res.writeHead(200, { 'Content-Type': 'application/javascript' });
      res.end(js);
    } catch (error) {
      debugLog('Error serving taskpane.js', { error: error.message });
      res.writeHead(404);
      res.end('File not found');
    }
  } else {
    res.writeHead(404);
    res.end('Not found');
  }
});

// Setup Socket.io server
const io = new SocketIOServer(httpServer, {
  cors: {
    origin: "*",
    methods: ["GET", "POST"]
  }
});

// Socket.io connection handling
io.on('connection', (socket) => {
  debugLog('Office Add-in connected', { socketId: socket.id });
  
  socket.on('disconnect', () => {
    debugLog('Office Add-in disconnected', { socketId: socket.id });
  });
  
  socket.on('error', (error) => {
    debugLog('Socket error', { socketId: socket.id, error: error.message });
  });
  
  socket.on('ready', (data) => {
    debugLog('Add-in ready', { socketId: socket.id, data });
  });
});

// Start HTTP server
httpServer.listen(PORT, () => {
  debugLog(`Server listening on http://localhost:${PORT}`);
});

// Create MCP Server
const server = new McpServer(
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

// Register EditTask tool
server.registerTool(
  {
    name: "edit_document",
    description: "Send edit commands to Word document via Office Add-in",
    inputSchema: {
      type: "object",
      properties: {
        content: {
          type: "string",
          description: "Text content to insert or edit"
        },
        action: {
          type: "string",
          enum: ["insert", "replace", "append"],
          description: "Type of edit action",
          default: "insert"
        },
        position: {
          type: "string",
          enum: ["cursor", "start", "end", "selection"],
          description: "Position for the edit",
          default: "cursor"
        }
      },
      required: ["content"]
    },
  },
  async (request) => {
    try {
      debugLog('Received edit_document request', request);
      
      const { content, action = "insert", position = "cursor" } = request.params.arguments;
      
      // Create EditTask
      const editTask = {
        content,
        action,
        position,
        timestamp: new Date().toISOString()
      };
      
      // Send to all connected Office Add-ins via Socket.io
      io.emit('ai-cmd', editTask);
      
      debugLog('Sent EditTask to Office Add-ins', editTask);
      
      return {
        content: [
          {
            type: "text",
            text: `Edit command sent: ${action} "${content}" at ${position}`
          }
        ]
      };
      
    } catch (error) {
      debugLog('Error in edit_document tool', { error: error.message, stack: error.stack });
      
      return {
        content: [
          {
            type: "text",
            text: `Error: ${error.message}`
          }
        ],
        isError: true
      };
    }
  }
);

// Register additional tool for document info
server.registerTool(
  {
    name: "get_document_status",
    description: "Get current document status from Office Add-in",
    inputSchema: {
      type: "object",
      properties: {}
    },
  },
  async (request) => {
    try {
      debugLog('Received get_document_status request');
      
      // Request status from Office Add-ins
      io.emit('get-status', { timestamp: new Date().toISOString() });
      
      return {
        content: [
          {
            type: "text",
            text: "Document status request sent to Office Add-in"
          }
        ]
      };
      
    } catch (error) {
      debugLog('Error in get_document_status tool', { error: error.message });
      
      return {
        content: [
          {
            type: "text",
            text: `Error: ${error.message}`
          }
        ],
        isError: true
      };
    }
  }
);

// Global error handling
process.on('uncaughtException', (error) => {
  debugLog('Uncaught exception', { error: error.message, stack: error.stack });
  process.exit(1);
});

process.on('unhandledRejection', (reason, promise) => {
  debugLog('Unhandled rejection', { reason, promise });
});

// Initialize MCP server with STDIO transport
async function main() {
  try {
    debugLog('Starting MCP Word Server');
    
    // Create STDIO transport for MCP communication
    const transport = new StdioServerTransport();
    
    // Connect MCP server
    await server.connect(transport);
    
    debugLog('MCP server connected and ready for STDIO communication');
    
  } catch (error) {
    debugLog('Failed to start MCP server', { error: error.message, stack: error.stack });
    process.exit(1);
  }
}

// Start the server
main();
    }
}

// Express and Socket.io Setup
function createWebServer(mcpServer) {
    const app = express();
    const server = http.createServer(app);
    const io = new Server(server, {
        cors: {
            origin: "*",
            methods: ["GET", "POST"]
        }
    });

    // Set socket.io reference in MCP server
    mcpServer.setSocketIO(io);

    // Serve static files
    app.use(express.static(path.join(__dirname, 'public')));

    // Health check endpoint
    app.get('/health', (req, res) => {
        res.json({
            status: 'healthy',
            timestamp: new Date().toISOString(),
            connectedClients: mcpServer.socketClients.size
        });
    });

    // Socket.io connection handling with enhanced error tracking
    io.on('connection', (socket) => {
        logger.info(`Office client connected: ${socket.id}`);
        mcpServer.socketClients.add(socket);

        // Enhanced client identification with Office.js info
        socket.on('client-info', (data) => {
            logger.info('Office client info received', { 
                socketId: socket.id, 
                officeVersion: data.officeVersion,
                platform: data.platform,
                host: data.host
            });
        });

        // Handle command execution results
        socket.on('command-result', (data) => {
            logger.info('Command execution result', { 
                socketId: socket.id, 
                commandId: data.commandId,
                success: data.success,
                error: data.error 
            });
        });

        // Enhanced document status with Word-specific info
        socket.on('document-status', (data) => {
            logger.debug('Document status update', { 
                socketId: socket.id, 
                documentName: data.documentName,
                wordCount: data.wordCount,
                selectionRange: data.selectionRange
            });
        });

        // WebSocket authentication placeholder
        socket.on('authenticate', (token) => {
            // TODO: Implement authentication logic
            logger.debug('Authentication attempt', { socketId: socket.id });
            socket.emit('auth-result', { success: true, message: 'Authentication not yet implemented' });
        });

        // Handle client errors
        socket.on('client-error', (error) => {
            logger.error('Client reported error', { socketId: socket.id, error });
        });

        // Handle test messages
        socket.on('test-message', (data) => {
            logger.debug('Test message received', { socketId: socket.id, data });
            socket.emit('test-response', { 
                message: 'Test message received successfully',
                timestamp: new Date().toISOString()
            });
        });

        socket.on('disconnect', (reason) => {
            logger.info(`Office client disconnected: ${socket.id}, reason: ${reason}`);
            mcpServer.socketClients.delete(socket);
        });
    });

    return { app, server, io };
}

// Enhanced main execution with better error handling
async function main() {
    try {
        logger.info('Starting MCP Word Server...', { 
            debug: DEBUG, 
            port: PORT,
            nodeVersion: process.version,
            platform: process.platform
        });

        const mcpServer = new WordMCPServer();

        // Check if running in STDIO mode (typical for MCP)
        if (process.stdin.isTTY === false) {
            // STDIO mode - start MCP server for Claude CLI integration
            logger.info('Running in STDIO mode for MCP client (Claude CLI)');
            await mcpServer.start();
        } else {
            // Interactive mode - start web server for Office Add-in
            logger.info('Running in interactive mode with web server for Office Add-in');
            
            const { server } = createWebServer(mcpServer);
            
            server.listen(PORT, () => {
                logger.info(`Web server started on http://localhost:${PORT}`);
                logger.info('Office Add-in manifest URL: http://localhost:${PORT}/manifest.xml');
                logger.info('Task pane URL: http://localhost:${PORT}/taskpane.html');
                
                if (DEBUG) {
                    logger.info(`Debug mode enabled. Detailed logs written to ${DEBUG_LOG_FILE}`);
                }
            });

            // Enhanced graceful shutdown
            const gracefulShutdown = () => {
                logger.info('Received shutdown signal, closing server gracefully...');
                server.close(() => {
                    logger.info('HTTP server closed');
                    if (DEBUG) {
                        logger.info('Final debug log written');
                    }
                    process.exit(0);
                });
                
                // Force close after 10 seconds
                setTimeout(() => {
                    logger.error('Could not close connections in time, forcefully shutting down');
                    process.exit(1);
                }, 10000);
            };

            process.on('SIGINT', gracefulShutdown);
            process.on('SIGTERM', gracefulShutdown);
        }

    } catch (error) {
        logger.error('Failed to start server', { 
            error: error.message, 
            stack: error.stack 
        });
        process.exit(1);
    }
}

// Handle unhandled errors
process.on('unhandledRejection', (reason, promise) => {
    logger.error('Unhandled Rejection', { reason, promise });
});

process.on('uncaughtException', (error) => {
    logger.error('Uncaught Exception', error);
    process.exit(1);
});

// Start the server
main();

// Export for testing
export { WordMCPServer, Logger };
