#!/usr/bin/env node

const { McpServer } = require('@modelcontextprotocol/sdk/server/mcp.js');
const { StdioServerTransport } = require('@modelcontextprotocol/sdk/server/stdio.js');
const http = require('http');
const express = require('express');
const { Server } = require('socket.io');
const path = require('path');
const fs = require('fs');

// Configuration
const PORT = process.env.PORT || 3000;
const DEBUG = process.argv.includes('--debug');
const DEBUG_LOG_FILE = 'debug.log';

// Logging utility
class Logger {
    constructor(debug = false) {
        this.debug = debug;
        if (debug) {
            // Clear debug log on start
            fs.writeFileSync(DEBUG_LOG_FILE, '');
        }
    }

    log(level, message, data = null) {
        const timestamp = new Date().toISOString();
        const logMessage = `[${timestamp}] [${level}] ${message}`;
        
        console.log(logMessage);
        
        if (this.debug) {
            let debugMessage = logMessage;
            if (data) {
                debugMessage += `\nData: ${JSON.stringify(data, null, 2)}`;
            }
            fs.appendFileSync(DEBUG_LOG_FILE, debugMessage + '\n');
        }
    }

    info(message, data) { this.log('INFO', message, data); }
    warn(message, data) { this.log('WARN', message, data); }
    error(message, data) { this.log('ERROR', message, data); }
    debug(message, data) { 
        if (this.debug) this.log('DEBUG', message, data); 
    }
}

const logger = new Logger(DEBUG);

// MCP Server Setup
class WordMCPServer {
    constructor() {
        this.server = new McpServer({
            name: 'mcp-word-server',
            version: '1.0.0',
            description: 'MCP server for Word document editing via Office Add-in'
        });
        
        this.socketClients = new Set();
        this.setupMCPTools();
        logger.info('MCP Word Server initialized');
    }

    setupMCPTools() {
        // Register edit_document tool
        this.server.registerTool({
            name: 'edit_document',
            description: 'Edit Word document content via Office Add-in',
            parameters: {
                type: 'object',
                properties: {
                    content: {
                        type: 'string',
                        description: 'Text content to insert or edit in the document'
                    },
                    action: {
                        type: 'string',
                        enum: ['insert', 'replace', 'append'],
                        default: 'insert',
                        description: 'Action to perform with the content'
                    },
                    position: {
                        type: 'string',
                        enum: ['cursor', 'start', 'end'],
                        default: 'cursor',
                        description: 'Position to perform the action'
                    }
                },
                required: ['content']
            }
        }, async (params) => {
            logger.info('Received edit_document request', params);
            
            try {
                // Send command to all connected Office Add-in clients
                const command = {
                    action: params.action || 'insert',
                    content: params.content,
                    position: params.position || 'cursor',
                    timestamp: new Date().toISOString()
                };

                this.broadcastToClients('ai-cmd', command);
                
                logger.info('Document edit command sent to Office clients', command);
                
                return {
                    success: true,
                    message: `Content ${command.action} command sent to ${this.socketClients.size} Office client(s)`,
                    clientCount: this.socketClients.size
                };
            } catch (error) {
                logger.error('Error in edit_document tool', error);
                return {
                    success: false,
                    error: error.message
                };
            }
        });

        // Register get_document_status tool
        this.server.registerTool({
            name: 'get_document_status',
            description: 'Get status of connected Office Add-in clients',
            parameters: {
                type: 'object',
                properties: {}
            }
        }, async () => {
            logger.debug('Getting document status');
            
            return {
                connectedClients: this.socketClients.size,
                serverStatus: 'running',
                timestamp: new Date().toISOString()
            };
        });

        logger.info('MCP tools registered successfully');
    }

    broadcastToClients(event, data) {
        if (this.io) {
            this.io.emit(event, data);
            logger.debug(`Broadcasted ${event} to ${this.socketClients.size} clients`, data);
        }
    }

    setSocketIO(io) {
        this.io = io;
    }

    async start() {
        const transport = new StdioServerTransport();
        await this.server.connect(transport);
        logger.info('MCP server started on STDIO transport');
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

    // Socket.io connection handling
    io.on('connection', (socket) => {
        logger.info(`Office client connected: ${socket.id}`);
        mcpServer.socketClients.add(socket);

        // Handle client identification
        socket.on('client-info', (data) => {
            logger.info('Client info received', { socketId: socket.id, data });
        });

        // Handle document status updates from client
        socket.on('document-status', (data) => {
            logger.debug('Document status update', { socketId: socket.id, data });
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

// Main execution
async function main() {
    try {
        logger.info('Starting MCP Word Server...', { debug: DEBUG, port: PORT });

        const mcpServer = new WordMCPServer();

        // Check if running in STDIO mode (typical for MCP)
        if (process.stdin.isTTY === false) {
            // STDIO mode - start MCP server
            logger.info('Running in STDIO mode for MCP client');
            await mcpServer.start();
        } else {
            // Interactive mode - start web server
            logger.info('Running in interactive mode with web server');
            
            const { server } = createWebServer(mcpServer);
            
            server.listen(PORT, () => {
                logger.info(`Web server started on http://localhost:${PORT}`);
                logger.info('Office Add-in can connect via Socket.io');
                
                if (DEBUG) {
                    logger.info(`Debug mode enabled. Logs written to ${DEBUG_LOG_FILE}`);
                }
            });

            // Handle graceful shutdown
            process.on('SIGINT', () => {
                logger.info('Shutting down server...');
                server.close(() => {
                    logger.info('Server shut down gracefully');
                    process.exit(0);
                });
            });
        }

    } catch (error) {
        logger.error('Failed to start server', error);
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
if (require.main === module) {
    main();
}

module.exports = { WordMCPServer, Logger };
