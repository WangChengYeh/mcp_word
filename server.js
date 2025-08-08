#!/usr/bin/env node

import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import http from 'http';
import express from 'express';
import { Server } from 'socket.io';
import path from 'path';
import fs from 'fs';
import { fileURLToPath } from 'url';

// ES6 module path resolution
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

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
        // Register edit_document tool with enhanced EditTask support
        this.server.registerTool({
            name: 'edit_document',
            description: 'Edit Word document content via Office Add-in with various EditTask types',
            parameters: {
                type: 'object',
                properties: {
                    content: {
                        type: 'string',
                        description: 'Text content to insert or edit in the document'
                    },
                    action: {
                        type: 'string',
                        enum: ['insert', 'replace', 'append', 'delete'],
                        default: 'insert',
                        description: 'Action to perform with the content'
                    },
                    position: {
                        type: 'string',
                        enum: ['cursor', 'start', 'end', 'selection'],
                        default: 'cursor',
                        description: 'Position to perform the action'
                    },
                    taskType: {
                        type: 'string',
                        enum: ['text', 'table', 'image', 'formatting'],
                        default: 'text',
                        description: 'Type of EditTask to perform'
                    },
                    formatting: {
                        type: 'object',
                        properties: {
                            bold: { type: 'boolean' },
                            italic: { type: 'boolean' },
                            fontSize: { type: 'number' },
                            color: { type: 'string' }
                        },
                        description: 'Formatting options for text'
                    }
                },
                required: ['content']
            }
        }, async (params) => {
            logger.info('Received edit_document request', params);
            
            try {
                // Enhanced command structure for different EditTask types
                const command = {
                    action: params.action || 'insert',
                    content: params.content,
                    position: params.position || 'cursor',
                    taskType: params.taskType || 'text',
                    formatting: params.formatting || {},
                    timestamp: new Date().toISOString(),
                    id: Math.random().toString(36).substr(2, 9)
                };

                this.broadcastToClients('ai-cmd', command);
                
                logger.info('Enhanced EditTask command sent to Office clients', command);
                
                return {
                    success: true,
                    message: `${command.taskType} ${command.action} command sent to ${this.socketClients.size} Office client(s)`,
                    clientCount: this.socketClients.size,
                    commandId: command.id,
                    taskType: command.taskType
                };
            } catch (error) {
                logger.error('Error in edit_document tool', error);
                return {
                    success: false,
                    error: error.message,
                    timestamp: new Date().toISOString()
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

        // Register new table creation tool
        this.server.registerTool({
            name: 'create_table',
            description: 'Create a table in the Word document',
            parameters: {
                type: 'object',
                properties: {
                    rows: {
                        type: 'number',
                        description: 'Number of rows in the table'
                    },
                    columns: {
                        type: 'number', 
                        description: 'Number of columns in the table'
                    },
                    headers: {
                        type: 'array',
                        items: { type: 'string' },
                        description: 'Header row content'
                    },
                    data: {
                        type: 'array',
                        items: {
                            type: 'array',
                            items: { type: 'string' }
                        },
                        description: 'Table data as array of rows'
                    }
                },
                required: ['rows', 'columns']
            }
        }, async (params) => {
            logger.info('Received create_table request', params);
            
            const command = {
                action: 'create_table',
                taskType: 'table',
                rows: params.rows,
                columns: params.columns,
                headers: params.headers || [],
                data: params.data || [],
                timestamp: new Date().toISOString(),
                id: Math.random().toString(36).substr(2, 9)
            };

            this.broadcastToClients('ai-cmd', command);
            
            return {
                success: true,
                message: `Table creation command sent (${params.rows}x${params.columns})`,
                clientCount: this.socketClients.size,
                commandId: command.id
            };
        });

        logger.info('Enhanced MCP tools registered successfully');
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
if (require.main === module) {
    main();
}

// Export for testing
export { WordMCPServer, Logger };
