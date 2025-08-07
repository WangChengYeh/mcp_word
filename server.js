import express from 'express';
import http from 'http';
import { Server } from 'socket.io';
import bodyParser from 'body-parser';
import archiver from 'archiver';
import path from 'path';
// Import Claude MCP SDK client
import { MCPClient } from 'claude-mcp-sdk';

const app = express();
const server = http.createServer(app);
const io = new Server(server);
// Initialize Claude MCP SDK client
const mcpClient = new MCPClient({
  apiKey: process.env.CLAUDE_API_KEY,
});

// Middleware
app.use(bodyParser.json());
app.use(express.static(path.join(process.cwd(), 'public')));

// WebSocket connection handling
io.on('connection', (socket) => {
  console.log(`Client connected: ${socket.id}`);
  socket.on('disconnect', () => {
    console.log(`Client disconnected: ${socket.id}`);
  });
});

// Endpoint to receive CLI POST, process EditTask via MCP SDK, and broadcast result
app.post('/mcp', async (req, res) => {
  try {
    const { content } = req.body;
    // Process EditTask through Claude MCP SDK
    const result = await mcpClient.requestEditTask({ content });
    // Broadcast editing result to Office Add-in client
    io.emit('ai-cmd', result);
    res.json({ status: 'ok', result });
  } catch (error) {
    console.error('EditTask error:', error);
    res.status(500).json({ status: 'error', message: error.message });
  }
});


// Start server
const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
  console.log(`Server listening on http://localhost:${PORT}`);
});
