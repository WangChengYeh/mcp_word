import express from 'express';
import http from 'http';
import { Server } from 'socket.io';
import bodyParser from 'body-parser';
import archiver from 'archiver';
import path from 'path';
// Import Claude MCP SDK client

const app = express();
const server = http.createServer(app);
const io = new Server(server);

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
// Endpoint to receive CLI POST, optionally process EditTask via SDK, and broadcast result
app.post('/mcp', async (req, res) => {
  try {
    const { content } = req.body;
    let result;
    try {
      // Dynamic import of Claude MCP SDK
      const { MCPClient } = await import('claude-mcp-sdk');
      const client = new MCPClient({ apiKey: process.env.CLAUDE_API_KEY });
      result = await client.requestEditTask({ content });
    } catch (sdkError) {
      console.warn('Claude MCP SDK unavailable, falling back to echo:', sdkError.message);
      // Fallback: echo content
      result = { content };
    }
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
