import express from 'express';
import http from 'http';
import { Server } from 'socket.io';
import bodyParser from 'body-parser';
import archiver from 'archiver';
import path from 'path';

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

// Endpoint to receive CLI POST and broadcast 'ai-cmd'
app.post('/mcp', (req, res) => {
  const payload = req.body;
  io.emit('ai-cmd', payload);
  res.json({ status: 'ok' });
});

// Endpoint to download project as ZIP (excluding node_modules)
app.get('/download', (req, res) => {
  res.attachment('project.zip');
  const archive = archiver('zip', { zlib: { level: 9 } });
  archive.pipe(res);
  archive.glob('**/*', {
    cwd: process.cwd(),
    ignore: ['node_modules/**', '.git/**']
  });
  archive.finalize();
});

// Start server
const PORT = process.env.PORT || 3000;
server.listen(PORT, () => {
  console.log(`Server listening on http://localhost:${PORT}`);
});
