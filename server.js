import path from 'path';
import { McpServer } from '@modelcontextprotocol/typescript-sdk';

// Initialize MCP Server
const staticDir = path.join(process.cwd(), 'public');
const port = process.env.PORT || 3000;
const mcpServer = new McpServer({
  apiKey: process.env.CLAUDE_API_KEY,
  staticDir,
});

// Register EditTask handler
mcpServer.on('EditTask', async ({ content }) => {
  // Process the edit request via SDK
  return await mcpServer.requestEditTask({ content });
});

// Start the MCP server
mcpServer.start(port)
  .then(() => console.log(`MCP Server running at http://localhost:${port}`))
  .catch((err) => console.error('Failed to start MCP Server:', err));
