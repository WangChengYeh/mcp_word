# MCP Word Add-in

A Model Context Protocol (MCP) server that enables AI-driven document editing workflows in Microsoft Word through a proxy server and Office.js Add-in.

## Overview

This project consists of:
- **MCP Server**: Node.js proxy server that implements MCP protocol and serves the Word Add-in
- **Office Add-in**: Client-side application that runs in Word's task pane and applies AI-generated edits

```mermaid
flowchart LR
  CLI/AI --> Proxy[Proxy Server (server.js)]
  Proxy --> Browser[Office.js Task Pane]
  Browser --> Word[Word Document]
```

## Features

- Real-time document editing through AI commands
- WebSocket-based communication between server and Add-in
- MCP protocol compliance for integration with AI tools
- Office.js integration for seamless Word document manipulation

## Prerequisites

- Node.js 18+ 
- Microsoft Word (Desktop or Online)
- Claude CLI or compatible MCP client

## Installation

1. Clone the repository:
```bash
git clone <repository-url>
cd mcp_word
```

2. Install dependencies:
```bash
npm install
```

## Usage

### 1. Start the MCP Server

```bash
npm start
```

The server will start on `http://localhost:3000` by default.

### 2. Sideload the Word Add-in

#### For Word Desktop:
1. Open Word
2. Go to Insert > My Add-ins > Upload My Add-in
3. Select the `public/manifest.xml` file
4. The Add-in will appear in your task pane

#### For Word Online:
1. Open Word Online
2. Go to Insert > Office Add-ins > Upload My Add-in
3. Select the `public/manifest.xml` file

### 3. Connect MCP Client

Use Claude CLI or another MCP-compatible client to send edit commands:

```bash
# Example using Claude CLI
claude-mcp-client --server "node server.js" --tool EditTask --args '{"content": "Insert a professional summary about AI technology."}'
```

### 4. Document Editing

The Add-in will automatically:
- Receive edit commands via WebSocket
- Apply changes to the active Word document
- Handle text insertion, replacement, and formatting

## MCP Server API

### Tools

#### EditTask
Processes document edit requests and forwards them to the Word Add-in.

**Parameters:**
- `content` (string): The text content to insert or edit instructions

**Example:**
```json
{
  "content": "Replace the first paragraph with a professional introduction about machine learning."
}
```

## Development

### Project Structure

```
mcp_word/
├── server.js           # MCP server and Express static server
├── public/
│   ├── manifest.xml    # Office Add-in manifest
│   ├── taskpane.html   # Add-in UI
│   └── taskpane.js     # Add-in logic and WebSocket client
├── package.json
├── SPEC.md            # Technical specification
└── README.md          # This file
```

### Key Components

- **server.js**: Implements MCP protocol using `@modelcontextprotocol/sdk`
- **taskpane.js**: Handles Office.js integration and WebSocket communication
- **manifest.xml**: Defines Add-in metadata and permissions

### Configuration

The server accepts the following environment variables:

- `PORT`: Server port (default: 3000)
- `HOST`: Server host (default: localhost)

### Debugging

1. **Server logs**: Check console output for MCP and WebSocket events
2. **Add-in debugging**: Use browser dev tools in Word (F12)
3. **Office.js errors**: Monitor the task pane console for Office API issues

## Troubleshooting

### Common Issues

**Add-in not loading:**
- Verify the server is running on port 3000
- Check that `manifest.xml` points to the correct URL
- Ensure Word has internet connectivity

**WebSocket connection failed:**
- Confirm the server is accessible at `http://localhost:3000`
- Check firewall settings
- Verify WebSocket support in your environment

**MCP client connection issues:**
- Ensure the server.js process is running
- Check that the MCP client supports the required protocol version
- Verify command syntax and parameters

### Logs

Server logs will show:
- MCP protocol messages
- WebSocket connections
- Edit command processing
- Error details

## Extensibility

### Adding New Edit Types

Extend the `EditTask` tool to support:
- Table manipulation
- Image insertion
- Advanced formatting
- Document structure changes

### Enhanced Features

- Authentication and authorization
- Multi-user collaboration
- Edit history and versioning
- Custom AI model integration

## License

[Your License Here]

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test with Word Add-in
5. Submit a pull request

## Support

For issues and questions:
- Check the troubleshooting section
- Review server logs
- Test with minimal reproduction cases
- Report bugs with detailed environment information
