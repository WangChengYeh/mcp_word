# MCP Word Add-in

A Model Context Protocol (MCP) server that enables AI-driven document editing workflows in Microsoft Word through a proxy server and Office.js Add-in.

## Overview

This project consists of:
- **MCP Server**: Node.js proxy server that implements MCP protocol and serves the Word Add-in
- **Office Add-in**: Client-side application that runs in Word's task pane and applies AI-generated edits

```mermaid
flowchart LR
  CLI/AI --> Proxy[server.js Proxy Server]
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

3. Configure Claude MCP client:

Add the MCP server to Claude using the transport configuration:

```bash
# Add the Word Editor MCP server
claude mcp add word-editor \
  --transport stdio \
  --command "node" \
  --args "server.js" \
  --cwd "/workspaces/mcp_word"

# Verify the server was added
claude mcp list
```

For custom environment variables:

```bash
claude mcp add word-editor \
  --transport stdio \
  --command "node" \
  --args "server.js" \
  --cwd "/workspaces/mcp_word" \
  --env "PORT=3000,NODE_ENV=production"
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

#### Using Claude MCP commands:

```bash
# Connect to the configured server
claude mcp connect word-editor

# Use the EditTask tool directly
claude mcp call word-editor EditTask \
  --content "Insert a professional summary about AI technology at the beginning of the document."

# Interactive mode with MCP server
claude --mcp word-editor
```

#### Example Claude MCP Commands:

```bash
# Insert text at cursor
claude mcp call word-editor EditTask \
  --content "Add a new paragraph about machine learning applications."

# Replace selected text  
claude mcp call word-editor EditTask \
  --content "Replace the current selection with: Artificial Intelligence is revolutionizing modern business processes."

# Format and style text
claude mcp call word-editor EditTask \
  --content "Make the first paragraph bold and add a professional heading: Executive Summary"
```

#### Remove MCP server configuration:

```bash
# Remove the server configuration
claude mcp remove word-editor
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
- Check that `manifest.xml` points to the correct URL- Ensure Word has internet connectivity**WebSocket connection failed:**- Confirm the server is accessible at `http://localhost:3000`- Check firewall settings- Verify WebSocket support in your environment**MCP client connection issues:**- Ensure the server.js process is running- Check that the MCP client supports the required protocol version- Verify command syntax and parameters### LogsServer logs will show:- MCP protocol messages- WebSocket connections- Edit command processing- Error details## Extensibility### Adding New Edit TypesExtend the `EditTask` tool to support:- Table manipulation- Image insertion- Advanced formatting- Document structure changes### Enhanced Features- Authentication and authorization- Multi-user collaboration- Edit history and versioning- Custom AI model integration## License[Your License Here]## Contributing1. Fork the repository2. Create a feature branch3. Make your changes4. Test with Word Add-in5. Submit a pull request## SupportFor issues and questions:- Check the troubleshooting section- Review server logs- Test with minimal reproduction cases- Report bugs with detailed environment information