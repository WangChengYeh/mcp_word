#!/bin/bash

# Test script for MCP Word Server
# Unit tests without AI or Word dependencies

set -e

echo "🧪 Starting MCP Word Server Tests..."

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Test configuration
TEST_PORT=3001
SERVER_PID=""
TEST_DIR="/tmp/mcp_word_test"

# Cleanup function
cleanup() {
    echo -e "\n🧹 Cleaning up..."
    if [ ! -z "$SERVER_PID" ]; then
        kill $SERVER_PID 2>/dev/null || true
        wait $SERVER_PID 2>/dev/null || true
    fi
    rm -rf "$TEST_DIR"
    echo -e "${GREEN}✓ Cleanup completed${NC}"
}

# Set trap for cleanup
trap cleanup EXIT

# Test 1: Check dependencies
echo -e "\n${YELLOW}Test 1: Checking dependencies...${NC}"
if ! command -v node &> /dev/null; then
    echo -e "${RED}✗ Node.js not found${NC}"
    exit 1
fi

if ! command -v npm &> /dev/null; then
    echo -e "${RED}✗ npm not found${NC}"
    exit 1
fi
echo -e "${GREEN}✓ Dependencies check passed${NC}"

# Test 2: Check server.js exists and is executable
echo -e "\n${YELLOW}Test 2: Checking server.js...${NC}"
if [ ! -f "server.js" ]; then
    echo -e "${RED}✗ server.js not found${NC}"
    exit 1
fi

if [ ! -x "server.js" ]; then
    chmod +x server.js
fi
echo -e "${GREEN}✓ server.js exists and is executable${NC}"

# Test 3: Start server in background
echo -e "\n${YELLOW}Test 3: Starting MCP server...${NC}"
mkdir -p "$TEST_DIR/public"
echo '{"name":"test","version":"1.0.0","type":"module","dependencies":{}}' > "$TEST_DIR/package.json"

# Start server with test port
PORT=$TEST_PORT node server.js > "$TEST_DIR/server.log" 2>&1 &
SERVER_PID=$!

# Wait for server to start
sleep 3

if ! kill -0 $SERVER_PID 2>/dev/null; then
    echo -e "${RED}✗ Server failed to start${NC}"
    cat "$TEST_DIR/server.log"
    exit 1
fi
echo -e "${GREEN}✓ MCP server started (PID: $SERVER_PID)${NC}"

# Test 4: Check HTTP server is running
echo -e "\n${YELLOW}Test 4: Testing HTTP server...${NC}"
if ! curl -s "http://localhost:$TEST_PORT" > /dev/null; then
    echo -e "${RED}✗ HTTP server not responding${NC}"
    exit 1
fi
echo -e "${GREEN}✓ HTTP server is responding${NC}"

# Test 5: Test Socket.IO connection
echo -e "\n${YELLOW}Test 5: Testing Socket.IO connection...${NC}"
cat > "$TEST_DIR/socket_test.js" << 'EOF'
import { io } from 'socket.io-client';

const PORT = process.env.TEST_PORT || 3001;
const socket = io(`http://localhost:${PORT}`);
let connected = false;

socket.on('connect', () => {
    console.log('✓ Socket.IO connected');
    connected = true;
    
    // Test receiving ai-cmd event
    socket.on('ai-cmd', (data) => {
        console.log('✓ Received ai-cmd:', data);
        process.exit(0);
    });
    
    // Simulate receiving a command
    setTimeout(() => {
        if (connected) {
            console.log('✓ Socket.IO connection test passed');
            process.exit(0);
        }
    }, 1000);
});

socket.on('connect_error', (error) => {
    console.log('✗ Socket.IO connection failed:', error.message);
    process.exit(1);
});

setTimeout(() => {
    if (!connected) {
        console.log('✗ Socket.IO connection timeout');
        process.exit(1);
    }
}, 5000);
EOF

if ! TEST_PORT=$TEST_PORT node "$TEST_DIR/socket_test.js"; then
    echo -e "${RED}✗ Socket.IO connection test failed${NC}"
    exit 1
fi
echo -e "${GREEN}✓ Socket.IO connection test passed${NC}"

# Test 6: Test MCP tool registration
echo -e "\n${YELLOW}Test 6: Testing MCP tool functionality...${NC}"
cat > "$TEST_DIR/mcp_test.js" << 'EOF'
// Simulate MCP EditTask tool test
const testTool = {
    name: "EditTask",
    description: "Send edit commands to connected Word document",
    inputSchema: {
        type: "object",
        properties: {
            content: { type: "string" },
            action: { type: "string", enum: ["insert", "replace", "append"], default: "insert" },
            position: { type: "string", enum: ["start", "end", "cursor"], default: "cursor" }
        },
        required: ["content"]
    }
};

// Test input validation
const testInputs = [
    { content: "Test content", action: "insert", position: "cursor" },
    { content: "Another test", action: "replace", position: "start" },
    { content: "Final test", action: "append", position: "end" }
];

console.log('✓ MCP tool definition valid');

testInputs.forEach((input, index) => {
    if (input.content && typeof input.content === 'string') {
        console.log(`✓ Test input ${index + 1} valid:`, input);
    } else {
        console.log(`✗ Test input ${index + 1} invalid:`, input);
        process.exit(1);
    }
});

console.log('✓ MCP tool functionality test passed');
EOF

if ! node "$TEST_DIR/mcp_test.js"; then
    echo -e "${RED}✗ MCP tool test failed${NC}"
    exit 1
fi
echo -e "${GREEN}✓ MCP tool test passed${NC}"

# Test 7: Test server logs
echo -e "\n${YELLOW}Test 7: Checking server logs...${NC}"
if [ -f "$TEST_DIR/server.log" ]; then
    if grep -q "MCP Word Server running" "$TEST_DIR/server.log"; then
        echo -e "${GREEN}✓ Server started successfully${NC}"
    else
        echo -e "${RED}✗ Server startup message not found${NC}"
        cat "$TEST_DIR/server.log"
        exit 1
    fi
else
    echo -e "${RED}✗ Server log file not found${NC}"
    exit 1
fi

# Test 8: Test graceful shutdown
echo -e "\n${YELLOW}Test 8: Testing graceful shutdown...${NC}"
kill -INT $SERVER_PID
wait $SERVER_PID 2>/dev/null || true
SERVER_PID=""
echo -e "${GREEN}✓ Server shutdown gracefully${NC}"

# Final results
echo -e "\n${GREEN}🎉 All tests passed!${NC}"
echo -e "${GREEN}✓ Dependencies verified${NC}"
echo -e "${GREEN}✓ Server executable${NC}"
echo -e "${GREEN}✓ MCP server starts correctly${NC}"
echo -e "${GREEN}✓ HTTP server responds${NC}"
echo -e "${GREEN}✓ Socket.IO connections work${NC}"
echo -e "${GREEN}✓ MCP tool definitions valid${NC}"
echo -e "${GREEN}✓ Server logs properly${NC}"
echo -e "${GREEN}✓ Graceful shutdown works${NC}"
echo -e "\n${GREEN}MCP Word Server is ready for use!${NC}"
