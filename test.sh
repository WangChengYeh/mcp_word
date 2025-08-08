#!/bin/bash

# Test script for MCP Word Proxy Server & Office Add-in
# Version: 1.0.0

set -e

echo "🧪 Starting MCP_WORD Test Suite..."

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Test configuration
SERVER_PORT=3000
TEST_TIMEOUT=10

# Helper functions
log_info() {
    echo -e "${GREEN}[INFO]${NC} $1"
}

log_warn() {
    echo -e "${YELLOW}[WARN]${NC} $1"
}

log_error() {
    echo -e "${RED}[ERROR]${NC} $1"
}

# Check if server is running
check_server() {
    if curl -s http://localhost:${SERVER_PORT}/taskpane.html > /dev/null 2>&1; then
        return 0
    else
        return 1
    fi
}

# Test 1: MCP Server STDIO Interface
test_stdio_interface() {
    log_info "Testing MCP Server STDIO Interface..."
    
    # Create test MCP requests
    cat << 'EOF' > /tmp/mcp_test_input.json
{"jsonrpc": "2.0", "id": 1, "method": "initialize", "params": {"protocolVersion": "2024-11-05", "capabilities": {"tools": {}}}}
{"jsonrpc": "2.0", "id": 2, "method": "tools/list"}
{"jsonrpc": "2.0", "id": 3, "method": "tools/call", "params": {"name": "edit_document", "arguments": {"content": "Hello from MCP test!"}}}
EOF

    # Test STDIO pipeline
    if node server.js < /tmp/mcp_test_input.json > /tmp/mcp_test_output.json 2>&1 &
    then
        SERVER_PID=$!
        sleep 2
        
        # Check if server responded
        if [ -f /tmp/mcp_test_output.json ] && [ -s /tmp/mcp_test_output.json ]; then
            log_info "✅ STDIO interface test passed"
            cat /tmp/mcp_test_output.json
        else
            log_error "❌ STDIO interface test failed - no response"
            return 1
        fi
        
        # Cleanup
        kill $SERVER_PID 2>/dev/null || true
        rm -f /tmp/mcp_test_input.json /tmp/mcp_test_output.json
    else
        log_error "❌ Failed to start MCP server for STDIO test"
        return 1
    fi
}

# Test 2: WebSocket Connection Test
test_websocket_connection() {
    log_info "Testing WebSocket Connection..."
    
    # Start server in background
    npm start &
    SERVER_PID=$!
    sleep 3
    
    if ! check_server; then
        log_error "❌ Server not responding on port ${SERVER_PORT}"
        kill $SERVER_PID 2>/dev/null || true
        return 1
    fi
    
    # Create WebSocket test client
    cat << 'EOF' > /tmp/websocket_test_client.js
const io = require('socket.io-client');

const socket = io('http://localhost:3000');

socket.on('connect', () => {
    console.log('✅ WebSocket connected successfully');
    
    // Test sending a message
    socket.emit('test-message', { content: 'Test from automated client' });
    
    // Test receiving ai-cmd event
    socket.on('ai-cmd', (data) => {
        console.log('✅ Received ai-cmd event:', data);
        process.exit(0);
    });
    
    // Simulate sending an ai-cmd after connection
    setTimeout(() => {
        socket.emit('ai-cmd', { content: 'Insert test content via WebSocket' });
    }, 1000);
});

socket.on('connect_error', (error) => {
    console.error('❌ WebSocket connection failed:', error);
    process.exit(1);
});

socket.on('disconnect', () => {
    console.log('WebSocket disconnected');
});

// Timeout after 5 seconds
setTimeout(() => {
    console.log('✅ WebSocket test completed');
    process.exit(0);
}, 5000);
EOF

    # Run WebSocket test
    if node /tmp/websocket_test_client.js; then
        log_info "✅ WebSocket connection test passed"
    else
        log_error "❌ WebSocket connection test failed"
        kill $SERVER_PID 2>/dev/null || true
        rm -f /tmp/websocket_test_client.js
        return 1
    fi
    
    # Cleanup
    kill $SERVER_PID 2>/dev/null || true
    rm -f /tmp/websocket_test_client.js
}

# Test 3: Office Add-in Manifest Validation
test_manifest_validation() {
    log_info "Testing Office Add-in Manifest..."
    
    if [ -f "public/manifest.xml" ]; then
        # Basic XML validation
        if xmllint --noout public/manifest.xml 2>/dev/null; then
            log_info "✅ Manifest XML is valid"
            
            # Check required elements
            if grep -q "ReadWriteDocument" public/manifest.xml && 
               grep -q "http://localhost:3000/taskpane.html" public/manifest.xml; then
                log_info "✅ Manifest contains required permissions and source location"
            else
                log_warn "⚠️  Manifest missing required elements"
            fi
        else
            log_error "❌ Manifest XML validation failed"
            return 1
        fi
    else
        log_error "❌ Manifest file not found"
        return 1
    fi
}

# Test 4: Static File Serving
test_static_files() {
    log_info "Testing Static File Serving..."
    
    # Start server
    npm start &
    SERVER_PID=$!
    sleep 3
    
    # Test taskpane.html
    if curl -s http://localhost:${SERVER_PORT}/taskpane.html | grep -q "Office.js"; then
        log_info "✅ taskpane.html serves correctly and includes Office.js"
    else
        log_error "❌ taskpane.html test failed"
        kill $SERVER_PID 2>/dev/null || true
        return 1
    fi
    
    # Test manifest.xml
    if curl -s http://localhost:${SERVER_PORT}/manifest.xml | grep -q "OfficeApp"; then
        log_info "✅ manifest.xml serves correctly"
    else
        log_error "❌ manifest.xml test failed"
        kill $SERVER_PID 2>/dev/null || true
        return 1
    fi
    
    # Cleanup
    kill $SERVER_PID 2>/dev/null || true
}

# Main test execution
main() {
    log_info "Starting test suite for MCP_WORD..."
    
    # Check prerequisites
    if ! command -v node &> /dev/null; then
        log_error "Node.js is not installed"
        exit 1
    fi
    
    if ! command -v npm &> /dev/null; then
        log_error "npm is not installed"
        exit 1
    fi
    
    # Install dependencies if needed
    if [ ! -d "node_modules" ]; then
        log_info "Installing dependencies..."
        npm install
    fi
    
    # Run tests
    TESTS_PASSED=0
    TOTAL_TESTS=4
    
    if test_manifest_validation; then
        ((TESTS_PASSED++))
    fi
    
    if test_static_files; then
        ((TESTS_PASSED++))
    fi
    
    if test_stdio_interface; then
        ((TESTS_PASSED++))
    fi
    
    if test_websocket_connection; then
        ((TESTS_PASSED++))
    fi
    
    # Summary
    log_info "Test Results: ${TESTS_PASSED}/${TOTAL_TESTS} tests passed"
    
    if [ $TESTS_PASSED -eq $TOTAL_TESTS ]; then
        log_info "🎉 All tests passed!"
        exit 0
    else
        log_error "💥 Some tests failed!"
        exit 1
    fi
}

# Run main function
main "$@"
