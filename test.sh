#!/bin/bash

# Colors for output
GREEN='\033[0;32m'
RED='\033[0;31m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# Server configuration
SERVER_URL="http://localhost:3000"
MCP_ENDPOINT="$SERVER_URL/mcp"

echo -e "${YELLOW}=== MCP Word Add-in Test Script ===${NC}"
echo "Server: $SERVER_URL"
echo "MCP Endpoint: $MCP_ENDPOINT"
echo

# Function to test server health
test_health() {
    echo -e "${YELLOW}Testing server health...${NC}"
    response=$(curl -s -w "HTTP_STATUS:%{http_code}" "$SERVER_URL/health")
    http_status=$(echo $response | tr -d '\n' | sed -e 's/.*HTTP_STATUS://')
    
    if [ "$http_status" -eq 200 ]; then
        echo -e "${GREEN}✓ Server is healthy${NC}"
        echo $response | sed -e 's/HTTP_STATUS:.*//' | jq . 2>/dev/null || echo $response | sed -e 's/HTTP_STATUS:.*//'
    else
        echo -e "${RED}✗ Server health check failed (HTTP $http_status)${NC}"
        return 1
    fi
    echo
}

# Function to send MCP EditTask request
test_edit_task() {
    local content="$1"
    local operation="${2:-insert}"
    local position="${3:-cursor}"
    
    echo -e "${YELLOW}Testing EditTask with:${NC}"
    echo "  Content: $content"
    echo "  Operation: $operation"
    echo "  Position: $position"
    
    # Create MCP request payload
    local payload=$(cat <<EOF
{
    "jsonrpc": "2.0",
    "id": "test-$(date +%s)",
    "method": "tools/call",
    "params": {
        "name": "EditTask",
        "arguments": {
            "content": "$content",
            "operation": "$operation",
            "position": "$position"
        }
    }
}
EOF
)
    
    echo "Sending request..."
    response=$(curl -s -w "HTTP_STATUS:%{http_code}" \
        -X POST \
        -H "Content-Type: application/json" \
        -d "$payload" \
        "$MCP_ENDPOINT")
    
    http_status=$(echo $response | tr -d '\n' | sed -e 's/.*HTTP_STATUS://')
    response_body=$(echo $response | sed -e 's/HTTP_STATUS:.*//')
    
    if [ "$http_status" -eq 200 ]; then
        echo -e "${GREEN}✓ Request successful (HTTP $http_status)${NC}"
        echo "$response_body" | jq . 2>/dev/null || echo "$response_body"
    else
        echo -e "${RED}✗ Request failed (HTTP $http_status)${NC}"
        echo "$response_body"
    fi
    echo
}

# Function to test WebSocket endpoint info
test_websocket_info() {
    echo -e "${YELLOW}Testing WebSocket endpoint info...${NC}"
    response=$(curl -s -w "HTTP_STATUS:%{http_code}" "$SERVER_URL")
    http_status=$(echo $response | tr -d '\n' | sed -e 's/.*HTTP_STATUS://')
    
    if [ "$http_status" -eq 200 ]; then
        echo -e "${GREEN}✓ Static server accessible${NC}"
    else
        echo -e "${RED}✗ Static server not accessible (HTTP $http_status)${NC}"
    fi
    echo
}

# Main execution
main() {
    # Check if server is running
    if ! curl -s "$SERVER_URL/health" > /dev/null 2>&1; then
        echo -e "${RED}✗ Server is not running at $SERVER_URL${NC}"
        echo "Please start the server with: npm start"
        exit 1
    fi
    
    # Run tests
    test_health
    test_websocket_info
    
    # Test different EditTask scenarios
    test_edit_task "Hello from MCP test script!" "insert" "cursor"
    test_edit_task "## Test Heading" "insert" "start"
    test_edit_task "This text will be appended." "append" "end"
    test_edit_task "Replacement text for selected content." "replace" "cursor"
    
    echo -e "${YELLOW}=== Test completed ===${NC}"
    echo "Check your Word document for the applied edits."
    echo "Monitor server logs with: npm run debug"
}

# Handle command line arguments
case "$1" in
    "health")
        test_health
        ;;
    "edit")
        if [ -z "$2" ]; then
            echo "Usage: $0 edit \"Your content here\" [operation] [position]"
            exit 1
        fi
        test_edit_task "$2" "$3" "$4"
        ;;
    "help"|"-h"|"--help")
        echo "Usage: $0 [command] [args...]"
        echo
        echo "Commands:"
        echo "  health              - Test server health"
        echo "  edit \"content\"       - Send edit command"
        echo "  help                - Show this help"
        echo
        echo "Examples:"
        echo "  $0                           # Run all tests"
        echo "  $0 health                    # Test server health only"
        echo "  $0 edit \"Hello World!\"       # Send simple edit"
        echo "  $0 edit \"Text\" insert start  # Insert at start"
        ;;
    "")
        main
        ;;
    *)
        echo "Unknown command: $1"
        echo "Use '$0 help' for usage information."
        exit 1
        ;;
esac
