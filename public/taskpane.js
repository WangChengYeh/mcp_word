Office.onReady(function(info) {
  if (info.host === Office.HostType.Word) {
    console.log('Word Add-in loaded successfully');
    initializeAddIn();
  } else {
    console.error('This Add-in is designed for Microsoft Word');
  }
});

let socket;
let isConnected = false;

function initializeAddIn() {
  // Establish WebSocket connection
  socket = io("https://localhost:3000", {transports: ["websocket"]});
  
  socket.on('connect', function() {
    console.log('Connected to MCP proxy server');
    isConnected = true;
    
    // Send initial status
    socket.emit('status', {
      type: 'connection',
      message: 'Office Add-in connected',
      host: 'Word',
      timestamp: new Date().toISOString()
    });
    
    updateStatus('Connected to server');
  });
  
  socket.on('disconnect', function() {
    console.log('Disconnected from MCP proxy server');
    isConnected = false;
    updateStatus('Disconnected from server');
  });
  
  // Listen for AI command events
  socket.on('ai-cmd', function(editCommand) {
    console.log('Received edit command:', editCommand);
    handleEditCommand(editCommand);
  });
  
  socket.on('connect_error', function(error) {
    console.error('Connection error:', error);
    updateStatus('Connection failed');
  });
}

async function handleEditCommand(editCommand) {
  const { editId, content, operation = 'insert', position = 'cursor' } = editCommand;
  
  try {
    await Word.run(async (context) => {
      let range;
      
      // Determine insertion point based on position
      switch (position) {
        case 'start':
          range = context.document.body.getRange('Start');
          break;
        case 'end':
          range = context.document.body.getRange('End');
          break;
        case 'cursor':
        default:
          // Use selection or cursor position
          range = context.document.getSelection();
          break;
      }
      
      // Perform edit operation
      switch (operation) {
        case 'insert':
          range.insertText(content, Word.InsertLocation.after);
          break;
        case 'replace':
          range.insertText(content, Word.InsertLocation.replace);
          break;
        case 'append':
          range.insertText('\n' + content, Word.InsertLocation.after);
          break;
        default:
          throw new Error(`Unsupported operation: ${operation}`);
      }
      
      await context.sync();
      
      // Send success response
      socket.emit('edit-complete', {
        editId,
        success: true,
        message: `Successfully ${operation}ed text at ${position}`,
        timestamp: new Date().toISOString()
      });
      
      console.log(`Edit operation completed: ${editId}`);
      updateStatus(`Edit completed: ${operation}`);
    });
    
  } catch (error) {
    console.error('Edit operation failed:', error);
    
    // Send error response
    socket.emit('edit-error', {
      editId,
      error: error.message || 'Unknown error occurred',
      timestamp: new Date().toISOString()
    });
    
    updateStatus(`Edit failed: ${error.message}`);
  }
}

function updateStatus(message) {
  const statusElement = document.getElementById('status');
  if (statusElement) {
    statusElement.textContent = `Status: ${message}`;
    statusElement.title = new Date().toISOString();
  }
  console.log(`Status update: ${message}`);
}

// Auto-start behavior - initialize when DOM is loaded
document.addEventListener('DOMContentLoaded', function() {
  console.log('DOM loaded, waiting for Office.js...');
  
  // Add status display
  const statusDiv = document.createElement('div');
  statusDiv.id = 'status';
  statusDiv.style.cssText = 'padding: 10px; background: #f0f0f0; border: 1px solid #ddd; margin: 10px;';
  statusDiv.textContent = 'Initializing...';
  document.body.appendChild(statusDiv);
  
  // Add connection indicator
  const indicator = document.createElement('div');
  indicator.style.cssText = 'width: 10px; height: 10px; border-radius: 50%; display: inline-block; margin-right: 5px; background: red;';
  indicator.id = 'connection-indicator';
  statusDiv.prepend(indicator);
  
  // Update indicator color based on connection status
  setInterval(() => {
    const indicator = document.getElementById('connection-indicator');
    if (indicator) {
      indicator.style.background = isConnected ? 'green' : 'red';
    }
  }, 1000);
});

// Export functions for potential manual testing
window.mcpWordAddin = {
  handleEditCommand,
  updateStatus,
  getConnectionStatus: () => isConnected
};
