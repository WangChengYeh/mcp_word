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
  // Use same-origin Socket.IO by default to match manifest SourceLocation
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
  
  // Listen for tool-named events (editTask)
  socket.on('editTask', function(editCommand) {
    console.log('Received editTask:', editCommand);
    handleEditCommand(editCommand);
  });
  
  socket.on('connect_error', function(error) {
    console.error('Connection error:', error);
    updateStatus('Connection failed');
  });
}

async function handleEditCommand(editCommand) {
  // Align with SPEC: { content, action: 'insert'|'replace'|'append', target: 'cursor'|'selection'|'document' }
  const { taskId, content, action = 'insert', target = 'selection' } = editCommand || {};
  
  try {
    await Word.run(async (context) => {
      let range;
      const body = context.document.body;
      
      // Determine insertion target
      switch (target) {
        case 'document':
          // For document-level actions, prefer end insertion; replace clears body first
          if (action === 'replace') {
            body.clear();
            // Insert at start after clearing
            body.insertParagraph(content, Word.InsertLocation.start);
            await context.sync();
            emitComplete(`Replaced entire document`);
            return;
          } else if (action === 'append' || action === 'insert') {
            range = body.getRange('End');
          }
          break;
        case 'cursor':
        case 'selection':
        default:
          range = context.document.getSelection();
          break;
      }
      
      // Perform edit operation
      switch (action) {
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
          throw new Error(`Unsupported action: ${action}`);
      }
      
      await context.sync();
      
      // Send success response
      socket.emit('edit-complete', {
        taskId,
        success: true,
        message: `Successfully ${action} on ${target}`,
        timestamp: new Date().toISOString()
      });
      
      console.log(`Edit completed: ${taskId || 'no-id'}`);
      updateStatus(`Edit completed: ${action}`);
    });
    
  } catch (error) {
    console.error('Edit operation failed:', error);
    
    // Send error response
    socket.emit('edit-error', {
      taskId,
      error: error.message || 'Unknown error occurred',
      timestamp: new Date().toISOString()
    });
    
    updateStatus(`Edit failed: ${error.message}`);
  }

  function emitComplete(message) {
    socket.emit('edit-complete', {
      taskId,
      success: true,
      message,
      timestamp: new Date().toISOString(),
    });
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
