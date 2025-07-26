// Aggressive debug logging at start
console.log("🔧 AUTHOPSY: Commands.js LOADING - " + new Date().toISOString());
alert("AUTHOPSY: Commands.js loaded!"); // Visible alert to confirm loading

Office.onReady(function (info) {
    // Initialize the add-in when Office is ready
    console.log("🔧 AUTHOPSY: Office.onReady called with:", info);
    alert("AUTHOPSY: Office.onReady called! Host: " + info.host); // Visible confirmation
    console.log("🔧 AUTHOPSY: Office context:", {
        host: info.host,
        platform: info.platform
    });
});

// Function to handle the ribbon button command
function action(event) {
    try {
        console.log("🔧 AUTHOPSY: action() function called");
        console.log("🔧 AUTHOPSY: Event object:", event);
        alert("AUTHOPSY: action() called!"); // Visible confirmation
        // The button action is handled by ShowTaskpane, so we just complete
        event.completed();
    } catch (error) {
        console.error("🔧 AUTHOPSY: Error in action function:", error);
        alert("AUTHOPSY ERROR: " + error.message);
        if (event && event.completed) {
            event.completed();
        }
    }
}

// Legacy function name for compatibility
function handleCommand(event) {
    console.log("🔧 AUTHOPSY: handleCommand called, forwarding to action");
    action(event);
}

// Additional function names that might be expected
function onAction(event) {
    console.log("🔧 AUTHOPSY: onAction called, forwarding to action");
    action(event);
}

// Make functions available globally
window.action = action;
window.handleCommand = handleCommand;
window.onAction = onAction;

// Debug: Log all available global functions
console.log("🔧 AUTHOPSY: Available window functions:", {
    action: typeof window.action,
    handleCommand: typeof window.handleCommand,
    onAction: typeof window.onAction
});

// Add global error handler
window.onerror = function(msg, url, lineNo, columnNo, error) {
    console.error('🔧 AUTHOPSY: Command file error:', {
        message: msg,
        source: url,
        line: lineNo,
        column: columnNo,
        error: error
    });
    alert("AUTHOPSY GLOBAL ERROR: " + msg);
    return false;
};

console.log("🔧 AUTHOPSY: Commands.js setup complete");
