Office.onReady(function (info) {
    // Initialize the add-in when Office is ready
    console.log("Authopsy add-in commands loaded", info);
    console.log("Office context:", {
        host: info.host,
        platform: info.platform
    });
});

// Function to handle the ribbon button command
function action(event) {
    try {
        console.log("Authopsy ribbon command executed");
        console.log("Event object:", event);
        // The button action is handled by ShowTaskpane, so we just complete
        event.completed();
    } catch (error) {
        console.error("Error in action function:", error);
        if (event && event.completed) {
            event.completed();
        }
    }
}

// Legacy function name for compatibility
function handleCommand(event) {
    console.log("handleCommand called, forwarding to action");
    action(event);
}

// Additional function names that might be expected
function onAction(event) {
    console.log("onAction called, forwarding to action");
    action(event);
}

// Make functions available globally
window.action = action;
window.handleCommand = handleCommand;
window.onAction = onAction;

// Add global error handler
window.onerror = function(msg, url, lineNo, columnNo, error) {
    console.error('Command file error:', {
        message: msg,
        source: url,
        line: lineNo,
        column: columnNo,
        error: error
    });
    return false;
};
