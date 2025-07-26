Office.onReady(function (info) {
    // Initialize the add-in when Office is ready
    console.log("Authopsy commands initialized for:", info.host);
});

// Primary function for ribbon commands
function action(event) {
    try {
        console.log("Authopsy ribbon command executed");
        // For ShowTaskpane actions, just complete the event
        event.completed();
    } catch (error) {
        console.error("Error in ribbon action:", error);
        if (event && event.completed) {
            event.completed();
        }
    }
}

// Compatibility functions
function handleCommand(event) {
    action(event);
}

function onAction(event) {
    action(event);
}

// Make functions globally available
window.action = action;
window.handleCommand = handleCommand;
window.onAction = onAction;
