Office.onReady(function () {
    // Initialize the add-in when Office is ready
    console.log("Authopsy add-in commands loaded");
});

// Function to handle any command actions
function handleCommand(event) {
    try {
        console.log("Command executed");
        event.completed();
    } catch (error) {
        console.error("Error in command handler:", error);
        event.completed();
    }
}

// Make functions available globally
window.handleCommand = handleCommand;
