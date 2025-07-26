Office.onReady(function () {
    // Initialize the add-in when Office is ready
    console.log("Authopsy add-in commands loaded");
});

// Function to handle any future command actions
function handleCommand(event) {
    try {
        // Add any command handling logic here if needed
        console.log("Command executed");
        event.completed();
    } catch (error) {
        console.error("Error in command handler:", error);
        event.completed();
    }
}
