Office.onReady(function () {
    // Initialize the add-in when Office is ready
    console.log("Authopsy add-in commands loaded");
});

// Function to handle message read events - automatically triggered
function onMessageRead(event) {
    try {
        console.log("📧 Message read event triggered - checking authentication");
        
        // Get the current item (email)
        const item = Office.context.mailbox.item;
        
        if (item && item.internetHeaders) {
            // Check authentication headers
            item.internetHeaders.getAsync(['Authentication-Results', 'ARC-Authentication-Results'], function(result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("🔍 Authentication headers retrieved");
                    
                    // Analyze authentication results
                    const authResults = analyzeAuthHeaders(result.value);
                    
                    // Show notification with results
                    showAuthNotification(authResults);
                } else {
                    console.log("⚠️ Could not retrieve authentication headers");
                    showAuthNotification({ dmarc: 'unknown', dkim: 'unknown', spf: 'unknown' });
                }
                
                // Complete the event
                event.completed();
            });
        } else {
            console.log("⚠️ No item or internetHeaders API available");
            event.completed();
        }
    } catch (error) {
        console.error("❌ Error in onMessageRead:", error);
        event.completed();
    }
}

// Analyze authentication headers
function analyzeAuthHeaders(headers) {
    const results = { dmarc: 'fail', dkim: 'fail', spf: 'fail' };
    
    // Look for authentication results in headers
    for (const header of headers) {
        const value = header.value.toLowerCase();
        
        if (value.includes('dmarc=pass')) results.dmarc = 'pass';
        if (value.includes('dkim=pass')) results.dkim = 'pass';
        if (value.includes('spf=pass')) results.spf = 'pass';
    }
    
    return results;
}

// Show authentication notification
function showAuthNotification(results) {
    const passCount = Object.values(results).filter(r => r === 'pass').length;
    const status = passCount === 3 ? '✅ Secure' : passCount >= 1 ? '⚠️ Partial' : '❌ Insecure';
    
    // Create notification message
    const message = `Email Authentication: ${status} | DMARC: ${results.dmarc === 'pass' ? '✓' : '✗'} | DKIM: ${results.dkim === 'pass' ? '✓' : '✗'} | SPF: ${results.spf === 'pass' ? '✓' : '✗'}`;
    
    // Show notification (this would appear as a banner/info bar)
    if (Office.context.mailbox.item && Office.context.mailbox.item.notificationMessages) {
        Office.context.mailbox.item.notificationMessages.addAsync("authopsy-results", {
            type: "informationalMessage",
            message: message,
            icon: "Icon.16x16",
            persistent: true
        });
    }
    
    console.log("📊 Authentication notification shown:", message);
}

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

// Make functions available globally
window.onMessageRead = onMessageRead;
window.handleCommand = handleCommand;
