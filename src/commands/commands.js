Office.onReady(function () {
    // Initialize the add-in when Office is ready
    console.log("Authopsy add-in commands loaded");
    
    // Check authentication when Office is ready
    checkAuthenticationStatus();
});

// Check authentication status and update dropdown labels
function checkAuthenticationStatus() {
    try {
        const item = Office.context.mailbox.item;
        
        if (item && item.internetHeaders) {
            item.internetHeaders.getAsync(['Authentication-Results', 'ARC-Authentication-Results'], function(result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    const authResults = analyzeAuthHeaders(result.value);
                    updateDropdownLabels(authResults);
                } else {
                    console.log("Could not retrieve authentication headers");
                    updateDropdownLabels({ dmarc: 'unknown', dkim: 'unknown', spf: 'unknown' });
                }
            });
        } else {
            console.log("No internetHeaders API available");
            updateDropdownLabels({ dmarc: 'unknown', dkim: 'unknown', spf: 'unknown' });
        }
    } catch (error) {
        console.error("Error checking authentication:", error);
        updateDropdownLabels({ dmarc: 'fail', dkim: 'fail', spf: 'fail' });
    }
}

// Update dropdown menu labels with authentication results
function updateDropdownLabels(results) {
    const getStatusIcon = (status) => {
        switch(status) {
            case 'pass': return '✅';
            case 'fail': return '❌';
            case 'unknown': return '❓';
            default: return '⚠️';
        }
    };
    
    // Update the ribbon menu items (this would require additional API support)
    console.log("Authentication Results:", {
        dmarc: `${getStatusIcon(results.dmarc)} DMARC: ${results.dmarc}`,
        dkim: `${getStatusIcon(results.dkim)} DKIM: ${results.dkim}`,
        spf: `${getStatusIcon(results.spf)} SPF: ${results.spf}`
    });
}

// Analyze authentication headers
function analyzeAuthHeaders(headers) {
    const results = { dmarc: 'fail', dkim: 'fail', spf: 'fail' };
    
    for (const header of headers) {
        const value = header.value.toLowerCase();
        
        if (value.includes('dmarc=pass')) results.dmarc = 'pass';
        if (value.includes('dkim=pass')) results.dkim = 'pass';
        if (value.includes('spf=pass')) results.spf = 'pass';
    }
    
    return results;
}

// Function to show DMARC details
function showDmarcDetails(event) {
    try {
        Office.context.mailbox.item.notificationMessages.addAsync("dmarc-details", {
            type: "informationalMessage",
            message: "DMARC: Domain-based Message Authentication, Reporting & Conformance - checks if email aligns with domain policy",
            icon: "Icon.16x16",
            persistent: false
        });
        event.completed();
    } catch (error) {
        console.error("Error showing DMARC details:", error);
        event.completed();
    }
}

// Function to show DKIM details
function showDkimDetails(event) {
    try {
        Office.context.mailbox.item.notificationMessages.addAsync("dkim-details", {
            type: "informationalMessage",
            message: "DKIM: DomainKeys Identified Mail - verifies email hasn't been tampered with using cryptographic signatures",
            icon: "Icon.16x16",
            persistent: false
        });
        event.completed();
    } catch (error) {
        console.error("Error showing DKIM details:", error);
        event.completed();
    }
}

// Function to show SPF details
function showSpfDetails(event) {
    try {
        Office.context.mailbox.item.notificationMessages.addAsync("spf-details", {
            type: "informationalMessage",
            message: "SPF: Sender Policy Framework - verifies the sending server is authorized to send email for this domain",
            icon: "Icon.16x16",
            persistent: false
        });
        event.completed();
    } catch (error) {
        console.error("Error showing SPF details:", error);
        event.completed();
    }
}

// Function to handle any future command actions
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
window.checkAuthenticationStatus = checkAuthenticationStatus;
window.showDmarcDetails = showDmarcDetails;
window.showDkimDetails = showDkimDetails;
window.showSpfDetails = showSpfDetails;
window.handleCommand = handleCommand;
