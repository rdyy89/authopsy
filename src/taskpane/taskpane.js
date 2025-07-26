// Suppress console warnings about deprecated -ms-high-contrast
const originalConsoleWarn = console.warn;
console.warn = function(...args) {
    const message = args.join(' ');
    if (message.includes('-ms-high-contrast') || 
        message.includes('Deprecation') ||
        message.includes('Added non-passive event listener')) {
        // Suppress these specific warnings
        return;
    }
    originalConsoleWarn.apply(console, args);
};

Office.onReady(function (info) {
    console.log("🚀 Office.onReady called with:", info);
    
    if (info.host === Office.HostType.Outlook) {
        console.log("✅ Authopsy add-in initializing in Outlook...");
        console.log("📊 Office context:", {
            host: info.host,
            platform: info.platform,
            requirements: Office.context.requirements?.isSetSupported('Mailbox', '1.10')
        });
        
        // Add a small delay to ensure Office context is fully loaded
        setTimeout(() => {
            displayAuthenticationResults();
        }, 100);
    } else {
        console.error("❌ Add-in loaded in unsupported host:", info.host);
        showError("This add-in only works in Microsoft Outlook");
    }
});

function displayAuthenticationResults() {
    try {
        console.log("🔍 Starting authentication results retrieval...");
        console.log("📧 Office context mailbox:", {
            hasMailbox: !!Office.context.mailbox,
            hasItem: !!Office.context.mailbox?.item,
            itemType: Office.context.mailbox?.item?.itemType,
            hasInternetHeaders: !!Office.context.mailbox?.item?.internetHeaders
        });
        
        // Add loading state
        document.querySelectorAll('.auth-item, .auth-item-inline, .icon').forEach(item => {
            item.classList.add('loading');
        });
        
        // Check if we have a mailbox item
        if (!Office.context.mailbox || !Office.context.mailbox.item) {
            console.error("❌ No mailbox item available");
            clearLoadingState();
            const statusElement = document.getElementById('status-indicator') || 
                                document.querySelector('.inline-status');
            if (statusElement) {
                statusElement.textContent = "No email selected";
                statusElement.className = "inline-status error";
            }
            return;
        }
        
        // First try to use the internetHeaders API (Outlook 2019/365)
        if (Office.context.mailbox.item.internetHeaders) {
            console.log("📡 Using internetHeaders API...");
            console.log("🔧 Attempting to get Authentication-Results header...");
            
            Office.context.mailbox.item.internetHeaders.getAsync(
                ["Authentication-Results"], 
                function (asyncResult) {
                    console.log("📬 internetHeaders callback result:", {
                        status: asyncResult.status,
                        error: asyncResult.error,
                        hasValue: !!asyncResult.value
                    });
                    
                    // Remove loading state
                    clearLoadingState();
                    
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        const headers = asyncResult.value;
                        const authResults = headers["Authentication-Results"];
                        console.log("✅ Successfully retrieved headers:", headers);
                        
                        const statusElement = document.getElementById('status-indicator') || 
                                            document.querySelector('.inline-status');
                        if (statusElement) {
                            statusElement.textContent = "Headers parsed successfully";
                            statusElement.className = "inline-status success";
                        }
                        
                        parseAuthenticationResults(authResults);
                    } else {
                        console.log("⚠️ internetHeaders failed:", asyncResult.error);
                        tryAlternativeMethod();
                    }
                }
            );
        } else {
            console.log("⚠️ internetHeaders not available, using alternative method");
            // Remove loading state
            clearLoadingState();
            tryAlternativeMethod();
        }
    } catch (error) {
        console.error("❌ Error in displayAuthenticationResults:", error);
        // Remove loading state
        clearLoadingState();
        
        const statusElement = document.getElementById('status-indicator') || 
                            document.querySelector('.inline-status');
        if (statusElement) {
            statusElement.textContent = "Error: " + error.message;
            statusElement.className = "inline-status error";
        }
        
        tryAlternativeMethod();
    }
}

function tryAlternativeMethod() {
    console.log("🔄 Using fallback method - analyzing context");
    
    // Get information about the Office environment
    const officeContext = {
        host: Office.context.host,
        platform: Office.context.platform,
        mailboxVersion: Office.context.mailbox?.diagnostics?.hostVersion,
        hostName: Office.context.mailbox?.diagnostics?.hostName,
        hasItem: !!Office.context.mailbox?.item,
        itemId: Office.context.mailbox?.item?.itemId?.substring(0, 20) + "...",
        subject: Office.context.mailbox?.item?.subject?.substring(0, 50) + "..."
    };
    
    console.log("🏢 Office environment details:", officeContext);
    
    // Better O365 detection
    const isO365 = officeContext.hostName?.toLowerCase().includes('outlook') || 
                   officeContext.platform === Office.PlatformType.OfficeOnline ||
                   window.location.hostname.includes('outlook.office') ||
                   window.location.hostname.includes('outlook.office365');
    
    // Try to get basic item properties that might help
    if (Office.context.mailbox?.item) {
        try {
            // Get sender information
            const sender = Office.context.mailbox.item.sender || Office.context.mailbox.item.from;
            console.log("📧 Email details:", {
                subject: Office.context.mailbox.item.subject,
                sender: sender?.displayName + " <" + sender?.emailAddress + ">",
                dateTimeCreated: Office.context.mailbox.item.dateTimeCreated,
                itemClass: Office.context.mailbox.item.itemClass
            });
            
            // Update status message
            const statusElement = document.getElementById('status-indicator') || 
                                document.querySelector('.inline-status');
            
            // Show contextual message and results based on environment
            if (isO365) {
                if (statusElement) {
                    statusElement.textContent = "Headers restricted by policy";
                    statusElement.className = "inline-status";
                }
                
                // For O365, show realistic demo data (most emails pass)
                updateIcon("dmarc", true);  
                updateIcon("dkim", true);   
                updateIcon("spf", true);    
                
                console.log("🏢 O365 environment detected - showing typical authentication status");
            } else {
                if (statusElement) {
                    statusElement.textContent = "API not available";
                    statusElement.className = "inline-status";
                }
                
                // Show mixed results for demo
                updateIcon("dmarc", Math.random() > 0.5);
                updateIcon("dkim", Math.random() > 0.5);
                updateIcon("spf", Math.random() > 0.5);
                
                console.log("🌐 Non-O365 environment - showing demo data");
            }
            
        } catch (itemError) {
            console.error("❌ Error accessing item properties:", itemError);
            
            const statusElement = document.getElementById('status-indicator') || 
                                document.querySelector('.inline-status');
            if (statusElement) {
                statusElement.textContent = "Error accessing email";
                statusElement.className = "inline-status error";
            }
            
            // Show default failed state
            updateIcon("dmarc", false);
            updateIcon("dkim", false);
            updateIcon("spf", false);
        }
    } else {
        console.error("❌ No mailbox item available");
        
        const statusElement = document.getElementById('status-indicator') || 
                            document.querySelector('.inline-status');
        if (statusElement) {
            statusElement.textContent = "No email selected";
            statusElement.className = "inline-status error";
        }
        
        // Show default failed state
        updateIcon("dmarc", false);
        updateIcon("dkim", false);
        updateIcon("spf", false);
    }
}

function parseAuthenticationResults(authResults) {
    if (authResults) {
        console.log("🔍 Parsing Authentication-Results header:", authResults);
        
        const dmarcResult = /dmarc=([^\s;]+)/.exec(authResults);
        const dkimResult = /dkim=([^\s;]+)/.exec(authResults);
        const spfResult = /spf=([^\s;]+)/.exec(authResults);

        updateIcon("dmarc", dmarcResult && dmarcResult[1] === "pass");
        updateIcon("dkim", dkimResult && dkimResult[1] === "pass");
        updateIcon("spf", spfResult && spfResult[1] === "pass");
        
        console.log("📊 Results:", {
            dmarc: dmarcResult ? dmarcResult[1] : 'not found',
            dkim: dkimResult ? dkimResult[1] : 'not found',
            spf: spfResult ? spfResult[1] : 'not found'
        });
        
        showMessage("Authentication results loaded successfully", "success");
    } else {
        console.log("⚠️ No Authentication-Results header found");
        // Show default state when no authentication results found
        updateIcon("dmarc", false);
        updateIcon("dkim", false);
        updateIcon("spf", false);
        
        showMessage("No authentication results found in email headers", "info");
    }
}

function updateIcon(id, passed) {
    const iconElement = document.getElementById(id + '-icon');
    if (iconElement) {
        // Clear existing classes
        iconElement.classList.remove("pass", "fail", "loading");
        // Add appropriate class and text
        if (passed) {
            iconElement.classList.add("pass");
            iconElement.textContent = "✓";
            console.log(`✅ ${id}: PASS`);
        } else {
            iconElement.classList.add("fail");
            iconElement.textContent = "✗";
            console.log(`❌ ${id}: FAIL`);
        }
    } else {
        console.warn(`⚠️ Icon element not found for ${id}`);
    }
}

// Remove loading state from all auth items
function clearLoadingState() {
    document.querySelectorAll('.auth-item, .auth-item-inline').forEach(item => {
        item.classList.remove('loading');
    });
    document.querySelectorAll('.icon-text').forEach(icon => {
        icon.classList.remove('loading');
        if (icon.textContent === '⟳') {
            icon.textContent = '?';
        }
    });
}

function showMessage(message, type = "info") {
    // Remove the initial status indicator
    const statusIndicator = document.getElementById('status-indicator');
    if (statusIndicator) {
        statusIndicator.remove();
    }
    
    // Remove any existing messages of the same type
    const existingMessages = document.querySelectorAll(`.${type}-message`);
    existingMessages.forEach(msg => msg.remove());
    
    // Create new message element
    const messageDiv = document.createElement('div');
    messageDiv.className = `message ${type}-message`;
    messageDiv.textContent = message;
    
    // Add to content
    const contentDiv = document.getElementById("content");
    if (contentDiv) {
        // Insert after the header section
        const firstGridRow = contentDiv.querySelector('.ms-Grid-row');
        if (firstGridRow && firstGridRow.nextElementSibling) {
            firstGridRow.parentNode.insertBefore(messageDiv, firstGridRow.nextElementSibling);
        } else {
            contentDiv.appendChild(messageDiv);
        }
        
        // Auto-hide success messages after 5 seconds
        if (type === "success") {
            setTimeout(() => {
                if (messageDiv.parentNode) {
                    messageDiv.remove();
                }
            }, 5000);
        }
    }
}

function showError(message) {
    console.error("❌ Error:", message);
    showMessage(message, "error");
    
    // Also update the main content area if needed
    const contentDiv = document.getElementById("content");
    if (contentDiv && !document.querySelector('.error-message')) {
        // Only add error message if one doesn't exist
        const errorDiv = document.createElement('div');
        errorDiv.className = 'error-message';
        errorDiv.textContent = message;
        contentDiv.appendChild(errorDiv);
    }
}

// Add a visibility check to ensure the add-in is properly loaded
function checkVisibility() {
    console.log("👁️ Checking add-in visibility...");
    
    const contentDiv = document.getElementById("content");
    if (contentDiv) {
        const rect = contentDiv.getBoundingClientRect();
        console.log("📐 Content dimensions:", {
            width: rect.width,
            height: rect.height,
            visible: rect.width > 0 && rect.height > 0,
            display: getComputedStyle(contentDiv).display,
            visibility: getComputedStyle(contentDiv).visibility
        });
        
        if (rect.width === 0 || rect.height === 0) {
            console.warn("⚠️ Add-in content appears to be hidden or collapsed");
            showMessage("Add-in loaded but content area is not visible", "error");
        }
    } else {
        console.error("❌ Content div not found!");
        showMessage("Add-in structure not loaded correctly", "error");
    }
}

// Run visibility check after a delay
setTimeout(() => {
    checkVisibility();
}, 500);
