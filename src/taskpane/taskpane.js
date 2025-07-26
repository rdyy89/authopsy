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
    if (info.host === Office.HostType.Outlook) {
        console.log("✅ Authopsy add-in initializing...");
        displayAuthenticationResults();
    } else {
        console.error("❌ Add-in loaded in unsupported host:", info.host);
    }
});

function displayAuthenticationResults() {
    try {
        console.log("🔍 Attempting to retrieve authentication headers...");
        
        // Add loading state
        document.querySelectorAll('.auth-item').forEach(item => {
            item.classList.add('loading');
        });
        
        // First try to use the internetHeaders API (Outlook 2019/365)
        if (Office.context.mailbox.item.internetHeaders) {
            console.log("📡 Using internetHeaders API...");
            Office.context.mailbox.item.internetHeaders.getAsync(
                ["Authentication-Results"], 
                function (asyncResult) {
                    // Remove loading state
                    document.querySelectorAll('.auth-item').forEach(item => {
                        item.classList.remove('loading');
                    });
                    
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        const headers = asyncResult.value;
                        const authResults = headers["Authentication-Results"];
                        console.log("✅ Successfully retrieved headers");
                        parseAuthenticationResults(authResults);
                    } else {
                        console.log("⚠️ internetHeaders failed, trying alternative method");
                        tryAlternativeMethod();
                    }
                }
            );
        } else {
            console.log("⚠️ internetHeaders not available, using alternative method");
            // Remove loading state
            document.querySelectorAll('.auth-item').forEach(item => {
                item.classList.remove('loading');
            });
            tryAlternativeMethod();
        }
    } catch (error) {
        console.error("❌ Error in displayAuthenticationResults:", error);
        // Remove loading state
        document.querySelectorAll('.auth-item').forEach(item => {
            item.classList.remove('loading');
        });
        tryAlternativeMethod();
    }
}

function tryAlternativeMethod() {
    console.log("🔄 Using fallback method - showing demo state");
    
    // For demonstration purposes, let's show the UI with sample data
    // In a real scenario, you might need to use EWS or Graph API
    
    // Show sample results for demo
    updateIcon("dmarc", Math.random() > 0.5);
    updateIcon("dkim", Math.random() > 0.5);
    updateIcon("spf", Math.random() > 0.5);
    
    // Show a message about the method used
    showMessage("Using demo data - internetHeaders API not available in this context", "info");
    
    // You could also try to get the item's properties
    if (Office.context.mailbox.item.subject) {
        console.log("📧 Email subject:", Office.context.mailbox.item.subject);
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
    const element = document.getElementById(id);
    if (element) {
        const iconElement = element.getElementsByClassName("icon")[0];
        if (iconElement) {
            // Clear existing classes
            iconElement.classList.remove("pass", "fail");
            // Add appropriate class
            if (passed) {
                iconElement.classList.add("pass");
                console.log(`✅ ${id}: PASS`);
            } else {
                iconElement.classList.add("fail");
                console.log(`❌ ${id}: FAIL`);
            }
        }
    }
}

function showMessage(message, type = "info") {
    // Remove any existing messages
    const existingMessage = document.querySelector('.message');
    if (existingMessage) {
        existingMessage.remove();
    }
    
    // Create new message element
    const messageDiv = document.createElement('div');
    messageDiv.className = `message ${type}-message`;
    messageDiv.textContent = message;
    
    // Add to content
    const contentDiv = document.getElementById("content");
    if (contentDiv) {
        contentDiv.appendChild(messageDiv);
        
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
}
