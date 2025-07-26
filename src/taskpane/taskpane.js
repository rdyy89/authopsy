Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        console.log("Authopsy add-in initializing...");
        displayAuthenticationResults();
    }
});

function displayAuthenticationResults() {
    try {
        // First try to use the internetHeaders API (Outlook 2019/365)
        if (Office.context.mailbox.item.internetHeaders) {
            Office.context.mailbox.item.internetHeaders.getAsync(
                ["Authentication-Results"], 
                function (asyncResult) {
                    if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                        const headers = asyncResult.value;
                        const authResults = headers["Authentication-Results"];
                        parseAuthenticationResults(authResults);
                    } else {
                        console.log("internetHeaders failed, trying alternative method");
                        tryAlternativeMethod();
                    }
                }
            );
        } else {
            // Fallback for older versions
            tryAlternativeMethod();
        }
    } catch (error) {
        console.error("Error in displayAuthenticationResults:", error);
        tryAlternativeMethod();
    }
}

function tryAlternativeMethod() {
    // For demonstration purposes, let's show the UI with default values
    // In a real scenario, you might need to use EWS or Graph API
    console.log("Using alternative method - showing default state");
    
    // Show default icons (failed state) since we can't get headers
    updateIcon("dmarc", false);
    updateIcon("dkim", false);
    updateIcon("spf", false);
    
    // You could also try to get the item's properties
    if (Office.context.mailbox.item.subject) {
        console.log("Email subject:", Office.context.mailbox.item.subject);
    }
}

function parseAuthenticationResults(authResults) {
    if (authResults) {
        console.log("Authentication-Results header:", authResults);
        
        const dmarcResult = /dmarc=([^\s;]+)/.exec(authResults);
        const dkimResult = /dkim=([^\s;]+)/.exec(authResults);
        const spfResult = /spf=([^\s;]+)/.exec(authResults);

        updateIcon("dmarc", dmarcResult && dmarcResult[1] === "pass");
        updateIcon("dkim", dkimResult && dkimResult[1] === "pass");
        updateIcon("spf", spfResult && spfResult[1] === "pass");
    } else {
        console.log("No Authentication-Results header found");
        // Show default state when no authentication results found
        updateIcon("dmarc", false);
        updateIcon("dkim", false);
        updateIcon("spf", false);
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
                console.log(`${id}: PASS`);
            } else {
                iconElement.classList.add("fail");
                console.log(`${id}: FAIL`);
            }
        }
    }
}

function showError(message) {
    const contentDiv = document.getElementById("content");
    if (contentDiv) {
        contentDiv.innerHTML = `
            <div class="ms-Grid" dir="ltr">
                <div class="ms-Grid-row">
                    <div class="ms-Grid-col ms-sm6 ms-md8 ms-lg10">
                        <h2 class="ms-font-xl">Authentication Results</h2>
                        <p class="ms-fontColor-redDark">${message}</p>
                    </div>
                </div>
            </div>
        `;
    }
}
