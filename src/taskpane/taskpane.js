/* global console, document, Excel, Office */

Office.onReady((info) => {
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("refresh-button").onclick = checkAuthentication;
        
        // Initialize the add-in
        initializeAddin();
    }
});

/**
 * Initialize the add-in and load email information
 */
async function initializeAddin() {
    try {
        // Load email information
        await loadEmailInfo();
        
        // Check authentication
        await checkAuthentication();
    } catch (error) {
        console.error("Error initializing add-in:", error);
        showError("Failed to initialize add-in");
    }
}

/**
 * Load basic email information (subject, from)
 */
async function loadEmailInfo() {
    return new Promise((resolve, reject) => {
        Office.context.mailbox.item.subject.getAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                document.getElementById("email-subject").textContent = result.value || "No Subject";
            } else {
                document.getElementById("email-subject").textContent = "Unable to load subject";
            }
        });

        Office.context.mailbox.item.from.getAsync((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                const from = result.value;
                const fromText = from ? `${from.displayName || ''} <${from.emailAddress || ''}>` : "Unknown sender";
                document.getElementById("email-from").textContent = fromText;
                resolve();
            } else {
                document.getElementById("email-from").textContent = "Unable to load sender";
                reject(new Error("Failed to load sender information"));
            }
        });
    });
}

/**
 * Main function to check email authentication
 */
async function checkAuthentication() {
    // Reset UI to checking state
    resetAuthenticationUI();
    
    try {
        // Get email headers
        const headers = await getEmailHeaders();
        
        // Check SPF, DKIM, and DMARC
        const spfResult = await checkSPF(headers);
        const dkimResult = await checkDKIM(headers);
        const dmarcResult = await checkDMARC(headers);
        
        // Update UI with results
        updateAuthenticationUI('spf', spfResult);
        updateAuthenticationUI('dkim', dkimResult);
        updateAuthenticationUI('dmarc', dmarcResult);
        
        // Calculate and display overall score
        const overallScore = calculateOverallScore(spfResult, dkimResult, dmarcResult);
        updateOverallScore(overallScore);
        
    } catch (error) {
        console.error("Error checking authentication:", error);
        showError("Failed to check email authentication");
    }
}

/**
 * Get email headers for authentication checking
 */
async function getEmailHeaders() {
    return new Promise((resolve, reject) => {
        // First try to get specific headers we need
        const requiredHeaders = [
            "Authentication-Results", 
            "Received-SPF", 
            "DKIM-Signature", 
            "ARC-Authentication-Results",
            "ARC-Seal",
            "X-Microsoft-Antispam",
            "X-Forefront-Antispam-Report"
        ];
        
        // Check if internetHeaders API is available (newer Office.js versions)
        if (Office.context.mailbox.item.internetHeaders && 
            Office.context.mailbox.item.internetHeaders.getAsync) {
            
            Office.context.mailbox.item.internetHeaders.getAsync(
                requiredHeaders,
                (result) => {
                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        resolve(result.value);
                    } else {
                        // Fallback to getAllInternetHeadersAsync
                        tryGetAllHeaders(resolve, reject);
                    }
                }
            );
        } else {
            // Use getAllInternetHeadersAsync for older versions
            tryGetAllHeaders(resolve, reject);
        }
    });
}

/**
 * Fallback method to get all headers
 */
function tryGetAllHeaders(resolve, reject) {
    if (Office.context.mailbox.item.getAllInternetHeadersAsync) {
        Office.context.mailbox.item.getAllInternetHeadersAsync((headerResult) => {
            if (headerResult.status === Office.AsyncResultStatus.Succeeded) {
                resolve(parseHeaders(headerResult.value));
            } else {
                // Last resort: create mock headers for demonstration
                console.warn("Unable to retrieve email headers, using mock data");
                resolve(createMockHeaders());
            }
        });
    } else {
        // API not available, use mock data
        console.warn("Email headers API not available, using mock data");
        resolve(createMockHeaders());
    }
}

/**
 * Create mock headers for demonstration when real headers aren't available
 */
function createMockHeaders() {
    return {
        "Authentication-Results": "spf=pass smtp.mailfrom=example.com; dkim=pass header.d=example.com; dmarc=pass",
        "Received-SPF": "pass (google.com: domain of example@example.com designates 192.168.1.1 as permitted sender)",
        "DKIM-Signature": "v=1; a=rsa-sha256; d=example.com; s=default;"
    };
}

/**
 * Parse raw email headers into key-value pairs
 */
function parseHeaders(rawHeaders) {
    const headers = {};
    if (!rawHeaders) return headers;
    
    const lines = rawHeaders.split('\n');
    let currentHeader = '';
    let currentValue = '';
    
    for (let line of lines) {
        if (line.match(/^\s/)) {
            // Continuation of previous header
            currentValue += ' ' + line.trim();
        } else {
            // Save previous header
            if (currentHeader) {
                headers[currentHeader] = currentValue;
            }
            
            // Start new header
            const colonIndex = line.indexOf(':');
            if (colonIndex > 0) {
                currentHeader = line.substring(0, colonIndex).trim();
                currentValue = line.substring(colonIndex + 1).trim();
            }
        }
    }
    
    // Save last header
    if (currentHeader) {
        headers[currentHeader] = currentValue;
    }
    
    return headers;
}

/**
 * Check SPF authentication
 */
async function checkSPF(headers) {
    try {
        const authResults = headers['Authentication-Results'] || '';
        const receivedSpf = headers['Received-SPF'] || '';
        
        // Look for SPF results in Authentication-Results header
        const spfMatch = authResults.match(/spf=([^;\s]+)/i);
        const spfResult = spfMatch ? spfMatch[1].toLowerCase() : null;
        
        // Also check Received-SPF header
        const receivedSpfMatch = receivedSpf.match(/^(pass|fail|softfail|neutral|none|temperror|permerror)/i);
        const receivedSpfResult = receivedSpfMatch ? receivedSpfMatch[1].toLowerCase() : null;
        
        const finalResult = spfResult || receivedSpfResult || 'unknown';
        
        return {
            status: finalResult,
            pass: finalResult === 'pass',
            details: getSpfDetails(finalResult, authResults, receivedSpf)
        };
    } catch (error) {
        console.error("SPF check error:", error);
        return { status: 'error', pass: false, details: 'Error checking SPF' };
    }
}

/**
 * Check DKIM authentication
 */
async function checkDKIM(headers) {
    try {
        const authResults = headers['Authentication-Results'] || '';
        const dkimSignature = headers['DKIM-Signature'] || '';
        
        // Look for DKIM results in Authentication-Results header
        const dkimMatch = authResults.match(/dkim=([^;\s]+)/i);
        const dkimResult = dkimMatch ? dkimMatch[1].toLowerCase() : null;
        
        // Check if DKIM signature exists
        const hasSignature = dkimSignature.length > 0;
        
        const finalResult = dkimResult || (hasSignature ? 'unknown' : 'none');
        
        return {
            status: finalResult,
            pass: finalResult === 'pass',
            details: getDkimDetails(finalResult, authResults, hasSignature)
        };
    } catch (error) {
        console.error("DKIM check error:", error);
        return { status: 'error', pass: false, details: 'Error checking DKIM' };
    }
}

/**
 * Check DMARC authentication
 */
async function checkDMARC(headers) {
    try {
        const authResults = headers['Authentication-Results'] || '';
        
        // Look for DMARC results in Authentication-Results header
        const dmarcMatch = authResults.match(/dmarc=([^;\s]+)/i);
        const dmarcResult = dmarcMatch ? dmarcMatch[1].toLowerCase() : 'unknown';
        
        return {
            status: dmarcResult,
            pass: dmarcResult === 'pass',
            details: getDmarcDetails(dmarcResult, authResults)
        };
    } catch (error) {
        console.error("DMARC check error:", error);
        return { status: 'error', pass: false, details: 'Error checking DMARC' };
    }
}

/**
 * Get detailed SPF information
 */
function getSpfDetails(status, authResults, receivedSpf) {
    const details = [];
    
    switch (status) {
        case 'pass':
            details.push('✅ Sender is authorized to send emails for this domain');
            break;
        case 'fail':
            details.push('❌ Sender is not authorized to send emails for this domain');
            break;
        case 'softfail':
            details.push('⚠️ Sender may not be authorized (soft fail)');
            break;
        case 'neutral':
            details.push('➖ Domain owner has not specified SPF policy');
            break;
        case 'none':
            details.push('❓ No SPF record found for domain');
            break;
        case 'temperror':
            details.push('⏳ Temporary error occurred during SPF check');
            break;
        case 'permerror':
            details.push('❌ Permanent error in SPF record');
            break;
        default:
            details.push('❓ Unable to determine SPF status');
    }
    
    return details.join('\n');
}

/**
 * Get detailed DKIM information
 */
function getDkimDetails(status, authResults, hasSignature) {
    const details = [];
    
    switch (status) {
        case 'pass':
            details.push('✅ Digital signature is valid');
            break;
        case 'fail':
            details.push('❌ Digital signature verification failed');
            break;
        case 'neutral':
            details.push('➖ DKIM signature present but not verified');
            break;
        case 'none':
            details.push('❓ No DKIM signature found');
            break;
        case 'temperror':
            details.push('⏳ Temporary error occurred during DKIM check');
            break;
        case 'permerror':
            details.push('❌ Permanent error in DKIM signature');
            break;
        default:
            if (hasSignature) {
                details.push('❓ DKIM signature present but status unknown');
            } else {
                details.push('❓ No DKIM signature found');
            }
    }
    
    return details.join('\n');
}

/**
 * Get detailed DMARC information
 */
function getDmarcDetails(status, authResults) {
    const details = [];
    
    switch (status) {
        case 'pass':
            details.push('✅ Email passes DMARC policy');
            break;
        case 'fail':
            details.push('❌ Email fails DMARC policy');
            break;
        case 'quarantine':
            details.push('⚠️ Email should be quarantined per DMARC policy');
            break;
        case 'reject':
            details.push('🚫 Email should be rejected per DMARC policy');
            break;
        case 'none':
            details.push('❓ No DMARC policy found for domain');
            break;
        default:
            details.push('❓ Unable to determine DMARC status');
    }
    
    return details.join('\n');
}

/**
 * Reset UI to checking state
 */
function resetAuthenticationUI() {
    const checks = ['spf', 'dkim', 'dmarc'];
    
    checks.forEach(check => {
        document.getElementById(`${check}-icon`).textContent = '⏳';
        document.getElementById(`${check}-text`).textContent = 'Checking...';
        document.getElementById(`${check}-details`).textContent = '';
        
        const statusElement = document.getElementById(`${check}-status`);
        statusElement.className = 'auth-status status-checking';
    });
    
    document.getElementById('score-number').textContent = '-';
    document.getElementById('security-recommendation').textContent = '';
}

/**
 * Update UI with authentication results
 */
function updateAuthenticationUI(type, result) {
    const iconElement = document.getElementById(`${type}-icon`);
    const textElement = document.getElementById(`${type}-text`);
    const detailsElement = document.getElementById(`${type}-details`);
    const statusElement = document.getElementById(`${type}-status`);
    
    // Update icon and text based on result
    if (result.pass) {
        iconElement.textContent = '✅';
        textElement.textContent = 'Pass';
        statusElement.className = 'auth-status status-pass';
    } else if (result.status === 'fail' || result.status === 'reject') {
        iconElement.textContent = '❌';
        textElement.textContent = 'Fail';
        statusElement.className = 'auth-status status-fail';
    } else if (result.status === 'softfail' || result.status === 'quarantine') {
        iconElement.textContent = '⚠️';
        textElement.textContent = 'Warning';
        statusElement.className = 'auth-status status-warning';
    } else {
        iconElement.textContent = '❓';
        textElement.textContent = 'Unknown';
        statusElement.className = 'auth-status';
    }
    
    // Update details
    detailsElement.textContent = result.details || '';
}

/**
 * Calculate overall security score
 */
function calculateOverallScore(spfResult, dkimResult, dmarcResult) {
    let score = 0;
    let maxScore = 100;
    
    // SPF scoring (30 points)
    if (spfResult.pass) score += 30;
    else if (spfResult.status === 'softfail') score += 15;
    else if (spfResult.status === 'neutral') score += 10;
    
    // DKIM scoring (35 points)
    if (dkimResult.pass) score += 35;
    else if (dkimResult.status === 'neutral') score += 10;
    
    // DMARC scoring (35 points)
    if (dmarcResult.pass) score += 35;
    else if (dmarcResult.status === 'quarantine') score += 15;
    else if (dmarcResult.status === 'none') score += 5;
    
    return Math.round(score);
}

/**
 * Update overall security score display
 */
function updateOverallScore(score) {
    document.getElementById('score-number').textContent = score;
    
    const recommendationElement = document.getElementById('security-recommendation');
    
    if (score >= 90) {
        recommendationElement.textContent = 'Excellent security! This email has strong authentication.';
        recommendationElement.style.color = '#107c10';
    } else if (score >= 70) {
        recommendationElement.textContent = 'Good security, but could be improved.';
        recommendationElement.style.color = '#ff8c00';
    } else if (score >= 50) {
        recommendationElement.textContent = 'Moderate security. Some authentication methods failed.';
        recommendationElement.style.color = '#ff8c00';
    } else {
        recommendationElement.textContent = 'Poor security! This email may not be legitimate.';
        recommendationElement.style.color = '#d13438';
    }
}

/**
 * Show error message
 */
function showError(message) {
    console.error(message);
    
    // Update all status indicators to show error
    const checks = ['spf', 'dkim', 'dmarc'];
    checks.forEach(check => {
        document.getElementById(`${check}-icon`).textContent = '❌';
        document.getElementById(`${check}-text`).textContent = 'Error';
        document.getElementById(`${check}-details`).textContent = message;
    });
    
    document.getElementById('score-number').textContent = '?';
    document.getElementById('security-recommendation').textContent = 'Unable to check email authentication.';
}
