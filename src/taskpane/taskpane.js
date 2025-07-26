Office.onReady(function (info) {
    if (info.host === Office.HostType.Outlook) {
        displayAuthenticationResults();
    }
});

function displayAuthenticationResults() {
    Office.context.mailbox.item.getHeadersAsync(function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            const headers = asyncResult.value;
            const authResults = headers.get("Authentication-Results");

            if (authResults) {
                const dmarcResult = /dmarc=([^ ]+)/.exec(authResults);
                const dkimResult = /dkim=([^ ]+)/.exec(authResults);
                const spfResult = /spf=([^ ]+)/.exec(authResults);

                updateIcon("dmarc", dmarcResult && dmarcResult[1] === "pass");
                updateIcon("dkim", dkimResult && dkimResult[1] === "pass");
                updateIcon("spf", spfResult && spfResult[1] === "pass");
            }
        }
    });
}

function updateIcon(id, passed) {
    const element = document.getElementById(id).getElementsByClassName("icon")[0];
    if (passed) {
        element.classList.add("pass");
    } else {
        element.classList.add("fail");
    }
}
