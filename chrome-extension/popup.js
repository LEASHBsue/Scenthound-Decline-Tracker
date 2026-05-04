// ==========================================
// MICROSOFT GRAPH & MSAL CONFIGURATION
// ==========================================
const msalConfig = {
    auth: {
        clientId: "47db5229-a470-4efa-895b-380a9da19af3", 
        authority: "https://login.microsoftonline.com/b9511ebb-4cb9-4992-a7a5-9e66d5fb2b29", 
        redirectUri: "chrome-extension://kiilekgkkanikdfbbjgjpkmjjpmameib/popup.html"
    }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
const loginRequest = { scopes: ["Sites.ReadWrite.All"] };

const SITE_ID = "aa335ef4-6dc1-4bb0-8f4c-d124ff03f81f"; 
const LIST_ID = "7b41b4fe-9308-4b25-b7a2-2ef34ddc8bbc"; 
const GRAPH_ENDPOINT = `https://graph.microsoft.com/v1.0/sites/${SITE_ID}/lists/${LIST_ID}/items`;
// ==========================================

// Restore the last manager's name
document.getElementById('managerName').value = localStorage.getItem('extManagerName') || '';

// Show/Hide "Other" text box
document.getElementById('reasonsBox').addEventListener('change', (e) => {
    const isOtherChecked = Array.from(document.querySelectorAll('input[type="checkbox"]:checked'))
                                .some(cb => cb.value.includes('Other'));
    document.getElementById('customReasonLabel').style.display = isOtherChecked ? 'block' : 'none';
    document.getElementById('customReason').style.display = isOtherChecked ? 'block' : 'none';
});

// Scrape MyTime automatically when popup opens
chrome.tabs.query({ active: true, currentWindow: true }, function(tabs) {
    let currentTab = tabs[0];
    document.getElementById('myTimeLink').value = currentTab.url;

    chrome.scripting.executeScript({
        target: { tabId: currentTab.id },
        function: scrapeMyTimeData,
    }, (results) => {
        if (results && results[0] && results[0].result) {
            let data = results[0].result;
            
            document.getElementById('dogName').value = data.dog;
            document.getElementById('parentName').value = data.parent;

            if (data.locationText) {
                const locSelect = document.getElementById('location');
                Array.from(locSelect.options).forEach(opt => {
                    if (data.locationText.toLowerCase().includes(opt.value.toLowerCase())) {
                        locSelect.value = opt.value;
                    }
                });
            }
        }
    });
});

function scrapeMyTimeData() {
    let parentEl = document.querySelector('h1.client-show__name');
    let parentName = parentEl ? parentEl.textContent.replace('', '').trim() : '';

    let dogEl = document.querySelector('span.qa-auto-child-label');
    let dogName = dogEl ? dogEl.textContent.replace('', '').replace('', '').trim() : '';

    let locationText = '';
    let iconBlocks = document.querySelectorAll('.icon-prefixed-block__text');
    for (let block of iconBlocks) {
        let text = block.textContent.trim();
        if (text.includes('Cedar Mill') || text.includes('Tigard') || text.includes('Hillsboro') || text.includes('Sherwood')) {
            locationText = text;
            break; 
        }
    }

    return {
        dog: dogName,
        parent: parentName,
        locationText: locationText
    };
}

// MSAL Authentication Helper
async function getAccessToken() {
    const accounts = msalInstance.getAllAccounts();
    if (accounts.length > 0) {
        try {
            const response = await msalInstance.acquireTokenSilent({
                ...loginRequest,
                account: accounts[0]
            });
            return response.accessToken;
        } catch (error) {
            console.warn("Silent token failed, falling back to popup", error);
        }
    }
    
    try {
        const response = await msalInstance.loginPopup(loginRequest);
        return response.accessToken;
    } catch (error) {
        console.error("Login failed:", error);
        return null;
    }
}

// Send Data to Microsoft Lists
document.getElementById('submitBtn').addEventListener('click', async () => {
    const managerName = document.getElementById('managerName').value;
    const location = document.getElementById('location').value;
    const dogName = document.getElementById('dogName').value;
    const parentName = document.getElementById('parentName').value;
    const myTimeLink = document.getElementById('myTimeLink').value;
    const customReason = document.getElementById('customReason').value;
    const notificationMethod = document.getElementById('notificationMethod').value;
    const notesField = document.getElementById('notesField').value;
    
    const selectedReasons = Array.from(document.querySelectorAll('input[type="checkbox"]:checked')).map(cb => cb.value);

    if (!managerName || selectedReasons.length === 0) {
        alert("Please enter your name and select at least one reason.");
        return;
    }

    let combinedName = "N/A";
    if (dogName && parentName) combinedName = `${dogName} & ${parentName}`;
    else if (dogName) combinedName = dogName;
    else if (parentName) combinedName = parentName;

    localStorage.setItem('extManagerName', managerName);
    document.getElementById('submitBtn').innerText = "Authenticating...";
    document.getElementById('submitBtn').disabled = true;

    // Grab the token before attempting to post
    const token = await getAccessToken();
    
    if (!token) {
        document.getElementById('status').innerText = "❌ Login required or failed.";
        document.getElementById('status').style.color = "#E74C3C";
        document.getElementById('submitBtn').innerText = "Try Again";
        document.getElementById('submitBtn').disabled = false;
        return;
    }

    document.getElementById('submitBtn').innerText = "Saving...";

    const payload = {
        fields: {
            Title: managerName,
            Location: location,
            DogName: dogName,
            ParentName: parentName,
            PetAndParentName: combinedName,
            MyTimeLink: myTimeLink,
            Reasons: selectedReasons.join(';'),
            CustomReason: customReason,
            NotificationMethod: notificationMethod,
            Notes: notesField
        }
    };

    try {
        const response = await fetch(GRAPH_ENDPOINT, {
            method: 'POST',
            headers: { 
                'Authorization': `Bearer ${token}`,
                'Content-Type': 'application/json' 
            },
            body: JSON.stringify(payload)
        });

        if (response.ok) {
            document.getElementById('status').innerText = "✅ Successfully Logged!";
            document.getElementById('status').style.color = "#27AE60";
            setTimeout(() => window.close(), 1500); 
        } else {
            throw new Error("Failed to save to Microsoft List");
        }
    } catch (error) {
        document.getElementById('status').innerText = "❌ Error saving entry.";
        document.getElementById('status').style.color = "#E74C3C";
        document.getElementById('submitBtn').innerText = "Try Again";
        document.getElementById('submitBtn').disabled = false;
    }
});
