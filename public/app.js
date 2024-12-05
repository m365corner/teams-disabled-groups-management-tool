const msalInstance = new msal.PublicClientApplication({
    auth: {
        clientId: "<client-id-goes-here>",
        authority: "https://login.microsoftonline.com/<tenant-id-goes-here>",
        redirectUri: "http://localhost:8000",
    },
});

let allGroups = []; // Store all fetched groups
let filteredGroups = []; // Store groups after filtering

// Login and Logout
async function login() {
    try {
        const loginResponse = await msalInstance.loginPopup({
            scopes: ["Group.Read.All", "Mail.Send"],
        });
        msalInstance.setActiveAccount(loginResponse.account);
        alert("Login successful.");
        await fetchAndDisplayGroups();
    } catch (error) {
        console.error("Login failed:", error);
        alert("Login failed.");
    }
}

function logout() {
    msalInstance.logoutPopup().then(() => alert("Logout successful."));
}

// Fetch and Display Groups
async function fetchAndDisplayGroups() {
    try {
        const response = await callGraphApi("/groups?$select=id,displayName,mail,groupTypes,mailEnabled,resourceProvisioningOptions");

        allGroups = await Promise.all(
            response.value
                .filter(group => 
                    !group.resourceProvisioningOptions.includes("Team") && 
                    (!group.groupTypes.includes("Unified") || group.resourceProvisioningOptions.length === 0)
                )
                .map(async group => ({
                    ...group,
                    memberCount: await fetchGroupMemberCount(group.id),
                }))
        );

        filteredGroups = allGroups; // Initialize filtered groups
        populateTable(filteredGroups);
    } catch (error) {
        console.error("Error fetching groups:", error);
        alert("Failed to fetch groups.");
    }
}

// Fetch Member Count
async function fetchGroupMemberCount(groupId) {
    if (!groupId) return 0;

    try {
        const response = await callGraphApi(`/groups/${groupId}/members`);
        return response.value.filter(member => member["@odata.type"] === "#microsoft.graph.user").length;
    } catch (error) {
        console.error(`Error fetching members for group ${groupId}:`, error);
        return 0;
    }
}

// Apply Filters and Search
function applyFilters() {
    const searchText = document.getElementById("searchBox").value.toLowerCase();
    const groupType = document.getElementById("groupTypeFilter").value;
    const mailEnabled = document.getElementById("mailEnabledFilter").value;

    filteredGroups = allGroups.filter(group => {
        const matchesSearch = searchText
            ? (group.displayName?.toLowerCase().includes(searchText) || group.mail?.toLowerCase().includes(searchText))
            : true;

            const matchesGroupType = groupType
            ? (groupType === "Security" && !group.groupTypes.length && !group.resourceProvisioningOptions.includes("Team")) ||
              (groupType === "Distribution" && !group.groupTypes.length && group.mailEnabled)
            : true;

        const matchesMailEnabled = mailEnabled
            ? (group.mailEnabled.toString() === mailEnabled)
            : true;

        return matchesSearch && matchesGroupType && matchesMailEnabled;
    });

    populateTable(filteredGroups);
}





// Populate Table
function populateTable(data) {
    const outputHeader = document.getElementById("outputHeader");
    const outputBody = document.getElementById("outputBody");
    outputHeader.innerHTML = "<th>Group Name</th><th>Group Mail</th><th>Group Type</th><th>Mail Enabled</th><th>Members</th>";
    outputBody.innerHTML = data
        .map(group => `
            <tr>
                <td>${group.displayName || "N/A"}</td>
                <td>${group.mail || "N/A"}</td>
                <td>${group.groupTypes.includes("Unified") ? "Unified" : group.groupTypes.join(", ") || "Security/Distribution"}</td>
                <td>${group.mailEnabled ? "Yes" : "No"}</td>
                <td>${group.memberCount}</td>
            </tr>
        `)
        .join("");
}

// Mail Report and Download CSV (Existing Functions)
async function sendReportAsMail() {
    const email = document.getElementById("adminEmail").value;
    if (!email) return alert("Please provide an admin email.");

    const headers = [...document.querySelectorAll("#outputHeader th")].map(th => th.textContent);
    const rows = [...document.querySelectorAll("#outputBody tr")].map(tr =>
        [...tr.querySelectorAll("td")].map(td => td.textContent)
    );

    const emailContent = rows.map(row => `<tr>${row.map(cell => `<td>${cell}</td>`).join("")}</tr>`).join("");
    const emailBody = `<table border="1"><thead><tr>${headers.map(header => `<th>${header}</th>`).join("")}</tr></thead><tbody>${emailContent}</tbody></table>`;

    const message = {
        message: {
            subject: "Tenant Groups Report",
            body: { contentType: "HTML", content: emailBody },
            toRecipients: [{ emailAddress: { address: email } }],
        },
    };
    try {
        await callGraphApi("/me/sendMail", "POST", message);
        alert("Report sent!");
    } catch (error) {
        console.error("Error sending report:", error);
        alert("Failed to send the report.");
    }
}


function downloadReportAsCSV() {
    const headers = [...document.querySelectorAll("#outputHeader th")].map(th => th.textContent);
    const rows = [...document.querySelectorAll("#outputBody tr")].map(tr =>
        [...tr.querySelectorAll("td")].map(td => td.textContent)
    );
    if (!rows.length) return alert("No data to download.");

    const csvContent = [headers.join(","), ...rows.map(row => row.join(","))].join("\n");
    const blob = new Blob([csvContent], { type: "text/csv" });
    const a = document.createElement("a");
    a.href = URL.createObjectURL(blob);
    a.download = "Tenant_Groups_Report.csv";
    a.click();
}



// Call Graph API (Unchanged)

async function callGraphApi(endpoint, method = "GET", body = null) {
    const account = msalInstance.getActiveAccount();
    if (!account) throw new Error("Please log in first.");

    const tokenResponse = await msalInstance.acquireTokenSilent({
        scopes: ["Mail.Send"],
        account,
    });

    const response = await fetch(`https://graph.microsoft.com/v1.0${endpoint}`, {
        method,
        headers: {
            Authorization: `Bearer ${tokenResponse.accessToken}`,
            "Content-Type": "application/json",
        },
        body: body ? JSON.stringify(body) : null, // Serialize the body if provided
    });

    // Check for non-OK responses
    if (!response.ok) {
        const errorText = await response.text();
        console.error("Graph API error response:", errorText);
        throw new Error(`Graph API call failed: ${response.status} ${response.statusText}`);
    }

    // Handle responses with no content or body
    const contentType = response.headers.get("content-type");
    if (!contentType || !contentType.includes("application/json")) {
        console.warn("No JSON content in response."); // Log when there's no JSON
        return {}; // Return an empty object for non-JSON responses
    }

    // Attempt to parse JSON for valid responses
    try {
        return await response.json();
    } catch (error) {
        console.error("Error parsing JSON response:", error);
        throw new Error("Failed to parse Graph API response.");
    }
}



