const express = require('express');
const axios = require('axios');
const path = require('path');
const app = express();
const PORT = process.env.PORT || 3000;

// Serve your static frontend files (like index.html)
app.use(express.static(path.join(__dirname, 'public')));

// Fetch Environment Variables securely from Azure App Service
const clientId = process.env.AzureAppID;
const tenantId = process.env.AzureTenantID;
const clientSecret = process.env.AzureClientSecret;

// Secure API endpoint for searching users in Entra ID
app.get('/api/searchUsers', async (req, res) => {
    const query = req.query.q;
    if (!query) {
        return res.status(400).json({ error: "Query parameter 'q' is required." });
    }

    try {
        // 1. Get Access Token from Entra ID via Client Credentials Flow
        const tokenResponse = await axios.post(
            `https://login.microsoftonline.com/${tenantId}/oauth2/v2.0/token`,
            new URLSearchParams({
                grant_type: 'client_credentials',
                client_id: clientId,
                client_secret: clientSecret,
                scope: 'https://graph.microsoft.com/.default'
            }).toString(),
            { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
        );

        const accessToken = tokenResponse.data.access_token;

        // 2. Query Microsoft Graph API to find matching users
        // We use $search and ConsistencyLevel: eventual to search across first, last, UPN, and email.
        const graphResponse = await axios.get(
            `https://graph.microsoft.com/v1.0/users?$search="displayName:${query}" OR "mail:${query}" OR "userPrincipalName:${query}"&$select=displayName,givenName,surname,mail,userPrincipalName,jobTitle&$top=10`,
            {
                headers: {
                    Authorization: `Bearer ${accessToken}`,
                    ConsistencyLevel: 'eventual'
                }
            }
        );

        // 3. Map Graph API response to the format your frontend HTML expects
        const users = graphResponse.data.value.map(user => ({
            first: user.givenName || '',
            last: user.surname || user.displayName,
            email: user.mail || 'N/A',
            upn: user.userPrincipalName,
            username: user.userPrincipalName.split('@')[0], // Extracting SamAccountName-style username
            title: user.jobTitle || 'No Title'
        }));

        res.json(users);

    } catch (error) {
        console.error("Error querying Microsoft Graph:", error.response?.data || error.message);
        res.status(500).json({ error: "Failed to fetch users from Entra ID." });
    }
});

app.listen(PORT, () => console.log(`Backend server running on port ${PORT}`));
