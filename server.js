const express = require('express');
const axios = require('axios');
const path = require('path');
const app = express();

app.use(express.static(path.join(__dirname, 'public')));

const clientId = process.env.AzureAppID;
const tenantId = process.env.AzureTenantID;
const clientSecret = process.env.AzureClientSecret;

app.get('/api/searchUsers', async (req, res) => {
    const query = req.query.q;
    if (!query) {
        return res.status(400).json({ error: 'Query parameter q is required.' });
    }

    try {
        // 1. Get Access Token from Entra ID via Client Credentials Flow
        const tokenResponse = await axios.post(
            'https://login.microsoftonline.com/' + tenantId + '/oauth2/v2.0/token',
            new URLSearchParams({
                grant_type: 'client_credentials',
                client_id: clientId,
                client_secret: clientSecret,
                scope: 'https://graph.microsoft.com/.default'
            }).toString(),
            { headers: { 'Content-Type': 'application/x-www-form-urlencoded' } }
        );

        const accessToken = tokenResponse.data.access_token;

        // 2. Query Microsoft Graph API
        const graphResponse = await axios.get(
            'https://graph.microsoft.com/v1.0/users?' +
            '$search="displayName:' + query + '" OR "mail:' + query + '" OR "userPrincipalName:' + query + '"' +
            '&$select=displayName,givenName,surname,mail,userPrincipalName,jobTitle' +
            '&$top=10',
            {
                headers: {
                    Authorization: 'Bearer ' + accessToken,
                    ConsistencyLevel: 'eventual'
                }
            }
        );

        // 3. Map Graph API response to frontend format
        const users = graphResponse.data.value.map(function(user) {
            return {
                first: user.givenName || '',
                last: user.surname || user.displayName,
                email: user.mail || 'N/A',
                upn: user.userPrincipalName,
                username: user.userPrincipalName.split('@')[0],
                title: user.jobTitle || 'No Title'
            };
        });

        res.json(users);

    } catch (error) {
        console.error('Error querying Microsoft Graph:', error.response ? error.response.data : error.message);
        res.status(500).json({ error: 'Failed to fetch users from Entra ID.' });
    }
});

// IISNode injects a Named Pipe path into process.env.PORT instead of a number.
// We must detect this and pass it directly to app.listen() without parsing it.
var port = process.env.PORT;
if (port && port.indexOf('pipe') !== -1) {
    // Running under IISNode in Azure App Service
    app.listen(port, function() {
        console.log('Backend server running on IISNode pipe: ' + port);
    });
} else {
    // Running locally or in Kudu console — bind to 127.0.0.1 to avoid IPv6 sandbox crash
    app.listen(3000, '127.0.0.1', function() {
        console.log('Backend server running on http://127.0.0.1:3000');
    });
}