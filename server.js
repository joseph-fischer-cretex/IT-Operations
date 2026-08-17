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

// httpPlatformHandler passes a clean numeric port via HTTP_PLATFORM_PORT
const port = process.env.PORT || 3000;
const host = process.env.PORT ? '0.0.0.0' : '127.0.0.1';

app.listen(port, host, function() {
    console.log('Server running on ' + host + ':' + port);
});