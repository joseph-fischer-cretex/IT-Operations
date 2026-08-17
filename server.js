const express = require('express');
const axios = require('axios');
const path = require('path');
const app = express();

app.use(express.static(__dirname));

const clientId = process.env.AzureAppID;
const tenantId = process.env.AzureTenantID;
const clientSecret = process.env.MICROSOFT_PROVIDER_AUTHENTICATION_SECRET;

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

// Under iisnode, process.env.PORT is a named pipe path (e.g. \\.\pipe\xxxx),
// not a numeric TCP port -- app.listen() must be called with just that path
// (no host argument) in that case. Locally/elsewhere, PORT (if set) is a
// real TCP port, so we bind to 0.0.0.0; with no PORT at all we fall back to
// localhost:3000 for local development.
const port = process.env.PORT || 3000;
const isNamedPipe = typeof port === 'string' && port.indexOf('\\\\.\\pipe\\') === 0;

function onListen() {
    console.log('Server running on ' + port);
}

if (isNamedPipe) {
    app.listen(port, onListen);
} else if (process.env.PORT) {
    app.listen(port, '0.0.0.0', onListen);
} else {
    app.listen(port, '127.0.0.1', onListen);
}