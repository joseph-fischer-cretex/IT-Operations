const express = require('express');
const axios = require('axios');
const path = require('path');
const fs = require('fs');
const { spawn } = require('child_process');
const app = express();

app.use(express.json());
app.use((req, res, next) => {
    // Don't let the .ps1 source files themselves be downloaded via static serving
    if (req.path.toLowerCase().startsWith('/scripts/')) {
        return res.status(404).end();
    }
    next();
});
app.use(express.static(__dirname));

const SCRIPTS_DIR = path.join(__dirname, 'scripts');
const SCRIPT_NAME_PATTERN = /^[a-zA-Z0-9_-]+\.ps1$/;
const LOCKDOWN_ROLES = ['ServiceDesk', 'PortalAdmin', 'PDAdmin'];

const clientId = process.env.AzureAppID;
const tenantId = process.env.AzureTenantID;
const clientSecret = process.env.MICROSOFT_PROVIDER_AUTHENTICATION_SECRET;

// Shared helper: acquire an app-only Graph token via client credentials
async function getGraphToken() {
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
    return tokenResponse.data.access_token;
}

app.get('/api/searchUsers', async (req, res) => {
    const query = req.query.q;
    if (!query) {
        return res.status(400).json({ error: 'Query parameter q is required.' });
    }

    try {
        const accessToken = await getGraphToken();

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

// Streams a user's Entra ID profile photo. Returns 404 (no body) if the
// user has no photo set, so the frontend can fall back to an initials avatar.
app.get('/api/userPhoto', async (req, res) => {
    const upn = req.query.upn;
    if (!upn) {
        return res.status(400).json({ error: 'Query parameter upn is required.' });
    }

    try {
        const accessToken = await getGraphToken();

        const photoResponse = await axios.get(
            'https://graph.microsoft.com/v1.0/users/' + encodeURIComponent(upn) + '/photo/$value',
            {
                headers: { Authorization: 'Bearer ' + accessToken },
                responseType: 'arraybuffer',
                validateStatus: null
            }
        );

        if (photoResponse.status === 200) {
            const contentType = photoResponse.headers['content-type'] || 'image/jpeg';
            res.set('Content-Type', contentType);
            res.set('Cache-Control', 'private, max-age=3600');
            return res.send(Buffer.from(photoResponse.data));
        }

        // 404 = no photo on file; anything else = unexpected Graph error
        return res.status(photoResponse.status === 404 ? 404 : 502).end();

    } catch (error) {
        console.error('Error fetching user photo:', error.response ? error.response.status : error.message);
        res.status(502).end();
    }
});

// ==========================================================
// CUSTOM SCRIPTS (scripts/*.ps1)
// ==========================================================
// Reads Azure Easy Auth's injected client principal header to determine the
// caller's app roles. Returns [] if the header is absent (e.g. local dev
// without Easy Auth in front of the app) -- script execution is denied in
// that case rather than silently allowed.
function getCallerRoles(req) {
    const header = req.headers['x-ms-client-principal'];
    if (!header) return [];
    try {
        const decoded = JSON.parse(Buffer.from(header, 'base64').toString('utf8'));
        const claims = decoded.claims || [];
        return claims
            .filter(c => c.typ === 'roles' || c.typ === 'http://schemas.microsoft.com/ws/2008/06/identity/claims/role')
            .map(c => c.val);
    } catch (err) {
        console.error('Failed to parse x-ms-client-principal header:', err.message);
        return [];
    }
}

function callerCanRunScripts(req) {
    const roles = getCallerRoles(req);
    return roles.some(r => LOCKDOWN_ROLES.includes(r));
}

// Parses the ".SCRIPT_NAME" / ".SCRIPT_DESCRIPTION" / ".REQUIRES_USER"
// comment-header convention out of the top of a .ps1 file's text.
function parseScriptMetadata(fileName, content) {
    const nameMatch = content.match(/^\s*\.SCRIPT_NAME\s+(.+)$/m);
    const descMatch = content.match(/^\s*\.SCRIPT_DESCRIPTION\s+(.+)$/m);
    const reqUserMatch = content.match(/^\s*\.REQUIRES_USER\s+(true|false)$/m);

    return {
        id: fileName,
        name: nameMatch ? nameMatch[1].trim() : fileName.replace(/\.ps1$/i, ''),
        description: descMatch ? descMatch[1].trim() : '',
        requiresUser: reqUserMatch ? reqUserMatch[1].trim().toLowerCase() === 'true' : false
    };
}

// Lists available custom scripts for the dashboard to render as buttons.
app.get('/api/scripts', (req, res) => {
    fs.readdir(SCRIPTS_DIR, (err, files) => {
        if (err) {
            if (err.code === 'ENOENT') return res.json([]); // no scripts folder yet
            console.error('Error reading scripts directory:', err.message);
            return res.status(500).json({ error: 'Failed to list scripts.' });
        }

        const scripts = files
            .filter(f => SCRIPT_NAME_PATTERN.test(f))
            .map(fileName => {
                try {
                    const content = fs.readFileSync(path.join(SCRIPTS_DIR, fileName), 'utf8');
                    return parseScriptMetadata(fileName, content);
                } catch (readErr) {
                    console.error('Error reading script ' + fileName + ':', readErr.message);
                    return null;
                }
            })
            .filter(Boolean);

        res.json(scripts);
    });
});

// Executes a custom script by file name. The selected user's UPN (if any)
// is passed as a -UserPrincipalName parameter; static app configuration
// (tenant/client id, client secret) is available to the script via the
// same environment variables the Node process already has, since spawned
// child processes inherit process.env by default.
app.post('/api/scripts/run', (req, res) => {
    if (!callerCanRunScripts(req)) {
        return res.status(403).json({ error: 'You do not have permission to run scripts.' });
    }

    const fileName = req.body && req.body.fileName;
    const upn = req.body && req.body.upn;

    if (!fileName || !SCRIPT_NAME_PATTERN.test(fileName)) {
        return res.status(400).json({ error: 'Invalid script file name.' });
    }

    const scriptPath = path.join(SCRIPTS_DIR, fileName);
    const resolved = path.resolve(scriptPath);
    if (!resolved.startsWith(path.resolve(SCRIPTS_DIR) + path.sep)) {
        return res.status(400).json({ error: 'Invalid script path.' });
    }
    if (!fs.existsSync(resolved)) {
        return res.status(404).json({ error: 'Script not found.' });
    }

    const args = ['-NoProfile', '-NonInteractive', '-ExecutionPolicy', 'Bypass', '-File', resolved];
    if (upn) {
        args.push('-UserPrincipalName', upn);
    }

    const child = spawn('powershell.exe', args, { windowsHide: true });

    let stdout = '';
    let stderr = '';
    child.stdout.on('data', d => { stdout += d.toString(); });
    child.stderr.on('data', d => { stderr += d.toString(); });

    child.on('error', (err) => {
        console.error('Failed to start script ' + fileName + ':', err.message);
        res.status(500).json({ error: 'Failed to start script: ' + err.message });
    });

    child.on('close', (code) => {
        res.json({
            fileName,
            exitCode: code,
            output: stdout,
            errorOutput: stderr,
            success: code === 0
        });
    });
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
