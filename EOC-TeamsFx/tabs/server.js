const express = require('express');
const path = require('path');
const crypto = require('crypto');
const axios = require('axios');
const { OnBehalfOfCredential } = require('@azure/identity');

const app = express();
const PORT = process.env.PORT || 3978;

// OBO token cache: sha256(ssoToken) → { token, expiry }
const oboCache = new Map();
const CACHE_TTL_MS = 45 * 60 * 1000; // 45 minutes

// Prune expired entries on the same cadence so the Map never grows unbounded
setInterval(() => {
    const now = Date.now();
    for (const [key, entry] of oboCache) {
        if (entry.expiry <= now) oboCache.delete(key);
    }
}, CACHE_TTL_MS).unref();

app.use(express.static(path.join(__dirname, 'build')));
app.use('/api/graph', express.json());

function hashToken(token) {
    return crypto.createHash('sha256').update(token).digest('hex');
}

async function resolveGraphToken(ssoToken) {
    const key = hashToken(ssoToken);
    const cached = oboCache.get(key);
    if (cached && cached.expiry > Date.now()) return cached.token;
    const credential = new OnBehalfOfCredential({
        tenantId: process.env.M365_TENANT_ID,
        clientId: process.env.M365_CLIENT_ID,
        clientSecret: process.env.M365_CLIENT_SECRET,
        userAssertionToken: ssoToken
    });
    const tokenResponse = await credential.getToken('https://graph.microsoft.com/.default');
    if (!tokenResponse) throw new Error('OBO returned null token');
    oboCache.set(key, {
        token: tokenResponse.token,
        expiry: Math.min(Date.now() + CACHE_TTL_MS, tokenResponse.expiresOnTimestamp ?? Date.now() + CACHE_TTL_MS)
    });
    return tokenResponse.token;
}

function extractSsoToken(req) {
    const auth = req.headers['authorization'];
    if (!auth || !auth.startsWith('Bearer ')) return null;
    return auth.slice(7);
}

// User-profile endpoint — returns user object only, no tokens exposed
app.get('/api/graph/me', async (req, res) => {
    const ssoToken = extractSsoToken(req);
    if (!ssoToken) return res.status(401).json({ error: 'Missing Authorization header' });
    try {
        const graphToken = await resolveGraphToken(ssoToken);
        const { data } = await axios.get('https://graph.microsoft.com/v1.0/me', {
            headers: { Authorization: `Bearer ${graphToken}` }
        });
        res.json(data);
    } catch (err) {
        console.error('[proxy] /api/graph/me error:', err.message);
        res.status(500).json({ error: 'Auth proxy failed' });
    }
});

// Cache-warming endpoint — evicts current entry and re-does OBO so the next
// real Graph call hits a warm cache rather than paying the OBO latency
app.post('/api/auth/refresh', async (req, res) => {
    const ssoToken = extractSsoToken(req);
    if (!ssoToken) return res.status(401).end();
    try {
        oboCache.delete(hashToken(ssoToken));
        await resolveGraphToken(ssoToken);
        res.sendStatus(204);
    } catch (err) {
        console.error('[proxy] /api/auth/refresh error:', err.message);
        res.sendStatus(500);
    }
});

const ALLOWED_GRAPH_PATHS = [
    /^\/v1\.0\/me(\/.*)?$/,
    /^\/v1\.0\/sites(\/.*)?$/,
    /^\/v1\.0\/groups(\/.*)?$/,
    /^\/v1\.0\/teams(\/.*)?$/,
    /^\/v1\.0\/users(\/.*)?$/,
    /^\/v1\.0\/planner\//,
    /^\/v1\.0\/appCatalogs\//,
    /^\/v1\.0\/invitations$/,
];

// Allowlist for generic proxy — /api/graph/v1.0/* only; does not affect /api/graph/me
app.use('/api/graph/v1.0', (req, res, next) => {
    if (!ALLOWED_GRAPH_PATHS.some(re => re.test('/v1.0' + req.path))) {
        return res.status(403).json({ error: 'Path not permitted' });
    }
    next();
});

// Generic Graph pass-through — strips /api/graph, forwards remainder to Graph
app.all('/api/graph/*', async (req, res) => {
    const ssoToken = extractSsoToken(req);
    if (!ssoToken) return res.status(401).json({ error: 'Missing Authorization header' });
    try {
        const graphToken = await resolveGraphToken(ssoToken);
        const graphPath = req.path.replace(/^\/api\/graph/, '');
        const qs = req.url.split('?').slice(1).join('?');
        const graphUrl = `https://graph.microsoft.com${graphPath}${qs ? '?' + qs : ''}`;
        const graphRes = await axios({
            method: req.method,
            url: graphUrl,
            headers: {
                Authorization: `Bearer ${graphToken}`,
                ...(req.headers['content-type'] && { 'Content-Type': req.headers['content-type'] })
            },
            data: ['GET', 'HEAD', 'DELETE'].includes(req.method.toUpperCase()) ? undefined : req.body,
            responseType: 'arraybuffer',
            validateStatus: () => true
        });
        if (graphRes.status === 204 || !graphRes.data || graphRes.data.length === 0) {
            return res.sendStatus(graphRes.status);
        }
        const upstreamType = graphRes.headers['content-type'];
        if (upstreamType) res.set('Content-Type', upstreamType);
        res.status(graphRes.status).send(Buffer.from(graphRes.data));
    } catch (err) {
        console.error('[proxy] Graph proxy error:', err.message);
        res.status(500).json({ error: 'Graph proxy failed' });
    }
});

app.get('*', (_req, res) => {
    res.sendFile(path.join(__dirname, 'build', 'index.html'));
});

app.listen(PORT, () => console.log(`Server listening on port ${PORT}`));
