/**
 * server.js — Test Plan Generator Local Server
 * ─────────────────────────────────────────────
 * Depends on Node.js built-ins + optional 'docx' package for Word export.
 *
 * Routes:
 *   GET  /ping              → health check
 *   GET  /*                 → serve files from ./public/
 *   POST /jira-proxy        → proxy Jira API calls (avoids CORS)
 *   POST /testrail-proxy    → proxy TestRail API calls (avoids CORS)
 *
 * Run: node server.js
 * Then open: http://localhost:3377
 */

const http  = require('http');
const https = require('https');
const fs    = require('fs');
const crypto = require('crypto');
const { execFile } = require('child_process');
const os   = require('os');
const { BedrockRuntimeClient, ConverseCommand } = require('@aws-sdk/client-bedrock-runtime');
let docxLib;
function loadDocx() {
  // Try cjs build first (most reliable), then fallback to package root
  const attempts = [
    () => require('./node_modules/docx/build/index.cjs'),
    () => require('docx'),
  ];
  for (const attempt of attempts) {
    try {
      const lib = attempt();
      if (lib && lib.Paragraph) return lib;
    } catch(e) {}
  }
  return null;
}
try {
  docxLib = loadDocx();
  if (!docxLib) {
    console.warn('[DOCX] Library not found. DOCX export disabled until dependencies are installed (npm install).');
  }
} catch(e) {
  console.warn('[DOCX] Load error:', e.message);
  docxLib = null;
}
const path  = require('path');

const PORT       = Number(process.env.PORT || 3377);
const HOST       = process.env.HOST || '127.0.0.1';
const PUBLIC_DIR = path.join(__dirname, 'public');

// ── MIME types ─────────────────────────────────────────
const MIME = {
  '.html': 'text/html; charset=utf-8',
  '.js':   'text/javascript',
  '.css':  'text/css',
  '.json': 'application/json',
  '.png':  'image/png',
  '.ico':  'image/x-icon',
  '.svg':  'image/svg+xml',
};

// ── Read POST body ──────────────────────────────────────
function readBody(req) {
  return new Promise((resolve, reject) => {
    let raw = '';
    req.on('data', chunk => raw += chunk);
    req.on('end', () => {
      try   { resolve(JSON.parse(raw)); }
      catch { reject(new Error('Invalid JSON')); }
    });
    req.on('error', reject);
  });
}

// ── Make outbound HTTPS call ────────────────────────────
function httpsCall(targetUrl, method, extraHeaders, body) {
  return new Promise((resolve, reject) => {
    let hostname, port, path;

    try {
      // TestRail uses: https://host/index.php?/api/v2/endpoint
      // The ? is part of the path (not a query separator for us)
      // We parse manually to keep the full path intact
      const withProto = targetUrl.startsWith('http') ? targetUrl : 'https://' + targetUrl;

      // Extract protocol + host + everything after
      const m = withProto.match(/^https?:\/\/([^/]+)(\/.*)$/s);
      if (!m) throw new Error('Bad URL: ' + targetUrl);

      const hostStr = m[1];  // e.g. "upland.testrail.com" or "host:443"
      path = m[2] || '/';    // e.g. "/index.php?/api/v2/get_current_user"

      const hp = hostStr.split(':');
      hostname = hp[0];
      port     = hp[1] ? parseInt(hp[1]) : 443;

      console.log('[HTTP] host=' + hostname + ':' + port);
      console.log('[HTTP] path=' + path);
    } catch(e) {
      return reject(new Error('URL parse error: ' + targetUrl + ' → ' + e.message));
    }

    const bodyStr = (body && method !== 'GET') ? JSON.stringify(body) : '';
    const headers = {
      'Accept':       'application/json',
      'Content-Type': 'application/json',
      ...extraHeaders,
    };
    if (bodyStr) headers['Content-Length'] = Buffer.byteLength(bodyStr);

    // TestRail uses /index.php?/api/v2/ — the ? must stay unencoded in path
    // http.request sends path as-is which is correct
    const req = https.request({
      hostname,
      port,
      path,
      method,
      headers,
      rejectUnauthorized: false  // Handle self-signed certs on some installs
    }, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => resolve({ status: res.statusCode, headers: res.headers, body: data }));
    });

    req.on('error', reject);
    req.setTimeout(20000, () => { req.destroy(); reject(new Error('Request timed out')); });
    if (bodyStr) req.write(bodyStr);
    req.end();
  });
}

function httpsCallBuffer(targetUrl, method, extraHeaders, rawBody) {
  return new Promise((resolve, reject) => {
    let hostname, port, path;
    try {
      const withProto = targetUrl.startsWith('http') ? targetUrl : 'https://' + targetUrl;
      const m = withProto.match(/^https?:\/\/([^/]+)(\/.*)$/s);
      if (!m) throw new Error('Bad URL: ' + targetUrl);
      const hostStr = m[1];
      path = m[2] || '/';
      const hp = hostStr.split(':');
      hostname = hp[0];
      port = hp[1] ? parseInt(hp[1], 10) : 443;
    } catch (e) {
      return reject(new Error('URL parse error: ' + targetUrl + ' → ' + e.message));
    }

    const headers = {
      ...extraHeaders,
    };

    const req = https.request({
      hostname,
      port,
      path,
      method,
      headers,
      rejectUnauthorized: false,
    }, res => {
      const chunks = [];
      res.on('data', chunk => chunks.push(Buffer.from(chunk)));
      res.on('end', () => resolve({ status: res.statusCode, headers: res.headers, body: Buffer.concat(chunks) }));
    });

    req.on('error', reject);
    req.setTimeout(30000, () => { req.destroy(); reject(new Error('Request timed out')); });
    if (rawBody) req.write(rawBody);
    req.end();
  });
}

// ── CORS ────────────────────────────────────────────────
function setCORS(req, res) {
  res.setHeader('Access-Control-Allow-Origin',  '*');
  res.setHeader('Access-Control-Allow-Methods', 'GET, POST, OPTIONS');
  res.setHeader('Access-Control-Allow-Headers', 'Content-Type, Authorization');
}

// ── JSON response ───────────────────────────────────────
function json(res, status, data) {
  res.writeHead(status, { 'Content-Type': 'application/json' });
  res.end(JSON.stringify(data));
}

// ── /ping ───────────────────────────────────────────────
function handlePing(res) {
  json(res, 200, { ok: true, server: 'Test Plan Generator', port: PORT });
}

// ── /jira-proxy ─────────────────────────────────────────
async function handleJiraProxy(req, res) {
  let body;
  try { body = await readBody(req); }
  catch(e) { return json(res, 400, { error: e.message }); }

  const userConfig = getPersonalConfigForRequest(req);
  const baseUrl = body.baseUrl || userConfig.jira_url;
  const issueKey = body.issueKey;
  const email = body.email || userConfig.jira_email;
  const token = body.token || userConfig.jira_token;

  if (!baseUrl || !issueKey || !email || !token)
    return json(res, 400, { error: 'Missing: baseUrl, issueKey, email, token' });

  const apiUrl = `${baseUrl.replace(/\/$/, '')}/rest/api/3/issue/${issueKey}` +
                 `?fields=*all&expand=names`;
  const auth   = Buffer.from(`${email}:${token}`).toString('base64');

  console.log(`[Jira] Fetching ${issueKey}...`);
  try {
    const r = await httpsCall(apiUrl, 'GET', { 'Authorization': `Basic ${auth}` }, null);
    console.log(`[Jira] ${r.status}`);
    if (r.status === 401) return json(res, 401, {
      error: '401 Unauthorized — Check your Jira email and API token.\n' +
             'Get token: Jira → Profile → Security → API Tokens → Create'
    });
    if (r.status === 403) return json(res, 403, {
      error: '403 Forbidden — You do not have access to this Jira project'
    });
    if (r.status === 404) return json(res, 404, {
      error: `404 Not Found — Issue "${issueKey}" does not exist or you have no access`
    });
    res.writeHead(r.status, { 'Content-Type': 'application/json' });
    res.end(r.body);
  } catch(e) {
    console.error('[Jira] Error:', e.message);
    json(res, 500, { error: 'Jira request failed: ' + e.message });
  }
}

async function handleJiraAttachmentProxy(req, res) {
  let body;
  try { body = await readBody(req); }
  catch(e) { return json(res, 400, { error: e.message }); }

  const userConfig = getPersonalConfigForRequest(req);
  const baseUrl = body.baseUrl || userConfig.jira_url;
  const { attachmentUrl, attachmentId, name, mimeType } = body;
  const email = body.email || userConfig.jira_email;
  const token = body.token || userConfig.jira_token;
  if (!baseUrl || (!attachmentUrl && !attachmentId) || !email || !token) {
    return json(res, 400, { error: 'Missing: baseUrl, attachmentUrl/attachmentId, email, token' });
  }

  let targetUrl = '';
  try {
    const base = new URL(baseUrl);
    if (attachmentId) {
      targetUrl = `${base.origin}/rest/api/3/attachment/content/${attachmentId}`;
    } else {
      const att = new URL(attachmentUrl);
      const host = att.host.toLowerCase();
      const trusted = host.endsWith('.atlassian.net') || host.endsWith('.atlassian.com') || host === base.host.toLowerCase();
      if (!trusted) return json(res, 400, { error: 'Untrusted attachment host' });
      targetUrl = att.toString();
    }
  } catch (e) {
    return json(res, 400, { error: 'Invalid attachment URL' });
  }

  const auth = Buffer.from(`${email}:${token}`).toString('base64');

  async function fetchWithRedirects(url, maxRedirects) {
    let current = url;
    for (let i = 0; i <= maxRedirects; i++) {
      const r = await httpsCallBuffer(current, 'GET', {
        'Authorization': `Basic ${auth}`,
        'Accept': '*/*',
      }, null);

      const location = r.headers?.location || r.headers?.Location;
      if ([301,302,303,307,308].includes(r.status) && location) {
        current = location.startsWith('http') ? location : (new URL(location, current)).toString();
        continue;
      }
      return r;
    }
    throw new Error('Too many redirects while fetching Jira attachment');
  }

  try {
    const r = await fetchWithRedirects(targetUrl, 4);

    if (r.status === 401) return json(res, 401, { error: 'Unauthorized to read Jira attachment' });
    if (r.status === 403) return json(res, 403, { error: 'Forbidden to read Jira attachment' });
    if (r.status !== 200) return json(res, r.status || 500, { error: 'Jira attachment returned ' + r.status });

    const contentType = String((r.headers && (r.headers['content-type'] || r.headers['Content-Type'])) || mimeType || 'image/png');
    if (!contentType.startsWith('image/')) {
      return json(res, 400, { error: 'Attachment is not an image' });
    }

    const b64 = r.body.toString('base64');
    return json(res, 200, {
      ok: true,
      name: name || 'jira-image',
      type: contentType,
      dataUrl: `data:${contentType};base64,${b64}`,
    });
  } catch (e) {
    return json(res, 500, { error: 'Jira attachment proxy failed: ' + e.message });
  }
}

// ── Local accounts and integration storage ───────────────
const SESSION_COOKIE = 'qa_session';
const SECURE_SESSION_COOKIE = process.env.COOKIE_SECURE === 'true';
const userSessions = new Map();
const LOCAL_STORE_FILE = process.env.APP_USER_STORE || path.join(__dirname, 'data', 'local-store.json');
const DEFAULT_BEDROCK_MODEL = 'global.anthropic.claude-sonnet-5-5';
const LEGACY_BEDROCK_MODEL = 'anthropic.claude-3-5-sonnet-20241022-v2:0';
let localStore = { users: [], sharedAi: { bedrock: {} } };

function loadLocalStore() {
  try {
    if (!fs.existsSync(LOCAL_STORE_FILE)) return;
    const stored = JSON.parse(fs.readFileSync(LOCAL_STORE_FILE, 'utf8'));
    localStore = {
      users: Array.isArray(stored.users) ? stored.users : [],
      sharedAi: { bedrock: stored.sharedAi?.bedrock || {} }
    };
    let modelUpdated = false;
    if (!localStore.sharedAi.bedrock.modelId || localStore.sharedAi.bedrock.modelId === LEGACY_BEDROCK_MODEL) {
      localStore.sharedAi.bedrock.modelId = DEFAULT_BEDROCK_MODEL;
      modelUpdated = true;
    }
    for (const user of localStore.users) {
      if (user.aiProvider === 'bedrock' && user.aiModel === LEGACY_BEDROCK_MODEL) {
        user.aiModel = DEFAULT_BEDROCK_MODEL;
        modelUpdated = true;
      }
    }
    if (modelUpdated) saveLocalStore();
  } catch (e) {
    console.error('[Users] Local store could not be loaded:', e.message);
    throw e;
  }
}

function saveLocalStore() {
  fs.mkdirSync(path.dirname(LOCAL_STORE_FILE), { recursive: true });
  fs.writeFileSync(LOCAL_STORE_FILE, JSON.stringify(localStore, null, 2), { mode: 0o600 });
}

function findStoredUser(userId) {
  return localStore.users.find(user => user.id === userId) || null;
}

function publicUser(user) {
  return {
    id: user.id,
    email: user.email,
    displayName: user.displayName,
    role: user.role,
    isActive: user.isActive,
    createdAt: user.createdAt,
    lastLoginAt: user.lastLoginAt || null
  };
}

function getUserAppConfig(user) {
  const integrations = user?.integrations || {};
  const tr = integrations.testrail || {};
  const jira = integrations.jira || {};
  const github = integrations.github || {};
  return {
    tr_url: tr.url || '', tr_project: tr.project || '', tr_suite: tr.suite || '',
    tr_email: tr.email || '', tr_api_key: tr.apiKey || '',
    jira_url: jira.url || '', jira_email: jira.email || '', jira_token: jira.token || '',
    ai_provider: user?.aiProvider || 'github',
    github_token: github.token || '',
    ai_model: user?.aiModel || (user?.aiProvider === 'bedrock' ? localStore.sharedAi.bedrock.modelId || DEFAULT_BEDROCK_MODEL : github.model || 'gpt-4.1'),
    ai_endpoint: github.endpoint || 'https://models.github.com/chat/completions',
    ai_max_tokens: github.maxTokens || ''
  };
}

function getSessionUser(req) {
  const current = getCurrentUser(req);
  return current ? findStoredUser(current.userId) : null;
}

function getPersonalConfigForRequest(req) {
  return getUserAppConfig(getSessionUser(req));
}

function migrateLegacyIntegrations() {
  const legacy = savedConfig || {};
  return {
    testrail: {
      url: legacy.tr_url || '', project: legacy.tr_project || '', suite: legacy.tr_suite || '',
      email: legacy.tr_email || '', apiKey: legacy.tr_api_key || ''
    },
    jira: { url: legacy.jira_url || '', email: legacy.jira_email || '', token: legacy.jira_token || '' },
    github: {
      token: legacy.github_token || '', model: legacy.ai_model || 'gpt-4.1',
      endpoint: legacy.ai_endpoint || 'https://models.github.com/chat/completions',
      maxTokens: legacy.ai_max_tokens || ''
    }
  };
}

function clearLegacyIntegrationsFromConfig() {
  for (const key of ['tr_url', 'tr_project', 'tr_suite', 'tr_email', 'tr_api_key', 'tr_password', 'tr_cookie', 'jira_url', 'jira_email', 'jira_token', 'github_token', 'ai_key', 'ai_provider', 'ai_model', 'ai_endpoint', 'ai_max_tokens']) {
    delete savedConfig[key];
  }
  try { fs.writeFileSync(CONFIG_FILE, JSON.stringify(savedConfig, null, 2)); }
  catch (e) { console.warn('[Config] Legacy integration values were migrated but could not be removed:', e.message); }
}

function toClientIntegrations(user) {
  const integrations = user.integrations || {};
  return {
    testrail: { ...(integrations.testrail || {}) },
    jira: { ...(integrations.jira || {}) },
    github: { ...(integrations.github || {}) }
  };
}

function parseCookies(header = '') {
  const out = {};
  for (const part of String(header).split(';')) {
    const clean = part.trim();
    if (!clean) continue;
    const idx = clean.indexOf('=');
    if (idx === -1) continue;
    const key = clean.slice(0, idx);
    const val = clean.slice(idx + 1);
    out[key] = decodeURIComponent(val);
  }
  return out;
}

function generateSessionToken() {
  return crypto.randomBytes(32).toString('hex');
}

function getCurrentUser(req) {
  const cookies = parseCookies(req.headers.cookie || '');
  const token = cookies[SESSION_COOKIE];
  if (!token) return null;
  const session = userSessions.get(token);
  if (!session) return null;
  const user = findStoredUser(session.userId);
  if (!user || !user.isActive) return null;
  return { userId: user.id, email: user.email, displayName: user.displayName, role: user.role, loginTime: session.loginTime };
}

function setSessionCookie(res, token) {
  const secure = SECURE_SESSION_COOKIE ? '; Secure' : '';
  res.setHeader('Set-Cookie', `${SESSION_COOKIE}=${token}; Path=/; HttpOnly; SameSite=Lax; Max-Age=86400${secure}`);
}

function clearSessionCookie(res) {
  const secure = SECURE_SESSION_COOKIE ? '; Secure' : '';
  res.setHeader('Set-Cookie', `${SESSION_COOKIE}=; Path=/; HttpOnly; SameSite=Lax; Max-Age=0${secure}`);
}

function hashPassword(password) {
  const salt = crypto.randomBytes(16).toString('hex');
  const hash = crypto.pbkdf2Sync(password, salt, 100000, 64, 'sha512').toString('hex');
  return `${salt}:${hash}`;
}

function verifyPassword(password, hashed) {
  if (!password || !hashed || typeof hashed !== 'string') return false;
  const idx = hashed.indexOf(':');
  if (idx < 1) return false;
  const salt = hashed.slice(0, idx);
  const expected = hashed.slice(idx + 1);
  const actual = crypto.pbkdf2Sync(password, salt, 100000, 64, 'sha512').toString('hex');
  return crypto.timingSafeEqual(Buffer.from(actual), Buffer.from(expected));
}

const CONFIG_FILE = process.env.APP_CONFIG_FILE || path.join(__dirname, 'config.json');
let savedConfig   = {};

function loadConfig() {
  try {
    if (fs.existsSync(CONFIG_FILE)) {
      savedConfig = JSON.parse(fs.readFileSync(CONFIG_FILE, 'utf8'));
      console.log('[Config] Loaded:', CONFIG_FILE);
    }
  } catch(e) { console.error('[Config] Load error:', e.message); }
}

// ── TestRail Basic auth header helper ───────────────
function trBasicAuth(email, apiKey) {
  if (!email || !apiKey) return null;
  return 'Basic ' + Buffer.from(email.trim() + ':' + apiKey.trim()).toString('base64');
}

// ── List all sections (for UI dropdown) ──────────────
async function handleGetMilestones(req, res) {
  const config = getPersonalConfigForRequest(req);
  const trUrl  = (config.tr_url || '').replace(/\/$/, '');
  const projId = config.tr_project || '';
  const trEmail  = config.tr_email || '';
  const trApiKey = config.tr_api_key || '';
  if (!trUrl || !projId || !trEmail || !trApiKey) {
    return json(res, 200, { milestones: [] });
  }
  const auth = trBasicAuth(trEmail, trApiKey);
  if (!auth) return json(res, 200, { milestones: [], warning: 'No TestRail API key configured' });

  const headers = { 'Accept': 'application/json', 'Content-Type': 'application/json', 'Authorization': auth };

  async function fetchAllMilestones(isCompleted) {
    // TestRail paginates — fetch all pages (limit 250 per page)
    let all = [], offset = 0;
    while (true) {
      const url = `${trUrl}/index.php?/api/v2/get_milestones/${projId}&is_completed=${isCompleted}&limit=250&offset=${offset}`;
      console.log(`[Milestones] GET ${url}`);
      const r = await httpsCallRaw(url, 'GET', headers, null);
      console.log(`[Milestones] status=${r.status} body=${r.body.slice(0,200)}`);
      if (r.status !== 200) break;
      const data = JSON.parse(r.body);
      // TestRail returns {milestones:[...], _links:{next:...}} or just array
      const page = data.milestones || (Array.isArray(data) ? data : []);
      all.push(...page);
      // Stop if no next page or got fewer than 250
      if (!data._links?.next || page.length < 250) break;
      offset += 250;
    }
    return all;
  }

  try {
    // Active and pending only (exclude completed)
    const active = await fetchAllMilestones(0);
    const topLevel = [...active];
    console.log('[Milestones] active/pending:', active.length);

    // Flatten: each top-level milestone may embed sub-milestones
    let all = [];
    for (const m of topLevel) {
      all.push(m);
      if (Array.isArray(m.milestones) && m.milestones.length) {
        all.push(...m.milestones);
      }
    }
    console.log('[Milestones] total:', all.length);

    const milestones = all
      .filter(m => !m.is_completed)
      .map(m => ({
      id: m.id,
      name: m.name,
      due_on: m.due_on || null,
      started_on: m.started_on || null,
      is_completed: m.is_completed || false,
      parent_id: m.parent_id || null,
      url: m.url || null,
    }));
    json(res, 200, { milestones });
  } catch(e) {
    console.error('[Milestones] Error:', e.message);
    json(res, 200, { milestones: [], error: e.message });
  }
}

async function handleGetSections(req, res) {
  const config = getPersonalConfigForRequest(req);
  const trUrl   = config.tr_url || '';
  const projId  = config.tr_project || '';
  // Accept suite_id from query string (e.g. /get-sections?suite_id=123) or fall back to config
  const qs      = new URLSearchParams(req.url.includes('?') ? req.url.split('?')[1] : '');
  const suiteId = qs.get('suite_id') || config.tr_suite || '';

  if (!trUrl || !projId)
    return json(res, 200, { sections: [], warning: 'TestRail not configured' });
  if (!suiteId)
    return json(res, 200, { sections: [], warning: 'No suite_id provided' });

  try {
    const trEmail  = config.tr_email || '';
    const trApiKey = config.tr_api_key || '';
    const auth = trBasicAuth(trEmail, trApiKey);
    if (!auth) return json(res, 200, { sections: [], warning: 'No TestRail API key configured' });
    const authHeaders = { 'Accept': 'application/json', 'Content-Type': 'application/json', 'Authorization': auth };
    const allSections = [];
    const seenIds = new Set();
    let offset = 0;
    const limit = 250;
    let pageCount = 0;
    while (true) {
      const url = `${trUrl.replace(/\/$/, '')}/index.php?/api/v2/get_sections/${projId}&suite_id=${suiteId}&limit=${limit}&offset=${offset}`;
      const r = await httpsCallRaw(url, 'GET', authHeaders, null);
      if (r.status !== 200) {
        if (allSections.length === 0) return json(res, 200, { sections: [], warning: 'TestRail returned ' + r.status });
        break;
      }
      const data = JSON.parse(r.body);
      const page = Array.isArray(data) ? data : (data.sections || []);
      pageCount += 1;
      let newOnPage = 0;
      for (const s of page) {
        const sid = String(s.id);
        if (!seenIds.has(sid)) {
          seenIds.add(sid);
          allSections.push(s);
          newOnPage += 1;
        }
      }

      // Stop conditions:
      // 1) empty page
      // 2) partial page (< limit)
      // 3) page had no new IDs (offset ignored by server / repeated page)
      // 4) safety guard against infinite loops
      if (page.length === 0 || page.length < limit || newOnPage === 0 || pageCount >= 40) break;
      offset += limit;
    }
    const sections = allSections;
    console.log(`[Sections] suite=${suiteId} loaded=${sections.length}`);
    json(res, 200, { sections: sections.map(s => ({ id: s.id, name: s.name, depth: s.depth || 0, parent_id: s.parent_id })) });
  } catch(e) {
    json(res, 500, { error: e.message });
  }
}

// ── /test-testrail — verify email+apiKey credentials ────────────────
async function handleTestTestrail(req, res) {
  let body = {};
  if (req.method === 'POST') {
    try { body = await readBody(req); } catch(e) {}
  }
  const config = getPersonalConfigForRequest(req);
  const trUrl   = (body.tr_url    || config.tr_url     || '').trim();
  const trEmail = (body.tr_email  || config.tr_email   || '').trim();
  const trApiKey= (body.tr_api_key|| config.tr_api_key || '').trim();
  if (!trUrl)    return json(res, 200, { ok: false, error: 'TestRail URL not set.' });
  if (!trEmail)  return json(res, 200, { ok: false, error: 'TestRail Email not set.' });
  if (!trApiKey) return json(res, 200, { ok: false, error: 'TestRail API Key not set.' });

  console.log(`[TestRail Test] url=${trUrl} email=${trEmail}`);

  const url  = `${trUrl.replace(/\/$/, '')}/index.php?/api/v2/get_current_user`;
  const auth = Buffer.from(trEmail + ':' + trApiKey).toString('base64');
  try {
    const r = await httpsCallRaw(url, 'GET', {
      'Authorization': 'Basic ' + auth,
      'Accept': 'application/json',
      'Content-Type': 'application/json',
    }, null);
    console.log(`[TestRail Test] status=${r.status} body=${r.body.slice(0,200)}`);
    if (r.status === 200) {
      const data = JSON.parse(r.body);
      return json(res, 200, { ok: true, name: data.name || data.email || 'Connected' });
    }
    const errBody = (() => { try { return JSON.parse(r.body); } catch(e) { return {}; } })();
    const errMsg = errBody.error || `HTTP ${r.status}`;
    // Give a clear hint for lockout vs wrong key
    if (errMsg.toLowerCase().includes('failed login') || errMsg.toLowerCase().includes('10 minutes')) {
      return json(res, 200, { ok: false, error: '⏳ Account locked out — wait 10 minutes then try again.' });
    }
    return json(res, 200, { ok: false, error: errMsg });
  } catch(e) {
    return json(res, 200, { ok: false, error: e.message });
  }
}

// ── /get-suites — return all suites for the configured project ────────
async function handleGetSuites(req, res) {
  const config = getPersonalConfigForRequest(req);
  const trUrl   = config.tr_url || '';
  const projId  = config.tr_project || '';
  const trEmail = config.tr_email || '';
  const trApiKey= config.tr_api_key || '';

  if (!trUrl || !projId) return json(res, 200, { suites: [] });

  const auth = trBasicAuth(trEmail, trApiKey);
  if (!auth) return json(res, 200, { suites: [], warning: 'No TestRail API key configured' });
  try {
    const url = `${trUrl.replace(/\/$/, '')}/index.php?/api/v2/get_suites/${projId}`;
    const r   = await httpsCallRaw(url, 'GET', { 'Accept': 'application/json', 'Content-Type': 'application/json', 'Authorization': auth }, null);
    if (r.status !== 200) return json(res, 200, { suites: [], warning: 'TestRail returned ' + r.status });
    const data   = JSON.parse(r.body);
    const suites = Array.isArray(data) ? data : (data.suites || []);
    json(res, 200, { suites: suites.map(s => ({ id: s.id, name: s.name, description: s.description || '' })) });
  } catch(e) {
    json(res, 500, { error: e.message, suites: [] });
  }
}

// ── /get-projects — return all TestRail projects user can access ─────
async function handleGetProjects(req, res) {
  let body = {};
  if (req.method === 'POST') {
    try { body = await readBody(req); }
    catch(e) { return json(res, 400, { error: e.message, projects: [] }); }
  }

  const config = getPersonalConfigForRequest(req);
  const trUrl    = (body.tr_url || config.tr_url || '').trim();
  const trEmail  = (body.tr_email || config.tr_email || '').trim();
  const trApiKey = (body.tr_api_key || config.tr_api_key || '').trim();

  if (!trUrl || !trEmail || !trApiKey) {
    return json(res, 200, { projects: [], warning: 'Missing TestRail URL/Email/API key' });
  }

  const auth = trBasicAuth(trEmail, trApiKey);
  if (!auth) return json(res, 200, { projects: [], warning: 'No TestRail API key configured' });

  try {
    const url = `${trUrl.replace(/\/$/, '')}/index.php?/api/v2/get_projects`;
    const r = await httpsCallRaw(url, 'GET', {
      'Accept': 'application/json',
      'Content-Type': 'application/json',
      'Authorization': auth,
    }, null);
    if (r.status !== 200) return json(res, 200, { projects: [], warning: 'TestRail returned ' + r.status });

    const data = JSON.parse(r.body);
    const projects = Array.isArray(data) ? data : (data.projects || []);
    return json(res, 200, {
      projects: projects.map(p => ({
        id: p.id,
        name: p.name,
        announcement: p.announcement || '',
        is_completed: !!p.is_completed,
        suite_mode: p.suite_mode,
      }))
    });
  } catch(e) {
    return json(res, 500, { error: e.message, projects: [] });
  }
}

// ── /get-config — expose safe config to browser ──────
const TOKEN_HISTORY_FILE = process.env.APP_TOKEN_HISTORY_FILE || path.join(__dirname, 'token-history.json');

function readTokenUsageHistory() {
  try {
    return fs.existsSync(TOKEN_HISTORY_FILE)
      ? JSON.parse(fs.readFileSync(TOKEN_HISTORY_FILE, 'utf8'))
      : [];
  } catch (e) {
    console.error('[Usage] History read failed:', e.message);
    return [];
  }
}

function writeTokenUsageHistory(history) {
  fs.writeFileSync(TOKEN_HISTORY_FILE, JSON.stringify(history, null, 2), 'utf8');
}

function recordTokenUsage(user, usage, model, context = {}) {
  const record = {
    id: crypto.randomUUID(),
    date: new Date().toISOString(),
    userId: user.id,
    userName: user.displayName,
    userEmail: user.email,
    userRole: user.role,
    provider: user.aiProvider || 'github',
    model: model || 'unknown',
    purpose: ['test-cases', 'test-plan'].includes(context.purpose) ? context.purpose : 'other',
    storyKey: String(context.storyKey || '').trim().slice(0, 200),
    storyId: String(context.storyId || '').trim().slice(0, 200),
    promptTokens: Number(usage.inputTokens ?? usage.prompt_tokens) || 0,
    completionTokens: Number(usage.outputTokens ?? usage.completion_tokens) || 0,
    totalTokens: Number(usage.totalTokens ?? usage.total_tokens) || 0,
    cost: null
  };
  if (!record.totalTokens) record.totalTokens = record.promptTokens + record.completionTokens;
  try {
    const history = readTokenUsageHistory();
    history.push(record);
    if (history.length > 10000) history.splice(0, history.length - 10000);
    writeTokenUsageHistory(history);
  } catch (e) {
    console.error('[Usage] History write failed:', e.message);
  }
  return record;
}

function handleGetTokenHistory(req, res) {
  const current = getCurrentUser(req);
  if (!current) return json(res, 401, { error: 'Not authenticated' });
  const history = readTokenUsageHistory().filter(record => record.userId === current.userId);
  return json(res, 200, history);
}

async function handleSaveTokenHistory(req, res) {
  const current = getCurrentUser(req);
  if (!current) return json(res, 401, { error: 'Not authenticated' });
  let body;
  try { body = await readBody(req); } catch(e) { return json(res, 400, { error: e.message }); }
  if (body?.action !== 'clear') return json(res, 400, { error: 'Usage records are managed by the server.' });
  try {
    const history = readTokenUsageHistory().filter(record => record.userId !== current.userId);
    writeTokenUsageHistory(history);
    return json(res, 200, { ok: true });
  } catch(e) {
    return json(res, 500, { error: e.message });
  }
}

function handleGetAdminUsage(req, res) {
  const current = getCurrentUser(req);
  if (!current || current.role !== 'Admin') return json(res, 403, { error: 'Admin access required.' });
  const query = new URL(req.url, `http://${req.headers.host}`).searchParams;
  const range = query.get('range') === 'month' ? 'month' : 'week';
  const now = new Date();
  const periodStart = range === 'month'
    ? new Date(Date.UTC(now.getUTCFullYear(), now.getUTCMonth(), 1))
    : new Date(now.getTime() - 7 * 24 * 60 * 60 * 1000);
  const records = readTokenUsageHistory().filter(record => {
    const date = new Date(record.date);
    return Number.isFinite(date.getTime()) && date >= periodStart;
  });
  const byUser = new Map();
  const getBucket = (userId, values = {}) => {
    if (!byUser.has(userId)) {
      byUser.set(userId, {
        userId,
        userName: values.userName || 'Unattributed legacy usage',
        userEmail: values.userEmail || '',
        role: values.userRole || '',
        calls: 0,
        promptTokens: 0,
        completionTokens: 0,
        totalTokens: 0,
        cost: 0,
        costKnown: true,
        stories: new Set(),
        plans: new Set(),
        models: new Set(),
        lastUsed: ''
      });
    }
    return byUser.get(userId);
  };

  localStore.users.forEach(user => getBucket(user.id, {
    userName: user.displayName,
    userEmail: user.email,
    userRole: user.role
  }));

  for (const record of records) {
    const bucket = getBucket(record.userId || 'legacy', record);
    bucket.calls++;
    bucket.promptTokens += Number(record.promptTokens) || 0;
    bucket.completionTokens += Number(record.completionTokens) || 0;
    bucket.totalTokens += Number(record.totalTokens) || 0;
    if (Number.isFinite(Number(record.cost))) bucket.cost += Number(record.cost);
    else bucket.costKnown = false;
    if (record.purpose === 'test-cases') bucket.stories.add(record.storyKey || record.storyId || record.id);
    if (record.purpose === 'test-plan') bucket.plans.add(record.storyKey || record.storyId || record.id);
    if (record.model) bucket.models.add(record.model);
    if (record.date > bucket.lastUsed) bucket.lastUsed = record.date;
  }

  const users = [...byUser.values()].map(bucket => ({
    ...bucket,
    stories: bucket.stories.size,
    plans: bucket.plans.size,
    models: [...bucket.models],
    cost: bucket.costKnown ? bucket.cost : null
  })).sort((a, b) => b.totalTokens - a.totalTokens);
  const totals = users.reduce((sum, user) => ({
    calls: sum.calls + user.calls,
    stories: sum.stories + user.stories,
    plans: sum.plans + user.plans,
    promptTokens: sum.promptTokens + user.promptTokens,
    completionTokens: sum.completionTokens + user.completionTokens,
    totalTokens: sum.totalTokens + user.totalTokens
  }), { calls: 0, stories: 0, plans: 0, promptTokens: 0, completionTokens: 0, totalTokens: 0 });

  return json(res, 200, { range, periodStart: periodStart.toISOString(), totals, users });
}

function handleGetConfig(req, res) {
  const current = getCurrentUser(req);
  if (!current) return json(res, 401, { error: 'Not authenticated' });
  const user = findStoredUser(current.userId);
  const config = getUserAppConfig(user);
  return json(res, 200, {
    ...config,
    tr_enabled: !!(config.tr_url && config.tr_api_key && config.tr_email),
    jira_enabled: !!(config.jira_url && config.jira_token),
    ai_enabled: config.ai_provider === 'bedrock'
      ? !!(localStore.sharedAi.bedrock.accessKeyId && localStore.sharedAi.bedrock.secretAccessKey)
      : !!config.github_token,
    bedrock: {
      configured: !!(localStore.sharedAi.bedrock.accessKeyId && localStore.sharedAi.bedrock.secretAccessKey),
      region: localStore.sharedAi.bedrock.region || 'us-east-1',
      modelId: localStore.sharedAi.bedrock.modelId || DEFAULT_BEDROCK_MODEL
    },
    user: publicUser(user)
  });
}

function handleGetAiSettings(req, res) {
  const current = getCurrentUser(req);
  if (!current) return json(res, 401, { error: 'Not authenticated' });
  const bedrock = localStore.sharedAi.bedrock || {};
  return json(res, 200, {
    isAdmin: current.role === 'Admin',
    configured: !!(bedrock.accessKeyId && bedrock.secretAccessKey),
    region: bedrock.region || 'us-east-1',
    modelId: bedrock.modelId || DEFAULT_BEDROCK_MODEL
  });
}

async function handleSaveAiSettings(req, res) {
  const current = getCurrentUser(req);
  if (!current || current.role !== 'Admin') return json(res, 403, { error: 'Admin access required.' });
  let body = {};
  try { body = await readBody(req); } catch (e) { return json(res, 400, { error: 'Invalid JSON body' }); }
  const bedrock = localStore.sharedAi.bedrock || (localStore.sharedAi.bedrock = {});
  if (body.region !== undefined) bedrock.region = String(body.region || '').trim();
  if (body.modelId !== undefined) bedrock.modelId = String(body.modelId || '').trim();
  if (body.accessKeyId) bedrock.accessKeyId = String(body.accessKeyId).trim();
  if (body.secretAccessKey) bedrock.secretAccessKey = String(body.secretAccessKey).trim();
  if (body.clearCredentials === true) {
    delete bedrock.accessKeyId;
    delete bedrock.secretAccessKey;
  }
  try { saveLocalStore(); }
  catch (e) { return json(res, 500, { error: 'Could not save shared Bedrock settings: ' + e.message }); }
  return json(res, 200, { ok: true, configured: !!(bedrock.accessKeyId && bedrock.secretAccessKey) });
}

async function handleLogin(req, res) {
  let body = {};
  try { body = await readBody(req); } catch (e) { return json(res, 400, { error: 'Invalid JSON body' }); }

  const email = String(body.email || '').trim().toLowerCase();
  const password = String(body.password || '');

  if (!email || !password) {
    return json(res, 400, { error: 'Email and password are required.' });
  }

  try {
    const user = localStore.users.find(item => item.email === email);
    if (!user) {
      return json(res, 401, { error: 'Invalid email or password.' });
    }
    if (!user.isActive) {
      return json(res, 403, { error: 'This user account is disabled.' });
    }

    if (!verifyPassword(password, user.passwordHash)) {
      return json(res, 401, { error: 'Invalid email or password.' });
    }

    const token = generateSessionToken();
    userSessions.set(token, {
      userId: user.id,
      loginTime: new Date().toISOString()
    });
    user.lastLoginAt = new Date().toISOString();
    saveLocalStore();

    setSessionCookie(res, token);
    return json(res, 200, {
      ok: true,
      user: {
        userId: user.id,
        email: user.email,
        displayName: user.displayName,
        role: user.role
      }
    });
  } catch (e) {
    console.error('[AUTH] Login failed:', e.message);
    return json(res, 500, { error: 'Login failed: ' + e.message });
  }
}

async function handleSetupStatus(req, res) {
  return json(res, 200, { setupRequired: localStore.users.length === 0 });
}

async function handleSetup(req, res) {
  let body = {};
  try { body = await readBody(req); } catch (e) { return json(res, 400, { error: 'Invalid JSON body' }); }
  if (localStore.users.length) return json(res, 409, { error: 'Initial setup has already been completed.' });

  const email = String(body.email || '').trim().toLowerCase();
  const displayName = String(body.displayName || '').trim();
  const password = String(body.password || '');
  if (!email || !displayName || !password) return json(res, 400, { error: 'Name, email, and password are required.' });
  if (password.length < 8) return json(res, 400, { error: 'Password must be at least 8 characters.' });

  const user = {
    id: crypto.randomUUID(),
    email,
    displayName,
    passwordHash: hashPassword(password),
    role: 'Admin',
    isActive: true,
    createdAt: new Date().toISOString(),
    lastLoginAt: new Date().toISOString(),
    aiProvider: 'github',
    integrations: migrateLegacyIntegrations()
  };
  localStore.users.push(user);
  try { saveLocalStore(); }
  catch (e) {
    localStore.users = [];
    return json(res, 500, { error: 'Could not save the initial account: ' + e.message });
  }
  clearLegacyIntegrationsFromConfig();

  const token = generateSessionToken();
  userSessions.set(token, { userId: user.id, loginTime: user.lastLoginAt });
  setSessionCookie(res, token);
  return json(res, 201, { ok: true, user: publicUser(user) });
}

async function handleLogout(req, res) {
  const cookies = parseCookies(req.headers.cookie || '');
  const token = cookies[SESSION_COOKIE];
  if (token) userSessions.delete(token);
  clearSessionCookie(res);
  return json(res, 200, { ok: true });
}

async function handleMe(req, res) {
  const current = getCurrentUser(req);
  if (!current) return json(res, 401, { error: 'Not authenticated' });
  return json(res, 200, { user: current });
}

async function handleCreateUser(req, res) {
  const current = getCurrentUser(req);
  if (!current || current.role !== 'Admin') {
    return json(res, 403, { error: 'Admin access required.' });
  }

  let body = {};
  try { body = await readBody(req); } catch (e) { return json(res, 400, { error: 'Invalid JSON body' }); }

  const email = String(body.email || '').trim().toLowerCase();
  const password = String(body.password || '');
  const displayName = String(body.displayName || body.name || email);
  const role = String(body.role || 'QA');

  if (!email || !password) {
    return json(res, 400, { error: 'Email and password are required.' });
  }
  if (password.length < 8) return json(res, 400, { error: 'Password must be at least 8 characters.' });
  if (!['Admin', 'QA'].includes(role)) return json(res, 400, { error: 'Role must be Admin or QA.' });

  if (localStore.users.some(user => user.email === email)) return json(res, 409, { error: 'User already exists.' });
  const user = {
    id: crypto.randomUUID(), email, displayName,
    passwordHash: hashPassword(password), role, isActive: true,
    createdAt: new Date().toISOString(), lastLoginAt: null,
    aiProvider: 'github',
    integrations: { testrail: {}, jira: {}, github: {} }
  };
  localStore.users.push(user);
  try { saveLocalStore(); }
  catch (e) {
    localStore.users = localStore.users.filter(item => item.id !== user.id);
    return json(res, 500, { error: 'User creation failed: ' + e.message });
  }
  return json(res, 201, { ok: true, user: publicUser(user) });
}

async function handleListUsers(req, res) {
  const current = getCurrentUser(req);
  if (!current || current.role !== 'Admin') {
    return json(res, 403, { error: 'Admin access required.' });
  }

  return json(res, 200, { users: localStore.users.map(publicUser).sort((a, b) => b.createdAt.localeCompare(a.createdAt)) });
}

async function handleGetMyIntegrations(req, res) {
  const current = getCurrentUser(req);
  if (!current) return json(res, 401, { error: 'Not authenticated' });
  const user = findStoredUser(current.userId);
  return json(res, 200, { integrations: toClientIntegrations(user) });
}

async function handleSaveIntegration(req, res) {
  const current = getCurrentUser(req);
  if (!current) return json(res, 401, { error: 'Not authenticated' });

  let body = {};
  try { body = await readBody(req); } catch (e) { return json(res, 400, { error: 'Invalid JSON body' }); }

  const user = findStoredUser(current.userId);
  if (!user) return json(res, 401, { error: 'Not authenticated' });
  const integrations = user.integrations || (user.integrations = {});
  const provider = String(body.provider || '').toLowerCase();
  if (!['testrail', 'jira', 'github'].includes(provider)) return json(res, 400, { error: 'Unsupported integration.' });
  integrations[provider] = { ...(integrations[provider] || {}), ...(body.values || {}) };
  try { saveLocalStore(); }
  catch (e) { return json(res, 500, { error: 'Could not save integration: ' + e.message }); }
  return json(res, 200, { ok: true });
}

async function handleSaveConfig(req, res) {
  let body;
  try { body = await readBody(req); }
  catch(e) { return json(res, 400, { error: e.message }); }
  const current = getCurrentUser(req);
  if (!current) return json(res, 401, { error: 'Not authenticated' });
  const user = findStoredUser(current.userId);
  const previous = getUserAppConfig(user);
  const value = key => body[key] === undefined ? previous[key] : String(body[key] || '').trim();
  const provider = body.ai_provider === 'bedrock' ? 'bedrock' : 'github';
  user.aiProvider = provider;
  user.aiModel = value('ai_model') || 'gpt-4.1';
  user.integrations = {
    testrail: {
      url: value('tr_url'), project: value('tr_project'), suite: value('tr_suite'),
      email: value('tr_email'), apiKey: value('tr_api_key')
    },
    jira: { url: value('jira_url'), email: value('jira_email'), token: value('jira_token') },
    github: {
      token: value('github_token'), model: value('ai_model') || 'gpt-4.1',
      endpoint: value('ai_endpoint') || 'https://models.github.com/chat/completions',
      maxTokens: value('ai_max_tokens')
    }
  };
  try { saveLocalStore(); }
  catch (e) { return json(res, 500, { error: 'Could not save settings: ' + e.message }); }
  return json(res, 200, { ok: true });
}

function deriveGitHubModelsCatalogUrl(aiEndpoint) {
  const raw = String(aiEndpoint || '').trim();
  if (!raw) return 'https://models.github.ai/catalog/models';

  try {
    const parsed = new URL(raw.startsWith('http') ? raw : 'https://' + raw);
    const host = String(parsed.hostname || '').toLowerCase();
    if (host === 'models.github.ai' || host === 'models.github.com') {
      return 'https://models.github.ai/catalog/models';
    }
  } catch (_) {}

  return null;
}

async function handleGetAiModels(req, res) {
  let body = {};
  if (req.method === 'POST') {
    try { body = await readBody(req); }
    catch(e) { return json(res, 400, { error: e.message }); }
  }

  const userConfig = getPersonalConfigForRequest(req);
  const endpoint = (body.ai_endpoint || userConfig.ai_endpoint || '').trim();
  const apiKey = (body.github_token || userConfig.github_token || '').trim();
  const catalogUrl = deriveGitHubModelsCatalogUrl(endpoint);

  if (!apiKey) {
    return json(res, 400, { error: 'Missing GitHub token. Enter a token first to load supported models.' });
  }

  if (!catalogUrl) {
    return json(res, 400, {
      error: 'Dynamic model loading is currently supported for GitHub Models endpoints only.'
    });
  }

  try {
    const r = await httpsCallRaw(catalogUrl, 'GET', {
      'Accept': 'application/vnd.github+json',
      'Authorization': 'Bearer ' + apiKey,
      'X-GitHub-Api-Version': '2026-03-10',
      'User-Agent': 'Test-Plan-Generator'
    }, null, 30000);

    if (r.status === 401) {
      return json(res, 401, { error: 'Invalid GitHub token for model catalog access.' });
    }
    if (r.status === 403) {
      return json(res, 403, { error: 'GitHub token is not authorized to read the model catalog.' });
    }
    if (r.status !== 200) {
      return json(res, r.status || 500, { error: 'GitHub Models catalog returned ' + r.status });
    }

    const parsed = JSON.parse(r.body || '[]');
    const models = (Array.isArray(parsed) ? parsed : [])
      .filter(model => {
        const input = Array.isArray(model?.supported_input_modalities) ? model.supported_input_modalities : [];
        const output = Array.isArray(model?.supported_output_modalities) ? model.supported_output_modalities : [];
        return input.includes('text') && output.includes('text');
      })
      .map(model => ({
        id: model.id || '',
        name: model.name || model.id || '',
        publisher: model.publisher || '',
        registry: model.registry || '',
        summary: model.summary || '',
        version: model.version || '',
        capabilities: Array.isArray(model.capabilities) ? model.capabilities : [],
        max_input_tokens: Number(model?.limits?.max_input_tokens) || null,
        max_output_tokens: Number(model?.limits?.max_output_tokens) || null,
        rate_limit_tier: model.rate_limit_tier || '',
        supported_input_modalities: Array.isArray(model.supported_input_modalities) ? model.supported_input_modalities : [],
        supported_output_modalities: Array.isArray(model.supported_output_modalities) ? model.supported_output_modalities : [],
        tags: Array.isArray(model.tags) ? model.tags : [],
      }))
      .sort((a, b) => a.name.localeCompare(b.name));

    return json(res, 200, { models, provider: 'github' });
  } catch (e) {
    return json(res, 500, { error: 'Could not load AI models: ' + e.message });
  }
}

// ── httpsCallRaw — for form POST with raw body ────────
function httpsCallRaw(targetUrl, method, extraHeaders, rawBody, timeoutMs = 20000) {
  return new Promise((resolve, reject) => {
    const withProto = targetUrl.startsWith('http') ? targetUrl : 'https://' + targetUrl;
    const m = withProto.match(/^https?:\/\/([^\/]+)(\/.*)?$/);
    if (!m) return reject(new Error('Bad URL: ' + targetUrl));
    const hp       = m[1].split(':');
    const hostname = hp[0];
    const port     = hp[1] ? parseInt(hp[1]) : 443;
    const path     = m[2] || '/';

    const headers = { ...extraHeaders };

    const req = https.request({
      hostname, port, path, method, headers,
      rejectUnauthorized: false
    }, res => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => resolve({ status: res.statusCode, headers: res.headers, body: data }));
    });
    req.on('error', reject);
    req.setTimeout(timeoutMs, () => { req.destroy(); reject(new Error(`Request timed out after ${timeoutMs/1000}s`)); });
    if (rawBody) req.write(rawBody);
    req.end();
  });
}

// ── /testrail-login ───────────────────────────────────
// Logs in with email+password, returns session cookie
async function handleTestRailLogin(req, res) {
  let body;
  try { body = await readBody(req); }
  catch(e) { return json(res, 400, { error: e.message }); }

  const { trUrl, email, password } = body;
  if (!trUrl || !email || !password)
    return json(res, 400, { error: 'Missing: trUrl, email, password' });

  const base     = trUrl.replace(/\/$/, '');
  const loginUrl = base + '/index.php?/auth/login';
  console.log('[Login] Fetching:', loginUrl);

  try {
    // Step 1: GET login page — collect cookies + CSRF token
    const page    = await httpsCallRaw(loginUrl, 'GET', { 'Accept': 'text/html' }, null);
    const cookies = (page.headers['set-cookie'] || []).map(c => c.split(';')[0]).join('; ');
    const csrfM   = page.body.match(/name="_token"[^>]*value="([^"]+)"/);
    const csrf    = csrfM ? csrfM[1] : '';
    console.log('[Login] Page status:', page.status, '| CSRF:', csrf ? 'found' : 'not found');

    // Step 2: POST credentials
    const form = '_token=' + encodeURIComponent(csrf) +
                 '&name='  + encodeURIComponent(email) +
                 '&password=' + encodeURIComponent(password) +
                 '&rememberme=0';

    const loginRes = await httpsCallRaw(loginUrl, 'POST', {
      'Content-Type':   'application/x-www-form-urlencoded',
      'Content-Length': Buffer.byteLength(form).toString(),
      'Cookie':         cookies,
      'Referer':        loginUrl,
      'Accept':         'text/html',
    }, form);

    console.log('[Login] POST status:', loginRes.status);
    const newCookies = (loginRes.headers['set-cookie'] || []).map(c => c.split(';')[0]).join('; ');
    const allCookies = [cookies, newCookies].filter(Boolean).join('; ');
    console.log('[Login] Cookies:', allCookies.substring(0, 80));

    // Step 3: Verify login worked by calling API with cookie
    const verify = await httpsCallRaw(base + '/index.php?/api/v2/get_current_user', 'GET', {
      'Cookie':       allCookies,
      'Accept':       'application/json',
      'Content-Type': 'application/json',
    }, null);

    console.log('[Login] Verify status:', verify.status, '| Body:', verify.body.slice(0, 100));

    if (verify.status === 200 && verify.body.includes('"id"')) {
      const user = JSON.parse(verify.body);
      console.log('[Login] SUCCESS — logged in as:', user.name);
      return json(res, 200, { ok: true, cookie: allCookies, user: user.name, role: user.role });
    }

    return json(res, 401, { error: 'Login failed — wrong email or password' });
  } catch(e) {
    console.error('[Login] Error:', e.message);
    return json(res, 500, { error: 'Login error: ' + e.message });
  }
}

// ── /testrail-proxy ───────────────────────────────────
async function handleTestRailProxy(req, res) {
  let body;
  try { body = await readBody(req); }
  catch(e) { return json(res, 400, { error: e.message }); }

  // Ping
  if (body.ping) return json(res, 200, { ok: true });

  const { url: targetUrl, email, key, cookie, body: payload } = body;

  if (!targetUrl) return json(res, 400, { error: 'Missing: url' });

  // GET for read endpoints, POST for write
  const isRead = [
    '/get_current_user', '/get_project/', '/get_suite/',
    '/get_sections/', '/get_section/', '/get_cases/',
    '/get_priorities', '/get_case_types', '/get_case_fields',
    '/get_milestones', '/get_runs', '/get_plans', '/get_statuses', '/get_users',
  ].some(e => targetUrl.includes(e));
  const method = isRead ? 'GET' : 'POST';

  // Support both cookie auth and Basic auth
  let headers = { 'Accept': 'application/json', 'Content-Type': 'application/json' };
  const userConfig = getPersonalConfigForRequest(req);
  const effectiveEmail = (email || userConfig.tr_email || '').trim();
  const effectiveKey   = (key   || userConfig.tr_api_key || '').trim();
  const auth = trBasicAuth(effectiveEmail, effectiveKey);
  if (!auth) return json(res, 400, { error: 'No TestRail credentials — add Email and API Key in Settings' });
  headers['Authorization'] = auth;
  console.log('[TestRail] Using Basic auth (API key)');

  console.log(`[TestRail] ${method} ${targetUrl}`);
  try {
    const r = await httpsCall(targetUrl, method, headers, isRead ? null : payload);
    console.log('[TestRail] Status:', r.status);
    console.log('[TestRail] Response headers:', JSON.stringify(r.headers));
    console.log('[TestRail] Body:', r.body.slice(0, 300));

    // HTML response = TestRail returned login page = wrong credentials
    const bodyTrimmed = r.body.trim();
    if (bodyTrimmed.startsWith('<!DOCTYPE') || bodyTrimmed.startsWith('<html')) {
      console.log('[TestRail] Got HTML response — credentials rejected, redirected to login');
      return json(res, 401, {
        error: 'TestRail returned a login page — credentials are wrong.\n\n' +
               'You must use an API KEY, not your login password.\n\n' +
               'How to get your API Key:\n' +
               '1. Open https://upland.testrail.com in your browser\n' +
               '2. Click your avatar (top-right corner)\n' +
               '3. Click My Settings\n' +
               '4. Click the API Keys tab\n' +
               '5. Click Add Key, give it any name, click Add Key\n' +
               '6. COPY the generated key\n' +
               '7. Paste it into the API Key field in this tool'
      });
    }

    // Always pass through raw response so client sees exact TestRail message
    console.log('[TestRail] Passing through status:', r.status);
    res.writeHead(r.status, { 'Content-Type': 'application/json' });
    res.end(r.body);
  } catch(e) {
    console.error('[TestRail] Error:', e.message);
    json(res, 500, { error: 'TestRail request failed: ' + e.message });
  }
}

// ── /testrail-attach-case ─────────────────────────────
async function handleTestRailAttachCase(req, res) {
  let body;
  try { body = await readBody(req); }
  catch(e) { return json(res, 400, { error: e.message }); }

  const userConfig = getPersonalConfigForRequest(req);
  const trUrl = (body.tr_url || userConfig.tr_url || '').trim();
  const trEmail = (body.tr_email || userConfig.tr_email || '').trim();
  const trApiKey = (body.tr_api_key || userConfig.tr_api_key || '').trim();
  const caseId = body.case_id;
  const fileName = (body.file_name || 'mockup.png').replace(/[\r\n"]/g, '_');
  const mimeType = body.mime_type || 'image/png';
  const dataUrl = body.data_url || '';

  if (!trUrl || !trEmail || !trApiKey || !caseId || !dataUrl) {
    return json(res, 400, { error: 'Missing required fields for attachment upload' });
  }

  const auth = trBasicAuth(trEmail, trApiKey);
  if (!auth) return json(res, 400, { error: 'Missing TestRail credentials' });

  const m = String(dataUrl).match(/^data:([^;]+);base64,(.+)$/);
  const b64 = m ? m[2] : String(dataUrl);
  let fileBuffer;
  try {
    fileBuffer = Buffer.from(b64, 'base64');
  } catch (e) {
    return json(res, 400, { error: 'Invalid base64 attachment payload' });
  }

  const boundary = '----CopilotBoundary' + Date.now().toString(16);
  const preamble = Buffer.from(
    `--${boundary}\r\n` +
    `Content-Disposition: form-data; name="attachment"; filename="${fileName}"\r\n` +
    `Content-Type: ${mimeType}\r\n\r\n`,
    'utf8'
  );
  const closing = Buffer.from(`\r\n--${boundary}--\r\n`, 'utf8');
  const formBody = Buffer.concat([preamble, fileBuffer, closing]);

  const url = `${trUrl.replace(/\/$/, '')}/index.php?/api/v2/add_attachment_to_case/${caseId}`;
  try {
    const r = await httpsCallRaw(url, 'POST', {
      'Authorization': auth,
      'Accept': 'application/json',
      'Content-Type': `multipart/form-data; boundary=${boundary}`,
      'Content-Length': String(formBody.length),
    }, formBody, 30000);

    if (r.status >= 200 && r.status < 300) {
      let parsed = {};
      try { parsed = JSON.parse(r.body || '{}'); } catch(_) {}
      return json(res, 200, { ok: true, ...parsed });
    }

    let err = 'TestRail returned ' + r.status;
    try {
      const parsed = JSON.parse(r.body || '{}');
      if (parsed.error) err = parsed.error;
    } catch(_) {}
    return json(res, r.status, { error: err });
  } catch (e) {
    return json(res, 500, { error: 'Attachment upload failed: ' + e.message });
  }
}

// ── Static files ────────────────────────────────────────
// ── /download-xlsx — generate Excel from test cases ──────────────────────
async function handleDownloadXlsx(req, res) {
  let body;
  try { body = await readBody(req); }
  catch(e) { return json(res, 400, { error: e.message }); }

  const { secName = 'Section', planName = 'Test Plan', cases = [] } = body;

  // Build Python script dynamically
  const rows = [];
  cases.forEach((tc, i) => {
    const steps = (tc.steps || []).map(s => typeof s === 'string' ? s : (s.step || '')).join('\n');
    rows.push({
      id: 'TC-' + (i + 1),
      title: tc.title || '',
      section: tc.sectionName || secName || '',
      reference: tc.reference || tc.ticketKey || '',
      preconditions: tc.preconditions || '',
      steps,
      expected: tc.expectedResult || '',
    });
  });

  const rowsJson = JSON.stringify(rows).replace(/\\/g, '\\\\').replace(/'/g, "\\'");
  const outPath  = path.join(os.tmpdir(), 'testcases_' + Date.now() + '.xlsx');

  const pyScript = `
import json, openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side, PatternFill
from openpyxl.utils import get_column_letter

rows = json.loads('${rowsJson}')
wb = openpyxl.Workbook()
ws = wb.active
ws.title = "Test Cases"

# Styles
navy  = "0F2557"
blue  = "EFF6FF"
white = "FFFFFF"
green = "F0FDF4"
hdr_font   = Font(name="Arial", bold=True, color=white, size=11)
hdr_fill   = PatternFill("solid", fgColor=navy)
hdr_align  = Alignment(horizontal="center", vertical="center", wrap_text=True)
cell_font  = Font(name="Arial", size=10)
cell_align = Alignment(vertical="top", wrap_text=True)
alt_fill   = PatternFill("solid", fgColor="F8F9FF")
exp_fill   = PatternFill("solid", fgColor=green)
thin = Side(style="thin", color="E3E6EF")
border = Border(left=thin, right=thin, top=thin, bottom=thin)

headers = ["TC #", "Title", "Section", "Reference", "Preconditions", "Test Steps", "Expected Result"]
col_widths = [8, 35, 18, 14, 28, 45, 35]

for ci, (h, w) in enumerate(zip(headers, col_widths), 1):
    cell = ws.cell(row=1, column=ci, value=h)
    cell.font = hdr_font
    cell.fill = hdr_fill
    cell.alignment = hdr_align
    cell.border = border
    ws.column_dimensions[get_column_letter(ci)].width = w

ws.row_dimensions[1].height = 28

for ri, row in enumerate(rows, 2):
    vals = [row["id"], row["title"], row["section"], row["reference"],
            row["preconditions"], row["steps"], row["expected"]]
    fill = alt_fill if ri % 2 == 0 else PatternFill("solid", fgColor=white)
    for ci, val in enumerate(vals, 1):
        cell = ws.cell(row=ri, column=ci, value=val)
        cell.font = cell_font
        cell.alignment = Alignment(vertical="top", wrap_text=True, horizontal="left")
        cell.border = border
        # Expected result column gets green tint
        if ci == 7:
            cell.fill = exp_fill
        else:
            cell.fill = fill
    ws.row_dimensions[ri].height = max(40, min(120, 15 * max(1, len(row["steps"].split("\\\\n")))))

ws.freeze_panes = "A2"
ws.auto_filter.ref = ws.dimensions

wb.save("${outPath}")
`;

  const tmpPy = path.join(os.tmpdir(), 'gen_xlsx_' + Date.now() + '.py');
  require('fs').writeFileSync(tmpPy, pyScript);

  // Try python3 first (Mac/Linux), fall back to python (Windows)
  function runPython(pyExe) {
    return new Promise((resolve, reject) => {
      execFile(pyExe, [tmpPy], { timeout: 30000 }, (err, stdout, stderr) => {
        if (err) reject(new Error(stderr || err.message));
        else resolve();
      });
    });
  }

  try {
    await runPython('python3').catch(() => runPython('python'));
    require('fs').unlink(tmpPy, () => {});

    const buffer = require('fs').readFileSync(outPath);
    require('fs').unlink(outPath, () => {});
    const fname = (planName + '-' + secName).replace(/[^a-z0-9-_ ]/gi, '_') + '.xlsx';
    res.writeHead(200, {
      'Content-Type':        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      'Content-Disposition': 'attachment; filename="' + fname + '"',
      'Content-Length':      buffer.length,
    });
    res.end(buffer);
  } catch(e) {
    console.error('[XLSX]', e.message);
    json(res, 500, { error: 'Excel generation failed: ' + e.message });
  }
}


// ── /download-docx — generate Word document from test plan or test cases ──
async function handleDownloadDocx(req, res) {
  let body;
  try { body = await readBody(req); }
  catch(e) { return json(res, 400, { error: e.message }); }

  const { secName = 'Section', planName = 'Test Plan', isBullet = true, cases = [] } = body;

  if (!docxLib) {
    return json(res, 500, { error: 'docx library not installed. Run: npm install docx' });
  }

  try {
    const { Document, Packer, Paragraph, TextRun, HeadingLevel, LevelFormat,
            AlignmentType, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType } = docxLib;

    const border = { style: BorderStyle.SINGLE, size: 1, color: 'E8ECF1' };
    const borders = { top: border, bottom: border, left: border, right: border };

    const styles = {
      default: { document: { run: { font: 'Arial', size: 24 } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 36, bold: true, font: 'Arial', color: '1E3A8A' },
          paragraph: { spacing: { before: 240, after: 120 }, outlineLevel: 0 } },
        { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 28, bold: true, font: 'Arial', color: '2563EB' },
          paragraph: { spacing: { before: 200, after: 80 }, outlineLevel: 1 } },
        { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 24, bold: true, font: 'Arial', color: '374151' },
          paragraph: { spacing: { before: 160, after: 60 }, outlineLevel: 2 } },
      ]
    };

    const numbering = {
      config: [
        { reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '\u2022',
          alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
        { reference: 'steps', levels: [{ level: 0, format: LevelFormat.DECIMAL, text: '%1.',
          alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      ]
    };

    const children = [
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(planName)] }),
      new Paragraph({ heading: HeadingLevel.HEADING_2, children: [new TextRun(secName)] }),
      new Paragraph({ children: [] }),
    ];

    if (isBullet) {
      // ── Test Plan: bullet points ──────────────────────────────────
      children.push(
        new Paragraph({ children: [new TextRun({ text: 'Test Scenarios', bold: true, font: 'Arial', size: 26 })] })
      );
      children.push(new Paragraph({ children: [] }));
      for (const c of cases) {
        const text = c.bullet || c.title || '';
        children.push(new Paragraph({
          numbering: { reference: 'bullets', level: 0 },
          children: [new TextRun({ text, font: 'Arial', size: 24 })]
        }));
      }
    } else {
      // ── Test Cases: structured cards ──────────────────────────────
      cases.forEach((tc, i) => {
        children.push(
          new Paragraph({ heading: HeadingLevel.HEADING_3,
            children: [new TextRun('TC-' + (i+1) + ': ' + (tc.title || ''))] })
        );

        if (tc.preconditions) {
          children.push(new Paragraph({ children: [
            new TextRun({ text: 'Preconditions: ', bold: true, font: 'Arial', size: 22 }),
            new TextRun({ text: tc.preconditions, font: 'Arial', size: 22, color: '4B5563' }),
          ]}));
        }

        if (tc.steps && tc.steps.length) {
          children.push(new Paragraph({ children: [
            new TextRun({ text: 'Test Steps:', bold: true, font: 'Arial', size: 22 }),
          ]}));
          (tc.steps || []).forEach((s, si) => {
            const stepText = typeof s === 'string' ? s : (s.step || '');
            children.push(new Paragraph({
              numbering: { reference: 'steps', level: 0 },
              children: [new TextRun({ text: stepText, font: 'Arial', size: 22 })]
            }));
          });
        }

        if (tc.expectedResult) {
          children.push(new Paragraph({ children: [
            new TextRun({ text: 'Expected Result: ', bold: true, font: 'Arial', size: 22 }),
            new TextRun({ text: tc.expectedResult, font: 'Arial', size: 22, color: '15803D' }),
          ]}));
        }
        children.push(new Paragraph({ children: [] }));
      });
    }

    const doc = new Document({ numbering, styles, sections: [{
      properties: {
        page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
      },
      children
    }]});

    const buffer = await Packer.toBuffer(doc);
    const fname  = (planName + '-' + secName).replace(/[^a-z0-9-_ ]/gi, '') + '.docx';
    res.writeHead(200, {
      'Content-Type':        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': 'attachment; filename="' + fname + '"',
      'Content-Length':      buffer.length
    });
    res.end(buffer);
  } catch(e) {
    console.error('[DOCX]', e.message);
    json(res, 500, { error: 'DOCX generation failed: ' + e.message });
  }
}

// ── /download-docx-full — full test plan DOCX with all sections ──────
async function handleDownloadDocxFull(req, res) {
  let body;
  try { body = await readBody(req); }
  catch(e) { return json(res, 400, { error: e.message }); }

  const { planName = 'Test Plan', sections = [] } = body;

  if (!docxLib) {
    return json(res, 500, { error: 'docx library not installed. Run: npm install docx' });
  }

  try {
    const { Document, Packer, Paragraph, TextRun, HeadingLevel, LevelFormat,
            AlignmentType, BorderStyle, ShadingType } = docxLib;

    const styles = {
      default: { document: { run: { font: 'Arial', size: 24 } } },
      paragraphStyles: [
        { id: 'Heading1', name: 'Heading 1', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 40, bold: true, font: 'Arial', color: '1E3A8A' },
          paragraph: { spacing: { before: 320, after: 160 }, outlineLevel: 0 } },
        { id: 'Heading2', name: 'Heading 2', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 32, bold: true, font: 'Arial', color: '2563EB' },
          paragraph: { spacing: { before: 240, after: 100 }, outlineLevel: 1 } },
        { id: 'Heading3', name: 'Heading 3', basedOn: 'Normal', next: 'Normal', quickFormat: true,
          run: { size: 26, bold: true, font: 'Arial', color: '374151' },
          paragraph: { spacing: { before: 180, after: 60 }, outlineLevel: 2 } },
      ]
    };

    const numbering = {
      config: [
        { reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '•',
          alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 720, hanging: 360 } } } }] },
      ]
    };

    const children = [
      // Cover title
      new Paragraph({ heading: HeadingLevel.HEADING_1, children: [new TextRun(planName)] }),
      new Paragraph({ children: [] }),
    ];

    let sectionIndex = 0;
    for (const sec of sections) {
      sectionIndex++;
      // Section heading (H2)
      children.push(
        new Paragraph({ heading: HeadingLevel.HEADING_2,
          children: [new TextRun(`${sectionIndex}. ${sec.sectionName || 'Section'}`)] })
      );

      for (const tkt of sec.tickets) {
        // Ticket heading (H3) with key + summary
        const tktLabel = [tkt.key, tkt.summary].filter(Boolean).join(' — ');
        children.push(
          new Paragraph({ heading: HeadingLevel.HEADING_3, children: [new TextRun(tktLabel || 'Ticket')] })
        );

        // Priority label
        if (tkt.priority) {
          children.push(new Paragraph({ children: [
            new TextRun({ text: 'Priority: ', bold: true, font: 'Arial', size: 22, color: '6B7280' }),
            new TextRun({ text: tkt.priority, font: 'Arial', size: 22, color: '374151' }),
          ]}));
        }
        children.push(new Paragraph({ children: [] }));

        // Groups: [{heading, bullets}]
        const groups = tkt.groups || (tkt.bullets ? [{ heading: 'Functional', bullets: tkt.bullets }] : []);
        console.log('[DOCX-FULL] ticket:', tkt.key, 'groups:', groups.map(g=>g.heading+':'+g.bullets.length));
        for (const group of groups) {
          // Category heading — shaded background strip with bold text
          children.push(new Paragraph({
            children: [new TextRun({ text: ' ' })],  // spacer above
            spacing: { before: 120, after: 0 },
          }));
          children.push(new Paragraph({
            children: [
              new TextRun({ text: '  ' + group.heading, bold: true, font: 'Arial', size: 24, color: '1E3A8A' }),
            ],
            shading: { type: ShadingType.CLEAR, color: 'auto', fill: 'DBEAFE' },
            spacing: { before: 0, after: 80 },
            indent: { left: 0 },
          }));

          for (const bullet of (group.bullets || [])) {
            children.push(new Paragraph({
              numbering: { reference: 'bullets', level: 0 },
              children: [new TextRun({ text: bullet, font: 'Arial', size: 24 })]
            }));
          }
          children.push(new Paragraph({ children: [] }));
        }
      }

      // Section divider (blank line)
      children.push(new Paragraph({ children: [] }));
    }

    const doc = new Document({ numbering, styles, sections: [{
      properties: {
        page: { size: { width: 12240, height: 15840 }, margin: { top: 1440, right: 1440, bottom: 1440, left: 1440 } }
      },
      children
    }]});

    const buffer = await Packer.toBuffer(doc);
    const fname  = planName.replace(/[^a-z0-9-_ ]/gi, '') + '-FullTestPlan.docx';
    res.writeHead(200, {
      'Content-Type':        'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
      'Content-Disposition': 'attachment; filename="' + fname + '"',
      'Content-Length':      buffer.length
    });
    res.end(buffer);
  } catch(e) {
    console.error('[DOCX-FULL]', e.message);
    json(res, 500, { error: 'DOCX generation failed: ' + e.message });
  }
}

async function handleBedrockAiProxy(res, prompt, maxTokens, images, selectedModel, currentUser, usageContext) {
  const bedrock = localStore.sharedAi.bedrock || {};
  if (!bedrock.accessKeyId || !bedrock.secretAccessKey) {
    return json(res, 400, { error: 'Amazon Bedrock is not configured by an Admin yet.' });
  }

  const content = [{ text: prompt }];
  for (const image of (Array.isArray(images) ? images : []).slice(0, 5)) {
    const match = String(image?.dataUrl || '').match(/^data:image\/(png|jpeg|jpg|gif|webp);base64,([A-Za-z0-9+/=]+)$/i);
    if (!match) continue;
    const format = match[1].toLowerCase() === 'jpg' ? 'jpeg' : match[1].toLowerCase();
    content.push({ image: { format, source: { bytes: Buffer.from(match[2], 'base64') } } });
  }

  const region = bedrock.region || 'us-east-1';
  const modelId = selectedModel || bedrock.modelId || DEFAULT_BEDROCK_MODEL;
  const inferenceConfig = { maxTokens };
  if (!modelId.includes('claude-sonnet-5-5')) inferenceConfig.temperature = 0.2;
  try {
    const client = new BedrockRuntimeClient({
      region,
      credentials: { accessKeyId: bedrock.accessKeyId, secretAccessKey: bedrock.secretAccessKey }
    });
    const result = await client.send(new ConverseCommand({
      modelId,
      messages: [{ role: 'user', content }],
      inferenceConfig
    }));
    const responseContent = (result.output?.message?.content || []).map(part => part.text || '').join('\n').trim();
    const usage = result.usage || {};
    const usageRecord = recordTokenUsage(currentUser, usage, modelId, usageContext);
    return json(res, 200, {
      content: responseContent,
      model: modelId,
      usageId: usageRecord.id,
      usage: {
        prompt_tokens: usage.inputTokens || 0,
        completion_tokens: usage.outputTokens || 0,
        total_tokens: usage.totalTokens || 0
      }
    });
  } catch (e) {
    console.error(`[Bedrock] Request failed: ${e.name || 'Error'}; model=${modelId}; region=${region}; detail=${e.message || ''}`);
    const message = e.name === 'ResourceNotFoundException'
      ? `Bedrock could not find model "${modelId}" in ${region}. Select an available model or inference profile, and verify model access in that region.`
      : e.name === 'AccessDeniedException'
      ? 'AWS denied this Bedrock request. Check the IAM permissions and model access for the configured account.'
      : e.name === 'ValidationException'
        ? 'Bedrock rejected the request: ' + String(e.message || 'Check the selected model settings.').slice(0, 300)
        : 'Bedrock request failed. Check the shared AWS settings and server log.';
    return json(res, 502, { error: message });
  }
}

// ── /ai-proxy — GitHub Models or Amazon Bedrock ──────────────────────
async function handleAiProxy(req, res) {
  let body;
  try { body = await readBody(req); }
  catch(e) { return json(res, 400, { error: e.message }); }

  const { prompt, maxTokens, images, purpose, storyKey, storyId } = body;
  if (!prompt) return json(res, 400, { error: 'Missing: prompt' });

  const currentUser = getSessionUser(req);
  if (!currentUser) return json(res, 401, { error: 'Not authenticated' });
  const userConfig = getUserAppConfig(currentUser);

  // Hard cap: GitHub Models rejects prompts over ~120k chars (~30k tokens)
  const MAX_PROMPT_CHARS = 40000;
  const trimmedPrompt = prompt.length > MAX_PROMPT_CHARS
    ? prompt.slice(0, MAX_PROMPT_CHARS) + '\n[Prompt truncated to fit model limit]'
    : prompt;

  const configuredTokens = parseInt(userConfig.ai_max_tokens || process.env.GITHUB_MAX_TOKENS || '4000', 10);
  const tokens = Math.max(1, parseInt(maxTokens, 10) || configuredTokens || 4000);
  const usageContext = { purpose, storyKey, storyId };
  if (currentUser.aiProvider === 'bedrock') {
    return handleBedrockAiProxy(res, trimmedPrompt, tokens, images, userConfig.ai_model, currentUser, usageContext);
  }

  const model    = userConfig.ai_model || 'gpt-4.1';
  const endpoint = userConfig.ai_endpoint || 'https://models.github.com/chat/completions';
  const apiKey   = userConfig.github_token || '';

  if (!apiKey)
    return json(res, 400, { error: 'Add your GitHub Models token in Settings, or select Amazon Bedrock.' });

  // ── Request log ──────────────────────────────────────────────────────
  const reqId = Date.now().toString(36).toUpperCase();
  console.log('');
  console.log('┌─────────────────────────────────────────────');
  console.log(`│ [AI REQUEST] ${reqId}`);
  console.log(`│  Provider : GitHub Copilot`);
  console.log(`│  Model    : ${model}`);
  console.log(`│  Endpoint : ${endpoint}`);
  console.log(`│  Tokens   : ${tokens}`);
  console.log(`│  Images   : ${Array.isArray(images) ? images.length : 0}`);
  console.log(`│  Prompt   : ${prompt.slice(0, 120).replace(/\n/g,' ')}...`);
  console.log('└─────────────────────────────────────────────');

  const imageParts = Array.isArray(images)
    ? images
        .filter(img => img && typeof img.dataUrl === 'string' && img.dataUrl.startsWith('data:image/'))
        .slice(0, 5)
        .map(img => ({ type: 'image_url', image_url: { url: img.dataUrl } }))
    : [];

  const userContent = imageParts.length
    ? [{ type: 'text', text: trimmedPrompt }, ...imageParts]
    : trimmedPrompt;

  const reqBody = JSON.stringify({
    model,
    messages:    [{ role: 'user', content: userContent }],
    temperature: 0.2,
    max_tokens:  tokens,
  });

  const startMs = Date.now();

  try {
    // AI responses for exhaustive prompts can take 60-180s — use 300s timeout
    const aiTimeoutMs = parseInt(process.env.AI_TIMEOUT_MS || '300000');
    const r = await httpsCallRaw(endpoint, 'POST', {
      'Content-Type':   'application/json',
      'Authorization':  'Bearer ' + apiKey,
      'Content-Length': Buffer.byteLength(reqBody).toString(),
    }, reqBody, aiTimeoutMs);

    const elapsed = ((Date.now() - startMs) / 1000).toFixed(1);

    // ── Response log ────────────────────────────────────────────────────
    console.log('');
    console.log('┌─────────────────────────────────────────────');
    console.log(`│ [AI RESPONSE] ${reqId}`);
    console.log(`│  Status   : ${r.status}`);
    console.log(`│  Time     : ${elapsed}s`);

    if (r.status !== 200) {
      console.log(`│  Error    : ${r.body.slice(0, 200)}`);
      console.log('└─────────────────────────────────────────────');
      let errMsg = r.body.slice(0, 400);
      if (r.status === 500) {
        errMsg = 'GitHub Models returned a 500 error. This usually means the prompt is too long or the model is temporarily overloaded. Try: (1) reducing your product context, (2) using fewer test plan bullets, or (3) waiting a moment and retrying.';
      } else if (r.status === 429) {
        errMsg = 'Rate limit reached. Wait a minute then try again.';
      } else if (r.status === 401) {
        errMsg = 'Invalid API key. Check your GitHub Copilot token in Settings.';
      }
      return json(res, r.status, { error: errMsg });
    }

    const data         = JSON.parse(r.body);
    const rawContent   = data.choices?.[0]?.message?.content;
    const content      = typeof rawContent === 'string'
      ? rawContent.trim()
      : Array.isArray(rawContent)
        ? rawContent.map(part => {
            if (typeof part === 'string') return part;
            if (part && typeof part === 'object') return part.text || part.content || '';
            return '';
          }).join('\n').trim()
        : (rawContent && typeof rawContent === 'object')
          ? String(rawContent.text || rawContent.content || '').trim()
          : '';
    const usage        = data.usage || {};
    const usageRecord  = recordTokenUsage(currentUser, usage, data.model || model, usageContext);
    console.log(`│  Prompt ↑ : ${usage.prompt_tokens     || '?'} tokens`);
    console.log(`│  Reply  ↓ : ${usage.completion_tokens || '?'} tokens`);
    console.log(`│  Total    : ${usage.total_tokens       || '?'} tokens`);
    console.log(`│  Preview  : ${content.slice(0, 100).replace(/\n/g,' ')}...`);
    console.log('└─────────────────────────────────────────────');

    const model    = data.model || '';
    json(res, 200, { content, model, usage, usageId: usageRecord.id });
  } catch(e) {
    const elapsed = ((Date.now() - startMs) / 1000).toFixed(1);
    console.log('');
    console.log('┌─────────────────────────────────────────────');
    console.log(`│ [AI ERROR] ${reqId}`);
    console.log(`│  Time  : ${elapsed}s`);
    console.log(`│  Error : ${e.message}`);
    console.log('└─────────────────────────────────────────────');
    const isTimeout = e.message && e.message.toLowerCase().includes('timed out');
    json(res, isTimeout ? 504 : 500, {
      error: isTimeout
        ? 'AI request timed out — the prompt may be too large. Try generating one ticket at a time.'
        : 'AI request failed: ' + e.message
    });
  }
}

// ── /jira-get-assignees — get users assignable to a Jira project ─────
async function handleJiraGetAssignees(req, res) {
  let body;
  try { body = await readBody(req); }
  catch(e) { return json(res, 400, { error: e.message }); }

  const { baseUrl, email, token, projectKey } = body;
  if (!baseUrl || !email || !token)
    return json(res, 400, { error: 'Missing: baseUrl, email, token' });

  const project = (projectKey || 'CIMPL').toUpperCase();
  const auth    = Buffer.from(`${email}:${token}`).toString('base64');
  const apiUrl  = `${baseUrl.replace(/\/$/, '')}/rest/api/3/user/assignable/search?project=${project}&maxResults=100`;
  console.log(`[Jira] Fetching assignees for project ${project}...`);
  try {
    const r = await httpsCall(apiUrl, 'GET', { 'Authorization': `Basic ${auth}` }, null);
    if (r.status === 200) {
      const users = JSON.parse(r.body);
      const list = (Array.isArray(users) ? users : []).map(u => ({
        id:    u.accountId,
        name:  u.displayName,
        email: u.emailAddress || '',
      }));
      console.log(`[Jira] Assignees found: ${list.length}`);
      return json(res, 200, list);
    }
    return json(res, r.status, { error: 'Could not fetch users: ' + r.body.slice(0, 200) });
  } catch(e) {
    json(res, 500, { error: 'Jira request failed: ' + e.message });
  }
}

// ── /jira-get-issue-types — get issue types for a project ────────────
async function handleJiraGetIssueTypes(req, res) {
  let body;
  try { body = await readBody(req); } catch(e) { return json(res, 400, { error: e.message }); }
  const { baseUrl, email, token, projectKey } = body;
  if (!baseUrl || !email || !token || !projectKey)
    return json(res, 400, { error: 'Missing required fields' });
  const auth = Buffer.from(`${email}:${token}`).toString('base64');
  const apiUrl = `${baseUrl.replace(/\/$/, '')}/rest/api/3/issuetype/project?projectId=${projectKey}`;
  // Try project-scoped endpoint first, fallback to all issue types
  try {
    // Jira Cloud: use the newer createmeta issuetypes endpoint
    const metaUrl = `${baseUrl.replace(/\/$/, '')}/rest/api/3/issue/createmeta/${projectKey}/issuetypes`;
    const r = await httpsCall(metaUrl, 'GET', {
      'Authorization': `Basic ${auth}`,
      'Content-Type': 'application/json',
    });
    if (r.status === 200) {
      const data = JSON.parse(r.body);
      const types = data.issueTypes || data.values || [];
      if (types.length) {
        return json(res, 200, types.map(t => ({ id: t.id, name: t.name, subtask: t.subtask })));
      }
    }
    // Fallback: all issue types for the instance
    const r2 = await httpsCall(`${baseUrl.replace(/\/$/, '')}/rest/api/3/issuetype`, 'GET', {
      'Authorization': `Basic ${auth}`,
      'Content-Type': 'application/json',
    });
    const types = JSON.parse(r2.body);
    return json(res, 200, Array.isArray(types) ? types.map(t => ({ id: t.id, name: t.name, subtask: t.subtask })) : []);
  } catch(e) {
    json(res, 500, { error: e.message });
  }
}

// ── /jira-get-components — get components for a Jira project ─────────
async function handleJiraGetComponents(req, res) {
  let body;
  try { body = await readBody(req); } catch(e) { return json(res, 400, { error: e.message }); }
  const { baseUrl, email, token, projectKey } = body;
  if (!baseUrl || !email || !token || !projectKey)
    return json(res, 400, { error: 'Missing required fields' });

  const auth = Buffer.from(`${email}:${token}`).toString('base64');
  const apiUrl = `${baseUrl.replace(/\/$/, '')}/rest/api/3/project/${encodeURIComponent(projectKey)}/components`;
  try {
    const r = await httpsCall(apiUrl, 'GET', {
      'Authorization': `Basic ${auth}`,
      'Content-Type': 'application/json',
    });
    if (r.status !== 200) {
      let errMsg = 'Could not fetch components';
      try {
        const errBody = JSON.parse(r.body);
        if (Array.isArray(errBody.errorMessages) && errBody.errorMessages.length) errMsg = errBody.errorMessages.join(', ');
      } catch(_) {}
      return json(res, r.status, { error: errMsg });
    }
    const data = JSON.parse(r.body);
    const list = Array.isArray(data) ? data : [];
    return json(res, 200, list.map(c => ({ id: c.id, name: c.name })));
  } catch(e) {
    return json(res, 500, { error: e.message });
  }
}

// ── /jira-create-subtask — create a QA test plan subtask in Jira ─────
async function handleJiraCreateSubtask(req, res) {
  let body;
  try { body = await readBody(req); }
  catch(e) { return json(res, 400, { error: e.message }); }

  const { baseUrl, email, token, parentKey, title, summary, assigneeId, issueType, labels, componentId } = body;
  if (!baseUrl || !email || !token || !parentKey)
    return json(res, 400, { error: 'Missing required fields: baseUrl, email, token, parentKey' });

  const auth       = Buffer.from(`${email}:${token}`).toString('base64');
  const projectKey = parentKey.split('-')[0];
  const apiUrl     = `${baseUrl.replace(/\/$/, '')}/rest/api/3/issue`;

  // Convert plain-text summary to Atlassian Document Format (ADF)
  // Make section headers bold (e.g. "FUNCTIONAL 4 scenarios").
  const lines = (summary || '').split('\n').filter(l => l.trim());
  const headerLineRe = /^[A-Z][A-Z\s\/]*\s-\s\d+\sscenarios?$/;
  const descContent = lines.length
    ? lines.map(line => {
        const text = line.trim();
        const isHeader = headerLineRe.test(text);
        return {
          type: 'paragraph',
          content: [{
            type: 'text',
            text,
            ...(isHeader ? { marks: [{ type: 'strong' }] } : {}),
          }],
        };
      })
    : [{ type: 'paragraph', content: [{ type: 'text', text: 'Test Plan' }] }];

  const issueFields = {
    project:     { key: projectKey },
    parent:      { key: parentKey },
    issuetype:   { name: issueType || 'Subtask' },
    summary:     title || `${parentKey} Test Plan`,
    labels:      (Array.isArray(labels) && labels.length ? labels : ['test-plan','AI']),
    description: { type: 'doc', version: 1, content: descContent },
    ...(componentId ? { components: [{ id: String(componentId) }] } : {}),
    ...(assigneeId ? { assignee: { id: assigneeId } } : {}),
  };

  console.log(`[Jira] Creating subtask under ${parentKey} — "${issueFields.summary}"`);
  try {
    const r = await httpsCall(apiUrl, 'POST', {
      'Authorization': `Basic ${auth}`,
      'Content-Type':  'application/json',
    }, { fields: issueFields });

    console.log('[Jira] Create subtask status:', r.status, r.body.slice(0, 150));

    if (r.status === 201) {
      const created = JSON.parse(r.body);
      // Best-effort label re-apply: some Jira configs can ignore labels on create for certain types.
      const labelsToApply = Array.isArray(issueFields.labels) ? issueFields.labels : [];
      try {
        await httpsCall(`${baseUrl.replace(/\/$/, '')}/rest/api/3/issue/${created.key}`, 'PUT', {
          'Authorization': `Basic ${auth}`,
          'Content-Type':  'application/json',
        }, {
          update: {
            labels: labelsToApply.map(l => ({ add: l })),
          },
        });
      } catch(labelErr) {
        console.warn('[Jira] Label re-apply warning:', labelErr.message);
      }
      return json(res, 200, { ok: true, key: created.key, id: created.id });
    }

    let errMsg = 'Jira returned ' + r.status;
    try {
      const errBody = JSON.parse(r.body);
      if (errBody.errorMessages?.length)  errMsg = errBody.errorMessages.join(', ');
      else if (errBody.errors)            errMsg = Object.values(errBody.errors).join(', ');
    } catch(_) {}
    return json(res, r.status, { error: errMsg });
  } catch(e) {
    console.error('[Jira] Create subtask error:', e.message);
    json(res, 500, { error: 'Jira request failed: ' + e.message });
  }
}

function serveStatic(req, res) {
  const safePath = req.url.split('?')[0].replace(/\.\./g, '');
  const currentUser = getCurrentUser(req);

  if (safePath === '/' && !currentUser) {
    const loginPath = path.join(PUBLIC_DIR, 'login.html');
    fs.readFile(loginPath, (err, data) => {
      if (err) { res.writeHead(500); return res.end('Login page not found'); }
      res.writeHead(200, {
        'Content-Type': 'text/html; charset=utf-8',
        'Cache-Control': 'no-cache, no-store, must-revalidate',
        'Pragma': 'no-cache',
        'Expires': '0'
      });
      return res.end(data);
    });
    return;
  }

  const filePath = path.join(PUBLIC_DIR, safePath === '/' ? 'index.html' : safePath);

  if (!filePath.startsWith(PUBLIC_DIR)) {
    res.writeHead(403); return res.end('Forbidden');
  }

  fs.readFile(filePath, (err, data) => {
    if (err) {
      res.writeHead(404, { 'Content-Type': 'text/plain' });
      return res.end('Not found: ' + req.url);
    }
    const ext  = path.extname(filePath).toLowerCase();
    const mime = MIME[ext] || 'text/plain';
    const headers = { 'Content-Type': mime };
    // Prevent browser caching HTML so changes are always picked up
    if (ext === '.html' || ext === '') {
      headers['Cache-Control'] = 'no-cache, no-store, must-revalidate';
      headers['Pragma']        = 'no-cache';
      headers['Expires']       = '0';
    }
    res.writeHead(200, headers);
    res.end(data);
  });
}

// ── Server ──────────────────────────────────────────────
const server = http.createServer(async (req, res) => {
  setCORS(req, res);
  if (req.method === 'OPTIONS') { res.writeHead(204); return res.end(); }

  const p = req.url.split('?')[0];
  console.log(`${req.method} ${p}`);

  if (p === '/ping')                                    return handlePing(res);
  if (req.method === 'GET'  && p === '/api/setup/status') return handleSetupStatus(req, res);
  if (req.method === 'POST' && p === '/api/setup')        return handleSetup(req, res);
  if (req.method === 'POST' && p === '/api/login')      return handleLogin(req, res);
  if (req.method === 'POST' && p === '/api/logout')     return handleLogout(req, res);
  if (req.method === 'GET'  && p === '/api/me')         return handleMe(req, res);
  if (req.method === 'GET'  && p === '/api/users')      return handleListUsers(req, res);
  if (req.method === 'POST' && p === '/api/users')      return handleCreateUser(req, res);
  if (req.method === 'GET'  && p === '/api/integrations') return handleGetMyIntegrations(req, res);
  if (req.method === 'POST' && p === '/api/integrations') return handleSaveIntegration(req, res);
  if (req.method === 'GET'  && p === '/api/ai-settings') return handleGetAiSettings(req, res);
  if (req.method === 'POST' && p === '/api/ai-settings') return handleSaveAiSettings(req, res);
  if (req.method === 'GET'  && p === '/api/admin/usage') return handleGetAdminUsage(req, res);

  const authenticatedPaths = new Set([
    '/get-config', '/save-config', '/get-milestones', '/get-sections', '/get-suites', '/get-projects',
    '/test-testrail', '/testrail-login', '/testrail-proxy', '/testrail-attach-case', '/jira-proxy',
    '/jira-attachment-proxy', '/jira-get-assignees', '/jira-get-issue-types', '/jira-get-components',
    '/jira-create-subtask', '/ai-models', '/ai-proxy', '/download-xlsx', '/download-docx',
    '/download-docx-full', '/get-token-history', '/save-token-history', '/api/admin/usage'
  ]);
  if (authenticatedPaths.has(p) && !getCurrentUser(req)) return json(res, 401, { error: 'Not authenticated' });

  if (req.method === 'POST' && p === '/jira-proxy')     return handleJiraProxy(req, res);
  if (req.method === 'POST' && p === '/jira-attachment-proxy') return handleJiraAttachmentProxy(req, res);
  if (req.method === 'POST' && p === '/testrail-login') return handleTestRailLogin(req, res);
  if (req.method === 'POST' && p === '/testrail-proxy') return handleTestRailProxy(req, res);
  if (req.method === 'POST' && p === '/testrail-attach-case') return handleTestRailAttachCase(req, res);
  if (req.method === 'POST' && p === '/save-config')    return handleSaveConfig(req, res);
  if (req.method === 'GET'  && p === '/get-milestones')  return handleGetMilestones(req, res);
  if (req.method === 'GET'  && p === '/get-sections')   return handleGetSections(req, res);
  if (req.method === 'GET'  && p === '/get-suites')     return handleGetSuites(req, res);
  if (req.method === 'POST' && p === '/get-projects')   return handleGetProjects(req, res);
  if (req.method === 'GET'  && p === '/get-config')     return handleGetConfig(req, res);
  if (req.method === 'POST' && p === '/ai-models')      return handleGetAiModels(req, res);
  if (req.method === 'GET'  && p === '/test-testrail')  return handleTestTestrail(req, res);
  if (req.method === 'POST' && p === '/test-testrail')  return handleTestTestrail(req, res);

  if (req.method === 'POST' && p === '/download-xlsx')    return handleDownloadXlsx(req, res);
  if (req.method === 'POST' && p === '/download-docx')      return handleDownloadDocx(req, res);
  if (req.method === 'POST' && p === '/download-docx-full') return handleDownloadDocxFull(req, res);
  if (req.method === 'POST' && p === '/ai-proxy')        return handleAiProxy(req, res);

  // ── Jira — get assignable users ──────────────────────
  if (req.method === 'POST' && p === '/jira-get-assignees') return handleJiraGetAssignees(req, res);
  // ── Jira — get issue types ───────────────────────────
  if (req.method === 'POST' && p === '/jira-get-issue-types') return handleJiraGetIssueTypes(req, res);
  // ── Jira — get components ────────────────────────────
  if (req.method === 'POST' && p === '/jira-get-components') return handleJiraGetComponents(req, res);
  // ── Jira — create subtask ────────────────────────────
  if (req.method === 'POST' && p === '/jira-create-subtask') return handleJiraCreateSubtask(req, res);

  if (req.method === 'GET'  && p === '/get-token-history')   return handleGetTokenHistory(req, res);
  if (req.method === 'POST' && p === '/save-token-history')  return handleSaveTokenHistory(req, res);

  serveStatic(req, res);
});

server.listen(PORT, HOST, () => {
  loadConfig();
  loadLocalStore();

  console.log('');
  console.log('  ╔══════════════════════════════════════════╗');
  console.log('  ║    Test Plan Generator — Ready ✅         ║');
  console.log('  ╠══════════════════════════════════════════╣');
  console.log(`  ║  Open  : http://localhost:${PORT}           ║`);
  console.log(`  ║  Ping  : http://localhost:${PORT}/ping      ║`);
  console.log('  ╠══════════════════════════════════════════╣');
  console.log('  ║  Ctrl+C to stop                          ║');
  console.log('  ╠══════════════════════════════════════════╣');

  console.log(`  ║  Accounts : ${localStore.users.length ? localStore.users.length + ' local user(s)' : 'First-run Admin setup required'}`.padEnd(53) + '║');
  const bedrockConfigured = !!(localStore.sharedAi.bedrock.accessKeyId && localStore.sharedAi.bedrock.secretAccessKey);
  console.log(`  ║  AI       : ${bedrockConfigured ? 'Shared Bedrock credentials configured' : 'Configure GitHub per user or shared Bedrock'}`.padEnd(53) + '║');
  console.log('  ╚══════════════════════════════════════════╝');
  console.log('');
});

server.on('error', err => {
  if (err.code === 'EADDRINUSE') {
    console.error(`\n  ❌ Port ${PORT} in use — stop other server or change PORT\n`);
  } else {
    console.error('  ❌ Error:', err.message);
  }
  process.exit(1);
});

process.on('SIGINT', () => {
  console.log('\n  👋 Stopped.\n');
  process.exit(0);
});
