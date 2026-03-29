/**
 * 共通: OAuth トークン取得 + HTTP リクエスト
 */
const { execSync } = require('child_process');
const https = require('https');
const fs = require('fs');
const path = require('path');
const os = require('os');

function getClaspCredentials() {
  const paths = [
    path.join(process.cwd(), '.clasprc.json'),
    path.join(os.homedir(), '.clasprc.json')
  ];
  for (const p of paths) {
    if (fs.existsSync(p)) {
      const data = JSON.parse(fs.readFileSync(p, 'utf8'));
      const creds = data.tokens?.default || data.tokens?.organization;
      if (creds?.refresh_token) return creds;
    }
  }
  throw new Error('No clasp credentials found. Run: npx clasp login');
}

function getAccessToken() {
  const creds = getClaspCredentials();
  const params = new URLSearchParams({
    client_id: creds.client_id,
    client_secret: creds.client_secret,
    refresh_token: creds.refresh_token,
    grant_type: 'refresh_token'
  });
  const result = execSync(
    `curl -s -X POST "https://oauth2.googleapis.com/token" -d "${params.toString()}"`,
    { encoding: 'utf8' }
  );
  let parsed;
  try { parsed = JSON.parse(result); } catch (e) {
    throw new Error(`Token refresh failed: invalid response`);
  }
  if (!parsed.access_token) {
    throw new Error(`Token refresh failed: ${parsed.error_description || parsed.error || 'no token'}`);
  }
  return parsed.access_token;
}

/**
 * HTTPS POST with JSON body, returns parsed JSON
 */
function postJSON(url, body, token) {
  return new Promise((resolve, reject) => {
    const data = typeof body === 'string' ? body : JSON.stringify(body);
    const parsed = new URL(url);
    const options = {
      hostname: parsed.hostname,
      path: parsed.pathname + parsed.search,
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'Content-Length': Buffer.byteLength(data),
        ...(token ? { 'Authorization': `Bearer ${token}` } : {})
      }
    };

    const req = https.request(options, (res) => {
      let body = '';
      res.on('data', chunk => { body += chunk; });
      res.on('end', () => {
        // Follow redirects for GAS (302 → GET with same path)
        if (res.statusCode >= 300 && res.statusCode < 400 && res.headers.location) {
          postJSON(res.headers.location, data, token).then(resolve).catch(reject);
          return;
        }
        try {
          resolve(JSON.parse(body));
        } catch (e) {
          reject(new Error(`Invalid JSON response (HTTP ${res.statusCode}): ${body.substring(0, 200)}`));
        }
      });
    });
    req.on('error', reject);
    req.setTimeout(30000, () => { req.destroy(); reject(new Error('Request timeout')); });
    req.write(data);
    req.end();
  });
}

/**
 * Sync wrapper using python urllib (handles POST redirects correctly)
 * curl -L converts POST→GET on redirect, losing the body.
 * Python urllib preserves POST method across redirects.
 */
function postJSONSync(url, body, token) {
  const payloadFile = path.join(os.tmpdir(), `gas-req-${Date.now()}.json`);
  const tokenFile = path.join(os.tmpdir(), `gas-tok-${Date.now()}.txt`);
  const scriptFile = path.join(os.tmpdir(), `gas-req-${Date.now()}.py`);
  const data = typeof body === 'string' ? body : JSON.stringify(body);

  fs.writeFileSync(payloadFile, data);
  fs.writeFileSync(tokenFile, token || '');
  fs.writeFileSync(scriptFile, [
    'import json, urllib.request',
    `payload = open(${JSON.stringify(payloadFile)}, "rb").read()`,
    `token = open(${JSON.stringify(tokenFile)}).read().strip()`,
    'headers = {"Content-Type": "application/json"}',
    'if token: headers["Authorization"] = "Bearer " + token',
    `req = urllib.request.Request(${JSON.stringify(url)}, data=payload, headers=headers)`,
    'try:',
    '    resp = urllib.request.urlopen(req, timeout=30)',
    '    print(resp.read().decode())',
    'except urllib.error.HTTPError as e:',
    '    print(json.dumps({"success":False,"error":"HTTP "+str(e.code),"body":e.read().decode()[:200]}))',
  ].join('\n'));

  try {
    const result = execSync(`python3 "${scriptFile}"`, { encoding: 'utf8', timeout: 30000 }).trim();
    try { return JSON.parse(result); } catch (e) {
      throw new Error(`Invalid JSON response: ${result.substring(0, 200)}`);
    }
  } finally {
    for (const f of [payloadFile, tokenFile, scriptFile]) {
      try { fs.unlinkSync(f); } catch (_) {}
    }
  }
}

function getConfig() {
  const configPath = path.join(__dirname, '..', 'config.json');
  if (!fs.existsSync(configPath)) {
    throw new Error('scripts/config.json not found. Copy from config.json.template and fill in values.');
  }
  return JSON.parse(fs.readFileSync(configPath, 'utf8'));
}

module.exports = { getAccessToken, postJSON, postJSONSync, getConfig };
