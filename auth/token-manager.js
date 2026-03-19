/**
 * Token management for Microsoft Graph API authentication
 */
const fs = require('fs');
const https = require('https');
const querystring = require('querystring');
const config = require('../config');

// Global variable to store tokens
let cachedTokens = null;
let refreshInFlight = null;

/**
 * Loads authentication tokens from the token file
 * @returns {object|null} - The loaded tokens or null if not available
 */
function loadTokenCache() {
  try {
    const tokenPath = config.AUTH_CONFIG.tokenStorePath;
    console.error(`[DEBUG] Attempting to load tokens from: ${tokenPath}`);
    console.error(`[DEBUG] HOME directory: ${process.env.HOME}`);
    console.error(`[DEBUG] Full resolved path: ${tokenPath}`);
    
    // Log file existence and details
    if (!fs.existsSync(tokenPath)) {
      console.error('[DEBUG] Token file does not exist');
      return null;
    }
    
    const stats = fs.statSync(tokenPath);
    console.error(`[DEBUG] Token file stats:
      Size: ${stats.size} bytes
      Created: ${stats.birthtime}
      Modified: ${stats.mtime}`);
    
    const tokenData = fs.readFileSync(tokenPath, 'utf8');
    console.error('[DEBUG] Token file contents length:', tokenData.length);
    console.error('[DEBUG] Token file first 200 characters:', tokenData.slice(0, 200));
    
    try {
      const tokens = JSON.parse(tokenData);
      console.error('[DEBUG] Parsed tokens keys:', Object.keys(tokens));
      
      // Log each key's value to see what's present
      Object.keys(tokens).forEach(key => {
        console.error(`[DEBUG] ${key}: ${typeof tokens[key]}`);
      });
      
      // Check for access token presence
      if (!tokens.access_token) {
        console.error('[DEBUG] No access_token found in tokens');
        return null;
      }
      
      // Check token expiration
      const now = Date.now();
      const expiresAt = tokens.expires_at || 0;
      
      console.error(`[DEBUG] Current time: ${now}`);
      console.error(`[DEBUG] Token expires at: ${expiresAt}`);
      
      if (now > expiresAt) {
        console.error('[DEBUG] Token has expired');
        return null;
      }
      
      // Update the cache
      cachedTokens = tokens;
      return tokens;
    } catch (parseError) {
      console.error('[DEBUG] Error parsing token JSON:', parseError);
      return null;
    }
  } catch (error) {
    console.error('[DEBUG] Error loading token cache:', error);
    return null;
  }
}

/**
 * Saves authentication tokens to the token file
 * @param {object} tokens - The tokens to save
 * @returns {boolean} - Whether the save was successful
 */
function saveTokenCache(tokens) {
  try {
    const tokenPath = config.AUTH_CONFIG.tokenStorePath;
    console.error(`Saving tokens to: ${tokenPath}`);
    
    fs.writeFileSync(tokenPath, JSON.stringify(tokens, null, 2));
    console.error('Tokens saved successfully');
    
    // Update the cache
    cachedTokens = tokens;
    return true;
  } catch (error) {
    console.error('Error saving token cache:', error);
    return false;
  }
}

/**
 * Attempts to refresh the access token using the stored refresh_token.
 * @returns {Promise<string|null>} - New access token or null if refresh fails
 */
async function refreshAccessToken() {
  // Read raw token file — bypass expiry check in loadTokenCache so we can still
  // get the refresh_token even when the access_token has already expired.
  let tokens = null;
  try {
    const tokenPath = config.AUTH_CONFIG.tokenStorePath;
    if (fs.existsSync(tokenPath)) {
      tokens = JSON.parse(fs.readFileSync(tokenPath, 'utf8'));
    }
  } catch (e) {
    console.error('[token-manager] Failed to read token file for refresh:', e);
  }

  if (!tokens || !tokens.refresh_token) {
    console.error('[token-manager] No refresh token available');
    return null;
  }

  const tenantId = process.env.OUTLOOK_TENANT_ID || 'common';
  const clientId = process.env.OUTLOOK_CLIENT_ID || '';
  const clientSecret = process.env.OUTLOOK_CLIENT_SECRET || '';

  const body = querystring.stringify({
    grant_type: 'refresh_token',
    refresh_token: tokens.refresh_token,
    client_id: clientId,
    client_secret: clientSecret,
    scope: 'https://graph.microsoft.com/.default offline_access'
  });

  return new Promise((resolve) => {
    const options = {
      hostname: 'login.microsoftonline.com',
      path: `/${tenantId}/oauth2/v2.0/token`,
      method: 'POST',
      headers: {
        'Content-Type': 'application/x-www-form-urlencoded',
        'Content-Length': Buffer.byteLength(body)
      }
    };

    const req = https.request(options, (res) => {
      let data = '';
      res.on('data', chunk => data += chunk);
      res.on('end', () => {
        if (res.statusCode !== 200) {
          console.error(`[token-manager] Token refresh failed: ${res.statusCode} ${data}`);
          resolve(null);
          return;
        }
        try {
          const parsed = JSON.parse(data);
          const newTokens = {
            ...tokens,
            access_token: parsed.access_token,
            refresh_token: parsed.refresh_token || tokens.refresh_token,
            expires_at: Date.now() + ((parsed.expires_in || 3600) * 1000)
          };
          saveTokenCache(newTokens);
          console.error('[token-manager] Token refreshed successfully');
          resolve(parsed.access_token);
        } catch (e) {
          console.error('[token-manager] Failed to parse refresh response:', e);
          resolve(null);
        }
      });
    });

    req.on('error', (e) => {
      console.error('[token-manager] Refresh request error:', e);
      resolve(null);
    });

    req.write(body);
    req.end();
  });
}

/**
 * Gets the current Graph API access token, loading from cache if necessary.
 * Automatically refreshes if expired or expiring within 5 minutes.
 * @returns {Promise<string|null>} - The access token or null if not available
 */
async function getAccessToken() {
  const REFRESH_THRESHOLD_MS = 5 * 60 * 1000; // 5 minutes

  if (cachedTokens && cachedTokens.access_token) {
    const expiresAt = cachedTokens.expires_at || 0;
    if (Date.now() < expiresAt - REFRESH_THRESHOLD_MS) {
      return cachedTokens.access_token;
    }
  }

  // Read raw token file to check expiry (loadTokenCache returns null on expiry,
  // so we read directly to decide whether to refresh or just return the token).
  let rawTokens = null;
  try {
    const tokenPath = config.AUTH_CONFIG.tokenStorePath;
    if (fs.existsSync(tokenPath)) {
      rawTokens = JSON.parse(fs.readFileSync(tokenPath, 'utf8'));
    }
  } catch (e) {
    // ignore parse errors
  }

  if (!rawTokens || !rawTokens.access_token) return null;

  const expiresAt = rawTokens.expires_at || 0;
  if (Date.now() < expiresAt - REFRESH_THRESHOLD_MS) {
    // Valid and not expiring soon — update cache and return
    cachedTokens = rawTokens;
    return rawTokens.access_token;
  }

  // Token expired or expiring soon — try refresh
  console.error('[token-manager] Token expired or expiring soon, attempting refresh');
  if (!refreshInFlight) {
    refreshInFlight = refreshAccessToken().finally(() => {
      refreshInFlight = null;
    });
  }
  return await refreshInFlight;
}

/**
 * Gets the current Flow API access token
 * @returns {string|null} - The Flow access token or null if not available
 */
function getFlowAccessToken() {
  const tokens = loadTokenCache();
  if (!tokens) return null;

  // Check if flow token exists and is not expired
  if (tokens.flow_access_token && tokens.flow_expires_at) {
    if (Date.now() < tokens.flow_expires_at) {
      return tokens.flow_access_token;
    }
  }

  return null;
}

/**
 * Saves Flow API tokens alongside existing Graph tokens
 * @param {object} flowTokens - The Flow tokens to save
 * @returns {boolean} - Whether the save was successful
 */
function saveFlowTokens(flowTokens) {
  try {
    const tokenPath = config.AUTH_CONFIG.tokenStorePath;

    // Load existing tokens
    let existingTokens = {};
    if (fs.existsSync(tokenPath)) {
      const tokenData = fs.readFileSync(tokenPath, 'utf8');
      existingTokens = JSON.parse(tokenData);
    }

    // Merge flow tokens
    const mergedTokens = {
      ...existingTokens,
      flow_access_token: flowTokens.access_token,
      flow_refresh_token: flowTokens.refresh_token,
      flow_expires_at: flowTokens.expires_at || (Date.now() + (flowTokens.expires_in || 3600) * 1000)
    };

    fs.writeFileSync(tokenPath, JSON.stringify(mergedTokens, null, 2));
    console.log('Flow tokens saved successfully');

    // Update cache
    cachedTokens = mergedTokens;
    return true;
  } catch (error) {
    console.error('Error saving Flow tokens:', error);
    return false;
  }
}

/**
 * Creates a test access token for use in test mode
 * @returns {object} - The test tokens
 */
function createTestTokens() {
  const testTokens = {
    access_token: "test_access_token_" + Date.now(),
    refresh_token: "test_refresh_token_" + Date.now(),
    expires_at: Date.now() + (3600 * 1000) // 1 hour
  };
  
  saveTokenCache(testTokens);
  return testTokens;
}

/** FOR TESTING ONLY — resets in-memory token cache */
function _resetCacheForTesting() {
  cachedTokens = null;
  refreshInFlight = null;
}

module.exports = {
  loadTokenCache,
  saveTokenCache,
  getAccessToken,      // now async
  getFlowAccessToken,
  saveFlowTokens,
  createTestTokens,
  refreshAccessToken,  // new
  _resetCacheForTesting   // for tests only
};
