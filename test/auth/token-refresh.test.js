const nock = require('nock');
const path = require('path');
const os = require('os');
const fs = require('fs');

const TOKEN_PATH = require('path').join(os.tmpdir(), '.outlook-mcp-tokens.json');

jest.mock('../../config', () => ({
  AUTH_CONFIG: {
    clientId: '',
    clientSecret: '',
    tenantId: 'common',
    tokenStorePath: require('path').join(require('os').tmpdir(), '.outlook-mcp-tokens.json'),
    authServerUrl: 'http://localhost:3333'
  }
}));

const { loadTokenCache, saveTokenCache, getAccessToken, refreshAccessToken, _resetCacheForTesting } = require('../../auth/token-manager');

beforeEach(() => {
  _resetCacheForTesting();
  if (fs.existsSync(TOKEN_PATH)) fs.unlinkSync(TOKEN_PATH);
});

afterEach(() => {
  nock.cleanAll();
  if (fs.existsSync(TOKEN_PATH)) fs.unlinkSync(TOKEN_PATH);
});

test('getAccessToken returns null when no token file exists', async () => {
  const token = await getAccessToken();
  expect(token).toBeNull();
});

test('getAccessToken returns token when not expired', async () => {
  saveTokenCache({
    access_token: 'valid-token',
    refresh_token: 'refresh-token',
    expires_at: Date.now() + 3600000
  });
  const token = await getAccessToken();
  expect(token).toBe('valid-token');
});

test('getAccessToken triggers refresh when token is expired', async () => {
  saveTokenCache({
    access_token: 'expired-token',
    refresh_token: 'my-refresh-token',
    expires_at: Date.now() - 1000
  });

  nock('https://login.microsoftonline.com')
    .post('/common/oauth2/v2.0/token')
    .reply(200, {
      access_token: 'new-access-token',
      refresh_token: 'new-refresh-token',
      expires_in: 3600
    });

  process.env.OUTLOOK_CLIENT_ID = 'test-client-id';
  process.env.OUTLOOK_CLIENT_SECRET = 'test-client-secret';
  process.env.OUTLOOK_TENANT_ID = 'common';

  const token = await getAccessToken();
  expect(token).toBe('new-access-token');

  const saved = JSON.parse(fs.readFileSync(TOKEN_PATH, 'utf8'));
  expect(saved.access_token).toBe('new-access-token');
  expect(saved.refresh_token).toBe('new-refresh-token');
});

test('getAccessToken triggers refresh when token expires within 5 minutes', async () => {
  saveTokenCache({
    access_token: 'expiring-soon-token',
    refresh_token: 'my-refresh-token',
    expires_at: Date.now() + (4 * 60 * 1000) // 4 min — inside 5 min window
  });

  nock('https://login.microsoftonline.com')
    .post('/common/oauth2/v2.0/token')
    .reply(200, {
      access_token: 'refreshed-token',
      refresh_token: 'new-refresh-token',
      expires_in: 3600
    });

  process.env.OUTLOOK_CLIENT_ID = 'test-client-id';
  process.env.OUTLOOK_CLIENT_SECRET = 'test-client-secret';
  process.env.OUTLOOK_TENANT_ID = 'common';

  const token = await getAccessToken();
  expect(token).toBe('refreshed-token');
});

test('getAccessToken returns null when refresh fails', async () => {
  saveTokenCache({
    access_token: 'expired-token',
    refresh_token: 'bad-refresh-token',
    expires_at: Date.now() - 1000
  });

  nock('https://login.microsoftonline.com')
    .post('/common/oauth2/v2.0/token')
    .reply(400, { error: 'invalid_grant' });

  const token = await getAccessToken();
  expect(token).toBeNull();
});
