jest.mock('../../auth', () => ({
  ensureAuthenticated: jest.fn().mockResolvedValue('test-token')
}));

jest.mock('../../utils/graph-api', () => ({
  callGraphAPI: jest.fn()
}));

const { callGraphAPI } = require('../../utils/graph-api');

// Set env before requiring the module
process.env.TEAMS_SELF_CHAT_ID = 'test-chat-id';
const { handleSendTeamsMessage } = require('../../teams/send-teams-message');

beforeEach(() => callGraphAPI.mockReset());

test('sends message to configured self-chat', async () => {
  callGraphAPI.mockResolvedValue({ id: 'msg-1' });

  const result = await handleSendTeamsMessage({ message: 'Hello Teams' });

  expect(callGraphAPI).toHaveBeenCalledWith(
    'test-token',
    'POST',
    '/me/chats/test-chat-id/messages',
    { body: { content: 'Hello Teams', contentType: 'text' } }
  );
  expect(result.success).toBe(true);
});

test('returns failure when TEAMS_SELF_CHAT_ID is not set', async () => {
  const savedChatId = process.env.TEAMS_SELF_CHAT_ID;
  delete process.env.TEAMS_SELF_CHAT_ID;
  // Re-require module after env change by using jest.resetModules
  jest.resetModules();
  const { handleSendTeamsMessage: fn } = require('../../teams/send-teams-message');
  const result = await fn({ message: 'test' });
  expect(result.success).toBe(false);
  expect(result.error).toContain('TEAMS_SELF_CHAT_ID');
  process.env.TEAMS_SELF_CHAT_ID = savedChatId;
});

test('returns failure when Graph API call fails', async () => {
  callGraphAPI.mockRejectedValue(new Error('500 Server Error'));
  const result = await handleSendTeamsMessage({ message: 'test' });
  expect(result.success).toBe(false);
  expect(result.error).toContain('500 Server Error');
});

test('returns error when message is missing', async () => {
  const result = await handleSendTeamsMessage({});
  expect(result.success).toBe(false);
  expect(result.error).toContain('message is required');
});
