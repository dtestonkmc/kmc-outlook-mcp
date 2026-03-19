jest.mock('../../auth', () => ({
  ensureAuthenticated: jest.fn().mockResolvedValue('test-token')
}));

jest.mock('../../utils/graph-api', () => ({
  callGraphAPI: jest.fn()
}));

const { callGraphAPI } = require('../../utils/graph-api');
const { handleMoveToTrash } = require('../../email/move-to-trash');

beforeEach(() => callGraphAPI.mockReset());

test('calls Graph API move endpoint with deleteditems destination', async () => {
  callGraphAPI.mockResolvedValue({ id: 'moved-message-id' });

  const result = await handleMoveToTrash({ messageId: 'abc123' });

  expect(callGraphAPI).toHaveBeenCalledWith(
    'test-token',
    'POST',
    '/me/messages/abc123/move',
    null,
    { destinationId: 'deleteditems' }
  );
  expect(result.content[0].text).toContain('moved to Deleted Items');
});

test('returns error when messageId is missing', async () => {
  const result = await handleMoveToTrash({});
  expect(result.content[0].text).toContain('messageId is required');
  expect(callGraphAPI).not.toHaveBeenCalled();
});

test('returns error message when Graph API fails', async () => {
  callGraphAPI.mockRejectedValue(new Error('403 Forbidden'));
  const result = await handleMoveToTrash({ messageId: 'abc123' });
  expect(result.content[0].text).toContain('Error');
});
