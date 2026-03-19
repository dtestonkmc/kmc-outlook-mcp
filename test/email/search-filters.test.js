const nock = require('nock');

// Stub auth before requiring search module
jest.mock('../../auth', () => ({
  ensureAuthenticated: jest.fn().mockResolvedValue('test-token')
}));

jest.mock('../../email/folder-utils', () => ({
  resolveFolderPath: jest.fn().mockResolvedValue('/me/mailFolders/inbox/messages')
}));

jest.mock('../../utils/graph-api', () => ({
  callGraphAPI: jest.fn(),
  callGraphAPIPaginated: jest.fn()
}));

const { callGraphAPIPaginated } = require('../../utils/graph-api');
const handleSearchEmails = require('../../email/search');

beforeEach(() => {
  callGraphAPIPaginated.mockReset();
  callGraphAPIPaginated.mockResolvedValue({ value: [] });
});

afterEach(() => nock.cleanAll());

test('passes receivedDateTimeBefore as OData $filter', async () => {
  const cutoff = '2026-03-19T04:00:00.000Z';
  await handleSearchEmails({ receivedDateTimeBefore: cutoff, count: 10 });

  const [, , , params] = callGraphAPIPaginated.mock.calls[0];
  expect(params.$filter).toContain(`receivedDateTime lt ${cutoff}`);
});

test('passes receivedDateTimeAfter as OData $filter', async () => {
  const since = '2026-03-18T20:00:00.000Z';
  await handleSearchEmails({ receivedDateTimeAfter: since, count: 10 });

  const [, , , params] = callGraphAPIPaginated.mock.calls[0];
  expect(params.$filter).toContain(`receivedDateTime gt ${since}`);
});

test('combines receivedDateTime before and after in single $filter', async () => {
  const before = '2026-03-19T04:00:00.000Z';
  const after = '2026-03-12T04:00:00.000Z';
  await handleSearchEmails({ receivedDateTimeBefore: before, receivedDateTimeAfter: after, count: 10 });

  const [, , , params] = callGraphAPIPaginated.mock.calls[0];
  expect(params.$filter).toContain(`receivedDateTime lt ${before}`);
  expect(params.$filter).toContain(`receivedDateTime gt ${after}`);
});

test('passes lastModifiedAfter as OData $filter', async () => {
  const since = '2026-03-19T07:30:00.000Z';
  await handleSearchEmails({ lastModifiedAfter: since, count: 10 });

  const [, , , params] = callGraphAPIPaginated.mock.calls[0];
  expect(params.$filter).toContain(`lastModifiedDateTime ge ${since}`);
});

test('combines date filters with isRead filter', async () => {
  const cutoff = '2026-03-19T04:00:00.000Z';
  await handleSearchEmails({ receivedDateTimeBefore: cutoff, unreadOnly: true, count: 10 });

  const [, , , params] = callGraphAPIPaginated.mock.calls[0];
  expect(params.$filter).toContain(`receivedDateTime lt ${cutoff}`);
  expect(params.$filter).toContain('isRead eq false');
});

test('does not fall back to recent emails when date filter is specified but returns 0 results', async () => {
  const cutoff = '2026-03-01T00:00:00.000Z';
  callGraphAPIPaginated.mockResolvedValue({ value: [] });

  const result = await handleSearchEmails({ receivedDateTimeBefore: cutoff, count: 10 });

  // Should only call the API once — no progressive fallback to recent emails
  expect(callGraphAPIPaginated).toHaveBeenCalledTimes(1);
  expect(result.content[0].text).toContain('No emails found');
});

test('does not mix in recent emails when a paginated date-filtered query returns empty on later pages', async () => {
  const cutoff = '2026-03-01T00:00:00.000Z';
  callGraphAPIPaginated
    .mockResolvedValueOnce({ value: [{ id: 'msg1', subject: 'Old email', from: { emailAddress: { name: 'A', address: 'a@b.com' } }, receivedDateTime: '2026-02-01T00:00:00Z', isRead: true }] })
    .mockResolvedValueOnce({ value: [] });

  const result = await handleSearchEmails({ receivedDateTimeBefore: cutoff, count: 50 });

  expect(result.content[0].text).toContain('Old email');
  const allCallParams = callGraphAPIPaginated.mock.calls.map(([, , , params]) => params);
  allCallParams.forEach(params => {
    expect(params.$filter).toBeDefined();
    expect(params.$filter).toContain('receivedDateTime lt');
  });
});
