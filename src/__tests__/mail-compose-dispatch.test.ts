import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import GraphClient from '../graph-client.js';
import { fetchWithResilience } from '../lib/graph-resilience.js';
import type AuthManager from '../auth.js';

vi.mock('../logger.js', () => ({
  default: { info: vi.fn(), warn: vi.fn(), error: vi.fn(), debug: vi.fn() },
}));
vi.mock('../lib/graph-resilience.js', () => ({
  fetchWithResilience: vi.fn(),
  getSharedBreaker: vi.fn(() => ({})),
  loadResilienceConfig: vi.fn(() => ({})),
}));
const key = 'MS365_MCP_REQUIRE_BOUND_MAIL_COMPOSE';
const old = process.env[key];
const fetchMock = vi.mocked(fetchWithResilience);
function client() {
  return new GraphClient(
    { getToken: vi.fn().mockResolvedValue('synthetic-token') } as unknown as AuthManager,
    { clientId: 'synthetic', tenantId: 'synthetic', cloudType: 'global' }
  );
}
beforeEach(() => {
  process.env[key] = '1';
  vi.clearAllMocks();
});
afterEach(() => {
  if (old === undefined) delete process.env[key];
  else process.env[key] = old;
});

describe('actual GraphClient outbound choke point', () => {
  it('blocks free native reply before any fetch', async () => {
    await expect(
      client().makeRequest('/me/messages/id/createReply', {
        method: 'POST',
        body: JSON.stringify({ Comment: 'unsigned business' }),
      })
    ).rejects.toThrow('BOUND_MAIL_COMPOSE_REQUIRED');
    expect(fetchMock).not.toHaveBeenCalled();
  });
  it('blocks equivalent batch through public MCP response path', async () => {
    const result = await client().graphRequest('/$batch', {
      method: 'POST',
      body: JSON.stringify({
        requests: [
          {
            id: '1',
            method: 'PATCH',
            url: '/me/messages/id',
            body: { body: { contentType: 'text', content: 'unsigned' } },
          },
        ],
      }),
    });
    expect(result.isError).toBe(true);
    expect(fetchMock).not.toHaveBeenCalled();
  });
  it('rejects an entire mixed batch before a safe subrequest can dispatch', async () => {
    await expect(
      client().makeRequest('/$batch', {
        method: 'POST',
        body: JSON.stringify({
          requests: [
            { id: '1', method: 'GET', url: '/me/messages/id' },
            { id: '2', method: 'POST', url: '/users/pia/messages/id/createReplyAll' },
          ],
        }),
      })
    ).rejects.toThrow();
    expect(fetchMock).not.toHaveBeenCalled();
  });
  it('permits a normal read with the same policy enabled', async () => {
    fetchMock.mockResolvedValue(
      new Response(JSON.stringify({ id: 'synthetic-read' }), {
        status: 200,
        headers: { 'content-type': 'application/json' },
      })
    );
    await client().makeRequest('/me/messages/id');
    expect(fetchMock).toHaveBeenCalledTimes(1);
  });
});
