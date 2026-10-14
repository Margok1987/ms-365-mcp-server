import { describe, expect, it, vi } from 'vitest';
import GraphClient from '../src/graph-client.js';
import logger from '../src/logger.js';

describe('GraphClient request logging', () => {
  it('omits attachment bodies and bearer tokens from Graph logs', async () => {
    const marker = 'SYNTHETIC_PDF_CONTENT_MUST_NOT_BE_LOGGED';
    const token = 'SYNTHETIC_GRAPH_TOKEN_MUST_NOT_BE_LOGGED';
    const logSpy = vi.spyOn(logger, 'info').mockImplementation(() => logger);
    const fetchMock = vi.fn(
      async () =>
        new Response(JSON.stringify({ responses: [] }), {
          status: 200,
          headers: { 'Content-Type': 'application/json' },
        })
    );
    vi.stubGlobal('fetch', fetchMock);

    try {
      const client = new GraphClient(
        {} as ConstructorParameters<typeof GraphClient>[0],
        { cloudType: 'global' } as ConstructorParameters<typeof GraphClient>[1]
      );
      const response = await client.graphRequest('/$batch', {
        method: 'POST',
        accessToken: token,
        body: JSON.stringify({
          requests: [
            {
              id: '1',
              method: 'POST',
              url: '/me/messages/qa/attachments',
              body: { name: 'synthetic.pdf', contentBytes: marker },
            },
          ],
        }),
      });

      expect(response.isError).not.toBe(true);
      expect(fetchMock).toHaveBeenCalledOnce();
      const messages = logSpy.mock.calls.map(([message]) => String(message));
      const requestLog = messages.find((message) =>
        message.includes('Calling /$batch with options:')
      );
      expect(requestLog).toContain('[body=OMITTED]');
      expect(requestLog).toContain('POST');
      expect(messages.join('\\n')).not.toContain(marker);
      expect(messages.join('\\n')).not.toContain(token);
    } finally {
      vi.unstubAllGlobals();
      logSpy.mockRestore();
    }
  });
});
