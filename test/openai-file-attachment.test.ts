import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest';
import { UTILITY_TOOLS } from '../src/graph-tools.js';
import { downloadHostFile } from '../src/lib/host-file-param.js';

const DRAFT_ATTACHMENTS = '/me/messages/AAA/attachments';
const CHUNK = 320 * 1024 * 12;
const tool = UTILITY_TOOLS.find((entry) => entry.name === 'get-upload-url')!;

function makeContext(graphClient: { makeRequest: ReturnType<typeof vi.fn> }) {
  return {
    graphClient: graphClient as never,
    authManager: {
      isOAuthModeEnabled: () => false,
      isMultiAccount: async () => false,
      getToken: async () => 'SERVER_TOKEN',
      getTokenForAccount: async () => 'SERVER_TOKEN',
    } as never,
    multiAccount: false,
    accountNames: [],
    hostFileDownloader: (file: Parameters<typeof downloadHostFile>[0], maxBytes: number) =>
      downloadHostFile(file, maxBytes, { fetchImpl: global.fetch }),
  };
}

function parse(result: { content: Array<{ text: string }> }) {
  return JSON.parse(result.content[0].text);
}

describe('get-upload-url host-file branch', () => {
  let originalFetch: typeof global.fetch;

  beforeEach(() => {
    originalFetch = global.fetch;
  });

  afterEach(() => {
    global.fetch = originalFetch;
    vi.restoreAllMocks();
  });

  it('downloads a host file and attaches a small file inline byte-identically', async () => {
    const source = Buffer.from('host-file-bytes');
    global.fetch = vi.fn(
      async () =>
        new Response(source, {
          status: 200,
          headers: {
            'content-length': String(source.length),
            'content-type': 'application/octet-stream',
          },
        })
    ) as never;

    const graphClient = { makeRequest: vi.fn().mockResolvedValue({ id: 'att-small' }) };
    const result = await tool.execute(
      {
        target: DRAFT_ATTACHMENTS,
        name: 'report.pdf',
        file: {
          download_url: 'https://1.1.1.1/small',
          file_id: 'file-small',
          mime_type: 'application/pdf',
          file_name: 'report.pdf',
        },
      },
      makeContext(graphClient)
    );

    expect(result.isError).toBeFalsy();
    expect(parse(result as never)).toEqual({
      name: 'report.pdf',
      size: source.length,
      id: 'att-small',
    });
    expect(graphClient.makeRequest).toHaveBeenCalledTimes(1);
    const [target, options] = graphClient.makeRequest.mock.calls[0];
    expect(target).toBe(DRAFT_ATTACHMENTS);
    expect(options.method).toBe('POST');
    expect(options.accessToken).toBe('SERVER_TOKEN');
    expect(JSON.parse(options.body)).toEqual({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'report.pdf',
      contentType: 'application/pdf',
      contentBytes: source.toString('base64'),
    });
  });

  it('downloads a host file once and streams a large file through a Graph upload session', async () => {
    const size = 4 * 1024 * 1024;
    const source = Buffer.alloc(size, 7);
    const puts: Array<{ range: string; length: number }> = [];

    global.fetch = vi.fn(async (input: unknown, init?: Parameters<typeof fetch>[1]) => {
      const url = String(input);
      if (url === 'https://1.1.1.1/large') {
        return new Response(source, {
          status: 200,
          headers: {
            'content-length': String(size),
            'content-type': 'application/octet-stream',
          },
        });
      }
      if (url === 'https://up.example/session' && init?.method === 'PUT') {
        const headers = init.headers as Record<string, string>;
        const body = init.body as ArrayBuffer;
        puts.push({ range: headers['Content-Range'], length: body.byteLength });
        return new Response(null, { status: 202 });
      }
      throw new Error(`Unexpected fetch ${url}`);
    }) as never;

    const graphClient = {
      makeRequest: vi.fn().mockResolvedValue({ uploadUrl: 'https://up.example/session' }),
    };
    const result = await tool.execute(
      {
        target: DRAFT_ATTACHMENTS,
        name: 'large.bin',
        file: {
          download_url: 'https://1.1.1.1/large',
          file_id: 'file-large',
          file_name: 'large.bin',
        },
      },
      makeContext(graphClient)
    );

    expect(result.isError).toBeFalsy();
    expect(parse(result as never)).toEqual({ name: 'large.bin', size });
    expect(graphClient.makeRequest).toHaveBeenCalledTimes(1);
    const [target, options] = graphClient.makeRequest.mock.calls[0];
    expect(target).toBe(`${DRAFT_ATTACHMENTS}/createUploadSession`);
    expect(options.accessToken).toBe('SERVER_TOKEN');
    expect(JSON.parse(options.body).AttachmentItem).toEqual({
      attachmentType: 'file',
      name: 'large.bin',
      size,
      contentType: 'application/octet-stream',
    });
    expect(puts).toEqual([
      { range: `bytes 0-${CHUNK - 1}/${size}`, length: CHUNK },
      { range: `bytes ${CHUNK}-${size - 1}/${size}`, length: size - CHUNK },
    ]);
  });
});
