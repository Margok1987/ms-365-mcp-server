import { afterEach, describe, expect, it, vi } from 'vitest';
import { downloadHostFile, isPublicAddress } from '../src/lib/host-file-param.js';

const hostFile = {
  download_url: 'https://1.1.1.1/file',
  file_id: 'file-1',
  mime_type: 'application/octet-stream',
  file_name: 'sample.bin',
};

describe('host file parameter download', () => {
  afterEach(() => {
    vi.restoreAllMocks();
  });

  it('downloads exact bytes to a private temporary file and cleans it up', async () => {
    const bytes = Buffer.from('hello');
    const fetchImpl = vi.fn(
      async () =>
        new Response(bytes, {
          status: 200,
          headers: {
            'content-length': String(bytes.length),
            'content-type': 'text/plain; charset=utf-8',
          },
        })
    );

    const downloaded = await downloadHostFile(hostFile, 1024, { fetchImpl });
    expect(downloaded.size).toBe(bytes.length);
    expect(downloaded.contentType).toBe('text/plain');
    expect(Buffer.from(await downloaded.handle.readFile())).toEqual(bytes);
    await downloaded.cleanup();
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('refuses non-HTTPS and non-public destinations before fetching', async () => {
    const fetchImpl = vi.fn();
    await expect(
      downloadHostFile({ ...hostFile, download_url: 'http://1.1.1.1/file' }, 1024, {
        fetchImpl,
      })
    ).rejects.toThrow(/HTTPS/);
    await expect(
      downloadHostFile({ ...hostFile, download_url: 'https://127.0.0.1/file' }, 1024, {
        fetchImpl,
      })
    ).rejects.toThrow(/non-public/);
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it('revalidates redirects and refuses a redirect to a private destination', async () => {
    const fetchImpl = vi.fn(
      async () =>
        new Response(null, {
          status: 302,
          headers: { location: 'https://127.0.0.1/private' },
        })
    );

    await expect(downloadHostFile(hostFile, 1024, { fetchImpl })).rejects.toThrow(/non-public/);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('enforces the declared and streamed byte limit', async () => {
    const declared = vi.fn(
      async () =>
        new Response(Buffer.from('x'), {
          status: 200,
          headers: { 'content-length': '11' },
        })
    );
    await expect(downloadHostFile(hostFile, 10, { fetchImpl: declared })).rejects.toThrow(
      /largest attachment/
    );

    const streamed = vi.fn(async () => new Response(Buffer.alloc(11, 1), { status: 200 }));
    await expect(downloadHostFile(hostFile, 10, { fetchImpl: streamed })).rejects.toThrow(
      /largest attachment/
    );
  });

  it('refuses a hostname when DNS returns any private address', async () => {
    const fetchImpl = vi.fn();
    await expect(
      downloadHostFile({ ...hostFile, download_url: 'https://files.example.test/file' }, 1024, {
        fetchImpl,
        lookupAll: async () => ['1.1.1.1', '10.0.0.1'],
      })
    ).rejects.toThrow(/non-public/);
    expect(fetchImpl).not.toHaveBeenCalled();
  });

  it('classifies common private, documentation, mapped, and public IP forms', () => {
    expect(isPublicAddress('1.1.1.1')).toBe(true);
    expect(isPublicAddress('10.0.0.1')).toBe(false);
    expect(isPublicAddress('169.254.1.1')).toBe(false);
    expect(isPublicAddress('127.0.0.1')).toBe(false);
    expect(isPublicAddress('192.0.2.1')).toBe(false);
    expect(isPublicAddress('198.51.100.1')).toBe(false);
    expect(isPublicAddress('203.0.113.1')).toBe(false);
    expect(isPublicAddress('198.51.101.1')).toBe(true);
    expect(isPublicAddress('203.0.114.1')).toBe(true);
    expect(isPublicAddress('::1')).toBe(false);
    expect(isPublicAddress('fd00::1')).toBe(false);
    expect(isPublicAddress('::ffff:127.0.0.1')).toBe(false);
    expect(isPublicAddress('::ffff:7f00:1')).toBe(false);
  });
});
