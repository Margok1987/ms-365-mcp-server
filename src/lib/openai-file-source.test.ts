import { mkdtemp, readFile, stat } from 'fs/promises';
import { tmpdir } from 'os';
import path from 'path';
import { describe, expect, it, vi } from 'vitest';
import {
  assertSafeFileDownloadUrl,
  stageOpenAIFile,
} from './openai-file-source.js';

async function collect(source: AsyncIterable<Uint8Array>): Promise<Buffer> {
  const chunks: Buffer[] = [];
  for await (const chunk of source) chunks.push(Buffer.from(chunk));
  return Buffer.concat(chunks);
}

describe('OpenAI file-param source', () => {
  it('stages a host-provided HTTPS file without base64', async () => {
    const tempRoot = await mkdtemp(path.join(tmpdir(), 'ms365-file-test-'));
    const fetchImpl = vi.fn(async () => new Response(new Uint8Array([1, 2, 3, 4])));
    const staged = await stageOpenAIFile(
      {
        download_url: 'https://files.example.test/file',
        file_id: 'file_123',
        file_name: 'report.pdf',
        mime_type: 'application/pdf',
      },
      { fetchImpl: fetchImpl as typeof fetch, tempRoot }
    );

    expect(staged.size).toBe(4);
    expect(staged.fileName).toBe('report.pdf');
    expect(await collect(staged.open())).toEqual(Buffer.from([1, 2, 3, 4]));
    expect(await readFile(staged.path)).toEqual(Buffer.from([1, 2, 3, 4]));
    await staged.cleanup();
    await expect(stat(staged.path)).rejects.toThrow();
  });

  it('follows a bounded HTTPS redirect', async () => {
    const tempRoot = await mkdtemp(path.join(tmpdir(), 'ms365-file-test-'));
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(
        new Response(null, {
          status: 302,
          headers: { location: 'https://cdn.example.test/file' },
        })
      )
      .mockResolvedValueOnce(new Response(new Uint8Array([9])));

    const staged = await stageOpenAIFile(
      { download_url: 'https://files.example.test/start', file_id: 'file_1' },
      { fetchImpl: fetchImpl as typeof fetch, tempRoot }
    );
    expect(staged.size).toBe(1);
    expect(fetchImpl).toHaveBeenCalledTimes(2);
    await staged.cleanup();
  });

  it.each([
    'http://files.example.test/file',
    'https://localhost/file',
    'https://127.0.0.1/file',
    'https://10.1.2.3/file',
    'https://169.254.1.1/file',
    'https://[::1]/file',
  ])('rejects unsafe source URL %s', (url) => {
    expect(() => assertSafeFileDownloadUrl(url)).toThrow();
  });

  it('rejects a redirect to a private target', async () => {
    const fetchImpl = vi.fn(async () =>
      new Response(null, {
        status: 302,
        headers: { location: 'https://127.0.0.1/private' },
      })
    );
    await expect(
      stageOpenAIFile(
        { download_url: 'https://files.example.test/start', file_id: 'file_2' },
        { fetchImpl: fetchImpl as typeof fetch }
      )
    ).rejects.toMatchObject({ code: 'UNSAFE_FILE_URL' });
  });

  it('rejects non-success file downloads', async () => {
    const fetchImpl = vi.fn(async () => new Response('no', { status: 403 }));
    await expect(
      stageOpenAIFile(
        { download_url: 'https://files.example.test/file', file_id: 'file_3' },
        { fetchImpl: fetchImpl as typeof fetch }
      )
    ).rejects.toMatchObject({ code: 'FILE_DOWNLOAD_REJECTED' });
  });
});
