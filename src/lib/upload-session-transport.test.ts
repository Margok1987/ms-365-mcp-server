import { describe, expect, it, vi } from 'vitest';
import {
  DEFAULT_UPLOAD_CHUNK_SIZE,
  GRAPH_UPLOAD_GRANULARITY,
  uploadChunksToSession,
} from './upload-session-transport.js';

async function* bytes(...parts: Uint8Array[]) {
  for (const part of parts) yield part;
}

function json(status: number, body: unknown) {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json' },
  });
}

describe('uploadChunksToSession', () => {
  it('uploads one final chunk and returns the DriveItem', async () => {
    const fetchImpl = vi.fn(async (_url: string, init?: RequestInit) => {
      expect(init?.method).toBe('PUT');
      expect(new Headers(init?.headers).get('content-range')).toBe('bytes 0-4/5');
      return json(201, { id: 'item-1', size: 5 });
    });

    const result = await uploadChunksToSession({
      uploadUrl: 'https://upload.example.test/session',
      totalBytes: 5,
      source: bytes(new Uint8Array([1, 2, 3, 4, 5])),
      fetchImpl,
      chunkSize: GRAPH_UPLOAD_GRANULARITY,
    });

    expect(result.driveItem.id).toBe('item-1');
    expect(result.bytesUploaded).toBe(5);
    expect(result.chunksUploaded).toBe(1);
  });

  it('uses sequential Graph ranges and validates nextExpectedRanges', async () => {
    const first = new Uint8Array(GRAPH_UPLOAD_GRANULARITY);
    const second = new Uint8Array(7);
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(
        json(202, { nextExpectedRanges: [`${GRAPH_UPLOAD_GRANULARITY}-`] })
      )
      .mockResolvedValueOnce(
        json(201, { id: 'item-2', size: GRAPH_UPLOAD_GRANULARITY + 7 })
      );

    const result = await uploadChunksToSession({
      uploadUrl: 'https://upload.example.test/session',
      totalBytes: GRAPH_UPLOAD_GRANULARITY + 7,
      source: bytes(first, second),
      fetchImpl,
      chunkSize: GRAPH_UPLOAD_GRANULARITY,
    });

    const firstHeaders = new Headers(fetchImpl.mock.calls[0][1]?.headers);
    const secondHeaders = new Headers(fetchImpl.mock.calls[1][1]?.headers);
    expect(firstHeaders.get('content-range')).toBe(
      `bytes 0-${GRAPH_UPLOAD_GRANULARITY - 1}/${GRAPH_UPLOAD_GRANULARITY + 7}`
    );
    expect(secondHeaders.get('content-range')).toBe(
      `bytes ${GRAPH_UPLOAD_GRANULARITY}-${GRAPH_UPLOAD_GRANULARITY + 6}/${GRAPH_UPLOAD_GRANULARITY + 7}`
    );
    expect(result.chunksUploaded).toBe(2);
  });

  it('rejects a non-320-KiB chunk size', async () => {
    await expect(
      uploadChunksToSession({
        uploadUrl: 'https://upload.example.test/session',
        totalBytes: 1,
        source: bytes(new Uint8Array([1])),
        fetchImpl: vi.fn(),
        chunkSize: 123,
      })
    ).rejects.toMatchObject({ code: 'INVALID_CHUNK_SIZE' });
  });

  it('rejects mismatched nextExpectedRanges', async () => {
    const fetchImpl = vi.fn().mockResolvedValue(
      json(202, { nextExpectedRanges: ['1-'] })
    );
    await expect(
      uploadChunksToSession({
        uploadUrl: 'https://upload.example.test/session',
        totalBytes: GRAPH_UPLOAD_GRANULARITY + 1,
        source: bytes(
          new Uint8Array(GRAPH_UPLOAD_GRANULARITY),
          new Uint8Array([1])
        ),
        fetchImpl,
        chunkSize: GRAPH_UPLOAD_GRANULARITY,
      })
    ).rejects.toMatchObject({ code: 'NEXT_RANGE_MISMATCH' });
  });

  it('reconciles a failed intermediate PUT before one safe retry', async () => {
    let putAttempts = 0;
    const fetchImpl = vi.fn(async (_url: string, init?: RequestInit) => {
      if (init?.method === 'GET') {
        return json(200, { nextExpectedRanges: ['0-'] });
      }
      putAttempts += 1;
      if (putAttempts === 1) throw new Error('network lost');
      if (putAttempts === 2) {
        return json(202, { nextExpectedRanges: [`${GRAPH_UPLOAD_GRANULARITY}-`] });
      }
      return json(201, { id: 'item-retried' });
    });

    const result = await uploadChunksToSession({
      uploadUrl: 'https://upload.example.test/session',
      totalBytes: GRAPH_UPLOAD_GRANULARITY + 1,
      source: bytes(
        new Uint8Array(GRAPH_UPLOAD_GRANULARITY),
        new Uint8Array([1])
      ),
      fetchImpl,
      chunkSize: GRAPH_UPLOAD_GRANULARITY,
    });
    expect(result.driveItem.id).toBe('item-retried');
    expect(putAttempts).toBe(3);
  });

  it('treats a final transport failure as uncertain instead of replaying it', async () => {
    const fetchImpl = vi.fn(async () => {
      throw new Error('connection reset');
    });

    await expect(
      uploadChunksToSession({
        uploadUrl: 'https://upload.example.test/session',
        totalBytes: 1,
        source: bytes(new Uint8Array([1])),
        fetchImpl,
        chunkSize: GRAPH_UPLOAD_GRANULARITY,
      })
    ).rejects.toMatchObject({ code: 'FINAL_CHUNK_RESULT_UNCERTAIN' });
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('rejects a source longer than declared', async () => {
    await expect(
      uploadChunksToSession({
        uploadUrl: 'https://upload.example.test/session',
        totalBytes: 1,
        source: bytes(new Uint8Array([1, 2])),
        fetchImpl: vi.fn(),
        chunkSize: GRAPH_UPLOAD_GRANULARITY,
      })
    ).rejects.toMatchObject({ code: 'SOURCE_TOO_LONG' });
  });

  it('rejects a source shorter than declared', async () => {
    const fetchImpl = vi
      .fn()
      .mockResolvedValue(json(202, { nextExpectedRanges: ['1-'] }));
    await expect(
      uploadChunksToSession({
        uploadUrl: 'https://upload.example.test/session',
        totalBytes: GRAPH_UPLOAD_GRANULARITY + 1,
        source: bytes(new Uint8Array([1])),
        fetchImpl,
        chunkSize: GRAPH_UPLOAD_GRANULARITY,
      })
    ).rejects.toMatchObject({ code: 'SOURCE_TOO_SHORT' });
  });

  it('keeps the default chunk size Graph-aligned', () => {
    expect(DEFAULT_UPLOAD_CHUNK_SIZE % GRAPH_UPLOAD_GRANULARITY).toBe(0);
    expect(DEFAULT_UPLOAD_CHUNK_SIZE).toBeLessThan(60 * 1024 * 1024);
  });
});
