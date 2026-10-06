import { backoffDelayMs, loadResilienceConfig, parseRetryAfterMs } from './graph-resilience.js';

export const GRAPH_UPLOAD_GRANULARITY = 320 * 1024;
export const DEFAULT_UPLOAD_CHUNK_SIZE = 16 * GRAPH_UPLOAD_GRANULARITY; // 5 MiB
export const MAX_GRAPH_FRAGMENT_SIZE = 60 * 1024 * 1024;

export type FetchLike = (input: string, init?: RequestInit) => Promise<Response>;

export interface UploadSessionTransportOptions {
  uploadUrl: string;
  totalBytes: number;
  source: AsyncIterable<Uint8Array>;
  fetchImpl?: FetchLike;
  chunkSize?: number;
  sleep?: (ms: number) => Promise<void>;
}

export interface UploadSessionTransportResult {
  driveItem: Record<string, unknown>;
  bytesUploaded: number;
  chunksUploaded: number;
}

export class UploadSessionTransportError extends Error {
  constructor(
    public readonly code: string,
    message: string,
    public readonly details: Record<string, unknown> = {}
  ) {
    super(message);
    this.name = 'UploadSessionTransportError';
  }
}

interface UploadSessionRetryPolicy {
  maxRetries: number;
  baseBackoffMs: number;
  maxBackoffMs: number;
  sleep: (ms: number) => Promise<void>;
}

function buildUploadSessionRetryPolicy(
  sleep?: (ms: number) => Promise<void>
): UploadSessionRetryPolicy {
  const config = loadResilienceConfig();
  return {
    maxRetries: config.maxRetries,
    baseBackoffMs: config.baseBackoffMs,
    maxBackoffMs: config.maxBackoffMs,
    sleep: sleep ?? ((ms) => new Promise((resolve) => setTimeout(resolve, ms))),
  };
}

async function fetchUploadSessionWithThrottleRetry(
  fetchImpl: FetchLike,
  uploadUrl: string,
  init: RequestInit,
  retryPolicy: UploadSessionRetryPolicy
): Promise<Response> {
  let attempt = 0;
  while (true) {
    const response = await fetchImpl(uploadUrl, init);
    if (response.status !== 429 || attempt >= retryPolicy.maxRetries) {
      return response;
    }

    const retryAfter = parseRetryAfterMs(response.headers.get('retry-after'));
    const delayMs =
      retryAfter ??
      backoffDelayMs(attempt, retryPolicy.baseBackoffMs, retryPolicy.maxBackoffMs);

    try {
      await response.arrayBuffer();
    } catch {
      // Best-effort drain before retrying so the underlying connection can be reused.
    }

    attempt += 1;
    await retryPolicy.sleep(delayMs);
  }
}

function assertUploadUrl(uploadUrl: string): void {
  let parsed: URL;
  try {
    parsed = new URL(uploadUrl);
  } catch {
    throw new UploadSessionTransportError('INVALID_UPLOAD_URL', 'uploadUrl is not a valid URL.');
  }
  if (parsed.protocol !== 'https:') {
    throw new UploadSessionTransportError(
      'INVALID_UPLOAD_URL',
      'uploadUrl must use HTTPS.'
    );
  }
  if (parsed.username || parsed.password) {
    throw new UploadSessionTransportError(
      'INVALID_UPLOAD_URL',
      'uploadUrl must not contain URL credentials.'
    );
  }
}

function assertChunkSize(chunkSize: number): void {
  if (!Number.isSafeInteger(chunkSize) || chunkSize <= 0) {
    throw new UploadSessionTransportError(
      'INVALID_CHUNK_SIZE',
      'chunkSize must be a positive safe integer.'
    );
  }
  if (chunkSize >= MAX_GRAPH_FRAGMENT_SIZE) {
    throw new UploadSessionTransportError(
      'INVALID_CHUNK_SIZE',
      'chunkSize must be smaller than 60 MiB.'
    );
  }
  if (chunkSize % GRAPH_UPLOAD_GRANULARITY !== 0) {
    throw new UploadSessionTransportError(
      'INVALID_CHUNK_SIZE',
      'chunkSize must be a multiple of 320 KiB.'
    );
  }
}

function rangeStart(range: unknown): number | null {
  if (typeof range !== 'string') return null;
  const match = /^(\d+)-/.exec(range);
  if (!match) return null;
  const value = Number(match[1]);
  return Number.isSafeInteger(value) ? value : null;
}

async function readJson(response: Response): Promise<Record<string, unknown>> {
  const text = await response.text();
  if (!text) return {};
  try {
    const parsed = JSON.parse(text);
    return parsed && typeof parsed === 'object' && !Array.isArray(parsed)
      ? (parsed as Record<string, unknown>)
      : {};
  } catch {
    throw new UploadSessionTransportError(
      'INVALID_PROVIDER_RESPONSE',
      'Upload session returned a non-JSON response.',
      { status: response.status }
    );
  }
}

function nextExpectedStart(body: Record<string, unknown>): number | null {
  const ranges = body.nextExpectedRanges;
  if (!Array.isArray(ranges) || ranges.length === 0) return null;
  return rangeStart(ranges[0]);
}

async function sessionExpectedStart(
  fetchImpl: FetchLike,
  uploadUrl: string,
  retryPolicy: UploadSessionRetryPolicy
): Promise<number | null> {
  const response = await fetchUploadSessionWithThrottleRetry(
    fetchImpl,
    uploadUrl,
    { method: 'GET' },
    retryPolicy
  );
  if (!response.ok) return null;
  return nextExpectedStart(await readJson(response));
}

async function putChunk(
  fetchImpl: FetchLike,
  uploadUrl: string,
  chunk: Uint8Array,
  start: number,
  totalBytes: number,
  retryPolicy: UploadSessionRetryPolicy,
  allowRetryAfterReconcile = true
): Promise<{ finalItem?: Record<string, unknown>; reconciledAccepted?: boolean }> {
  const end = start + chunk.byteLength - 1;
  const isFinal = end + 1 === totalBytes;
  let response: Response;

  try {
    response = await fetchUploadSessionWithThrottleRetry(
      fetchImpl,
      uploadUrl,
      {
        method: 'PUT',
        headers: {
          'Content-Length': String(chunk.byteLength),
          'Content-Range': `bytes ${start}-${end}/${totalBytes}`,
        },
        body: Buffer.from(chunk),
      },
      retryPolicy
    );
  } catch (error) {
    if (isFinal) {
      throw new UploadSessionTransportError(
        'FINAL_CHUNK_RESULT_UNCERTAIN',
        'The final chunk transport failed after dispatch; reconcile the destination DriveItem before any replay.',
        { start, end, cause: error instanceof Error ? error.message : String(error) }
      );
    }

    const expected = await sessionExpectedStart(fetchImpl, uploadUrl, retryPolicy);
    if (expected !== null && expected > end) {
      return { reconciledAccepted: true };
    }
    if (expected === start && allowRetryAfterReconcile) {
      return putChunk(fetchImpl, uploadUrl, chunk, start, totalBytes, retryPolicy, false);
    }
    throw new UploadSessionTransportError(
      'CHUNK_RESULT_UNCERTAIN',
      'Chunk transport failed and upload-session status did not prove a safe retry.',
      { start, end, nextExpectedStart: expected }
    );
  }

  if (response.status === 202) {
    if (isFinal) {
      throw new UploadSessionTransportError(
        'UNEXPECTED_INTERMEDIATE_FINAL',
        'Provider returned 202 for the final declared byte range.',
        { start, end, totalBytes }
      );
    }
    const body = await readJson(response);
    const expected = nextExpectedStart(body);
    if (expected !== end + 1) {
      throw new UploadSessionTransportError(
        'NEXT_RANGE_MISMATCH',
        'Provider nextExpectedRanges did not match the uploaded range.',
        { start, end, expected: end + 1, actual: expected }
      );
    }
    return {};
  }

  if (response.status === 200 || response.status === 201) {
    if (!isFinal) {
      throw new UploadSessionTransportError(
        'UNEXPECTED_EARLY_COMPLETION',
        'Provider completed the upload before all declared bytes were sent.',
        { start, end, totalBytes }
      );
    }
    return { finalItem: await readJson(response) };
  }

  const body = await readJson(response).catch(() => ({}));
  throw new UploadSessionTransportError(
    'UPLOAD_CHUNK_REJECTED',
    `Upload chunk failed with HTTP ${response.status}.`,
    { start, end, status: response.status, provider: body }
  );
}

async function* fixedChunks(
  source: AsyncIterable<Uint8Array>,
  chunkSize: number
): AsyncGenerator<Uint8Array> {
  let pending = Buffer.alloc(0);
  for await (const part of source) {
    if (!(part instanceof Uint8Array)) {
      throw new UploadSessionTransportError(
        'INVALID_SOURCE_CHUNK',
        'Source yielded a non-byte chunk.'
      );
    }
    pending = Buffer.concat([pending, Buffer.from(part)]);
    while (pending.length >= chunkSize) {
      yield pending.subarray(0, chunkSize);
      pending = pending.subarray(chunkSize);
    }
  }
  if (pending.length > 0) yield pending;
}

export async function uploadChunksToSession(
  options: UploadSessionTransportOptions
): Promise<UploadSessionTransportResult> {
  const fetchImpl = options.fetchImpl ?? fetch;
  const chunkSize = options.chunkSize ?? DEFAULT_UPLOAD_CHUNK_SIZE;
  const retryPolicy = buildUploadSessionRetryPolicy(options.sleep);

  assertUploadUrl(options.uploadUrl);
  assertChunkSize(chunkSize);
  if (!Number.isSafeInteger(options.totalBytes) || options.totalBytes <= 0) {
    throw new UploadSessionTransportError(
      'INVALID_TOTAL_BYTES',
      'totalBytes must be a positive safe integer.'
    );
  }

  let offset = 0;
  let chunksUploaded = 0;
  let finalItem: Record<string, unknown> | undefined;

  for await (const chunk of fixedChunks(options.source, chunkSize)) {
    if (offset + chunk.byteLength > options.totalBytes) {
      throw new UploadSessionTransportError(
        'SOURCE_TOO_LONG',
        'Source contains more bytes than declared.',
        { offset, chunkBytes: chunk.byteLength, totalBytes: options.totalBytes }
      );
    }

    const result = await putChunk(
      fetchImpl,
      options.uploadUrl,
      chunk,
      offset,
      options.totalBytes,
      retryPolicy
    );
    offset += chunk.byteLength;
    chunksUploaded += 1;
    if (result.finalItem) finalItem = result.finalItem;
  }

  if (offset !== options.totalBytes) {
    throw new UploadSessionTransportError(
      'SOURCE_TOO_SHORT',
      'Source ended before the declared byte count.',
      { bytesRead: offset, totalBytes: options.totalBytes }
    );
  }
  if (!finalItem) {
    throw new UploadSessionTransportError(
      'MISSING_FINAL_DRIVEITEM',
      'All declared bytes were sent but the provider did not return a final DriveItem.'
    );
  }

  return {
    driveItem: finalItem,
    bytesUploaded: offset,
    chunksUploaded,
  };
}
