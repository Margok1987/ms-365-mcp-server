import { createHash } from 'crypto';
import { backoffDelayMs, loadResilienceConfig, parseRetryAfterMs } from './graph-resilience.js';
import type { FetchLike } from './upload-session-transport.js';

export const EVENT_ATTACHMENT_MIN_UPLOAD_SESSION_BYTES = 3 * 1024 * 1024;
export const EVENT_ATTACHMENT_MAX_BYTES = 150 * 1024 * 1024;
export const EVENT_ATTACHMENT_CHUNK_BYTES = 4 * 1024 * 1024;

export interface EventAttachmentUploadSource {
  name: string;
  size: number;
  open(): AsyncIterable<Uint8Array>;
}

export interface EventAttachmentEventSnapshot {
  id?: string;
  type?: string;
  isCancelled?: boolean;
  isOrganizer?: boolean;
  attendees?: unknown[];
  changeKey?: string;
  lastModifiedDateTime?: string;
}

export interface EventAttachmentSnapshot {
  id?: string;
  name?: string;
  isInline?: boolean;
  contentId?: string | null;
  size?: number;
}

export interface EventAttachmentUploadDependencies {
  readEvent(eventId: string): Promise<EventAttachmentEventSnapshot>;
  listAttachments(eventId: string): Promise<EventAttachmentSnapshot[]>;
  createUploadSession(input: {
    eventId: string;
    name: string;
    size: number;
    isInline: boolean;
    contentId?: string;
  }): Promise<{ uploadUrl: string }>;
  readAttachmentBytes(eventId: string, attachmentId: string): Promise<AsyncIterable<Uint8Array>>;
  fetchImpl?: FetchLike;
  sleep?: (ms: number) => Promise<void>;
}

export interface EventAttachmentUploadInput {
  eventId: string;
  idKind: 'restImmutableEntryId';
  source: EventAttachmentUploadSource;
  isInline?: boolean;
  contentId?: string;
}

export interface EventAttachmentUploadResult {
  attachmentId: string;
  name: string;
  sourceBytes: number;
  sha256: string;
  chunksUploaded: number;
  reconciledAfterUncertainFinal: boolean;
  eventChangeKeyBefore?: string;
  eventLastModifiedBefore?: string;
}

export class EventAttachmentUploadError extends Error {
  constructor(
    public readonly code: string,
    message: string,
    public readonly details: Record<string, unknown> = {}
  ) {
    super(message);
    this.name = 'EventAttachmentUploadError';
  }
}

function requiredString(value: unknown, field: string): string {
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new EventAttachmentUploadError('INVALID_INPUT', `${field} must be a non-empty string.`);
  }
  return value.trim();
}

function assertSource(source: EventAttachmentUploadSource): void {
  requiredString(source.name, 'source.name');
  if (!Number.isSafeInteger(source.size)) {
    throw new EventAttachmentUploadError('INVALID_FILE_SIZE', 'source.size must be a safe integer.');
  }
  if (source.size < EVENT_ATTACHMENT_MIN_UPLOAD_SESSION_BYTES) {
    throw new EventAttachmentUploadError(
      'FILE_TOO_SMALL_FOR_UPLOAD_SESSION',
      'Large event attachment upload requires a file of at least 3 MiB; use add-event-attachment for smaller files.'
    );
  }
  if (source.size > EVENT_ATTACHMENT_MAX_BYTES) {
    throw new EventAttachmentUploadError(
      'FILE_TOO_LARGE',
      'Event attachment upload supports files up to 150 MiB.'
    );
  }
}

function assertSyntheticInternalEvent(eventId: string, event: EventAttachmentEventSnapshot): void {
  if (event.id !== eventId) {
    throw new EventAttachmentUploadError(
      'EVENT_IDENTITY_MISMATCH',
      'Immutable event readback did not return the exact requested event id.'
    );
  }
  if (event.type !== 'singleInstance') {
    throw new EventAttachmentUploadError(
      'EVENT_TYPE_NOT_ALLOWED',
      'This qualified upload operation is limited to singleInstance events.'
    );
  }
  if (event.isCancelled === true) {
    throw new EventAttachmentUploadError('EVENT_CANCELLED', 'Cancelled events are not writable fixtures.');
  }
  if (event.isOrganizer !== true) {
    throw new EventAttachmentUploadError(
      'EVENT_NOT_ORGANIZER_OWNED',
      'This qualified upload operation requires an organizer-owned event.'
    );
  }
  if (!Array.isArray(event.attendees)) {
    throw new EventAttachmentUploadError(
      'EVENT_ATTENDEES_UNKNOWN',
      'Exact preflight must return the attendee collection.'
    );
  }
  if (event.attendees.length !== 0) {
    throw new EventAttachmentUploadError(
      'EVENT_HAS_ATTENDEES',
      'This qualified upload operation refuses events with attendees to avoid meeting-update fan-out.'
    );
  }
}

async function hashBytes(source: AsyncIterable<Uint8Array>): Promise<string> {
  const hash = createHash('sha256');
  for await (const part of source) {
    if (!(part instanceof Uint8Array)) {
      throw new EventAttachmentUploadError('INVALID_SOURCE_CHUNK', 'Byte source yielded a non-byte chunk.');
    }
    hash.update(part);
  }
  return hash.digest('hex');
}

async function* fixedChunks(
  source: AsyncIterable<Uint8Array>,
  chunkSize: number
): AsyncGenerator<Uint8Array> {
  let pending = Buffer.alloc(0);
  for await (const part of source) {
    if (!(part instanceof Uint8Array)) {
      throw new EventAttachmentUploadError('INVALID_SOURCE_CHUNK', 'Byte source yielded a non-byte chunk.');
    }
    pending = Buffer.concat([pending, Buffer.from(part)]);
    while (pending.length >= chunkSize) {
      yield pending.subarray(0, chunkSize);
      pending = pending.subarray(chunkSize);
    }
  }
  if (pending.length > 0) yield pending;
}

function parseNextExpectedStart(body: Record<string, unknown>): number | null {
  const ranges = body.nextExpectedRanges;
  if (!Array.isArray(ranges) || typeof ranges[0] !== 'string') return null;
  const match = /^(\d+)-/.exec(ranges[0]);
  if (!match) return null;
  const n = Number(match[1]);
  return Number.isSafeInteger(n) ? n : null;
}

async function readJson(response: Response): Promise<Record<string, unknown>> {
  const text = await response.text();
  if (!text) return {};
  try {
    const value = JSON.parse(text);
    return value && typeof value === 'object' && !Array.isArray(value)
      ? (value as Record<string, unknown>)
      : {};
  } catch {
    throw new EventAttachmentUploadError(
      'INVALID_UPLOAD_SESSION_RESPONSE',
      'Upload session returned non-JSON progress data.',
      { status: response.status }
    );
  }
}

function attachmentIdFromLocation(raw: string | null): string | null {
  if (!raw) return null;
  try {
    const url = new URL(raw);
    const segments = url.pathname.split('/').filter(Boolean);
    const marker = segments.findIndex((segment) => segment.toLowerCase() === 'attachments');
    if (marker < 0 || marker + 1 >= segments.length) return null;
    const decoded = decodeURIComponent(segments[marker + 1]);
    return decoded.length > 0 ? decoded : null;
  } catch {
    return null;
  }
}

interface RetryPolicy {
  maxRetries: number;
  baseBackoffMs: number;
  maxBackoffMs: number;
  sleep: (ms: number) => Promise<void>;
}

function retryPolicy(sleep?: (ms: number) => Promise<void>): RetryPolicy {
  const config = loadResilienceConfig();
  return {
    maxRetries: config.maxRetries,
    baseBackoffMs: config.baseBackoffMs,
    maxBackoffMs: config.maxBackoffMs,
    sleep: sleep ?? ((ms) => new Promise((resolve) => setTimeout(resolve, ms))),
  };
}

async function putRangeWithExplicitThrottleRetry(
  fetchImpl: FetchLike,
  uploadUrl: string,
  init: RequestInit,
  policy: RetryPolicy
): Promise<Response> {
  let attempt = 0;
  while (true) {
    const response = await fetchImpl(uploadUrl, init);
    if (response.status !== 429 || attempt >= policy.maxRetries) return response;
    const retryAfter =
      parseRetryAfterMs(response.headers.get('retry-after')) ??
      backoffDelayMs(attempt, policy.baseBackoffMs, policy.maxBackoffMs);
    try {
      await response.arrayBuffer();
    } catch {
      // Best-effort drain before retrying an explicit provider throttle.
    }
    attempt += 1;
    await policy.sleep(retryAfter);
  }
}

async function reconcileCreatedAttachment(
  eventId: string,
  sourceHash: string,
  beforeIds: ReadonlySet<string>,
  deps: EventAttachmentUploadDependencies,
  preferredId?: string | null
): Promise<{ attachmentId: string } | null> {
  const after = await deps.listAttachments(eventId);
  const candidates = after
    .map((item) => item.id)
    .filter((id): id is string => typeof id === 'string' && id.length > 0 && !beforeIds.has(id));

  if (preferredId && !beforeIds.has(preferredId) && !candidates.includes(preferredId)) {
    candidates.unshift(preferredId);
  }

  const unique = [...new Set(candidates)];
  const matches: string[] = [];
  for (const attachmentId of unique) {
    try {
      const actualHash = await hashBytes(await deps.readAttachmentBytes(eventId, attachmentId));
      if (actualHash === sourceHash) matches.push(attachmentId);
    } catch {
      // A candidate that cannot be read exactly cannot prove success.
    }
  }
  return matches.length === 1 ? { attachmentId: matches[0] } : null;
}

export async function uploadLargeFileAttachmentToEvent(
  input: EventAttachmentUploadInput,
  deps: EventAttachmentUploadDependencies
): Promise<EventAttachmentUploadResult> {
  const eventId = requiredString(input.eventId, 'eventId');
  if (input.idKind !== 'restImmutableEntryId') {
    throw new EventAttachmentUploadError(
      'INVALID_ID_KIND',
      'eventId must be declared as restImmutableEntryId.'
    );
  }
  assertSource(input.source);

  const event = await deps.readEvent(eventId);
  assertSyntheticInternalEvent(eventId, event);

  const before = await deps.listAttachments(eventId);
  const beforeIds = new Set(
    before
      .map((item) => item.id)
      .filter((id): id is string => typeof id === 'string' && id.length > 0)
  );

  const sourceHash = await hashBytes(input.source.open());
  const session = await deps.createUploadSession({
    eventId,
    name: input.source.name,
    size: input.source.size,
    isInline: input.isInline === true,
    ...(input.contentId ? { contentId: input.contentId } : {}),
  });
  if (!session || typeof session.uploadUrl !== 'string' || session.uploadUrl.length === 0) {
    throw new EventAttachmentUploadError(
      'UPLOAD_SESSION_INVALID',
      'Provider did not return an event attachment uploadUrl.'
    );
  }

  let offset = 0;
  let chunksUploaded = 0;
  const fetchImpl = deps.fetchImpl ?? fetch;
  const policy = retryPolicy(deps.sleep);

  for await (const chunk of fixedChunks(input.source.open(), EVENT_ATTACHMENT_CHUNK_BYTES)) {
    if (offset + chunk.byteLength > input.source.size) {
      throw new EventAttachmentUploadError('SOURCE_TOO_LONG', 'Byte source exceeds declared size.');
    }

    const start = offset;
    const end = start + chunk.byteLength - 1;
    const isFinal = end + 1 === input.source.size;
    let response: Response;

    try {
      response = await putRangeWithExplicitThrottleRetry(
        fetchImpl,
        session.uploadUrl,
        {
          method: 'PUT',
          headers: {
            'Content-Length': String(chunk.byteLength),
            'Content-Range': `bytes ${start}-${end}/${input.source.size}`,
            'Content-Type': 'application/octet-stream',
          },
          body: Buffer.from(chunk),
        },
        policy
      );
    } catch (error) {
      if (!isFinal) {
        throw new EventAttachmentUploadError(
          'CHUNK_RESULT_UNCERTAIN',
          'An intermediate attachment range lost its transport receipt. Do not replay the range blindly; the upload session must be abandoned/reconciled.',
          { start, end, cause: error instanceof Error ? error.message : String(error) }
        );
      }

      const reconciled = await reconcileCreatedAttachment(
        eventId,
        sourceHash,
        beforeIds,
        deps
      ).catch(() => null);
      if (!reconciled) {
        throw new EventAttachmentUploadError(
          'FINAL_CHUNK_RESULT_UNCERTAIN',
          'The final attachment range lost its transport receipt and exact event readback did not prove one unique matching new attachment.',
          { start, end, cause: error instanceof Error ? error.message : String(error) }
        );
      }
      offset += chunk.byteLength;
      chunksUploaded += 1;
      return {
        attachmentId: reconciled.attachmentId,
        name: input.source.name,
        sourceBytes: input.source.size,
        sha256: sourceHash,
        chunksUploaded,
        reconciledAfterUncertainFinal: true,
        ...(event.changeKey ? { eventChangeKeyBefore: event.changeKey } : {}),
        ...(event.lastModifiedDateTime
          ? { eventLastModifiedBefore: event.lastModifiedDateTime }
          : {}),
      };
    }

    if (!isFinal) {
      if (response.status !== 200) {
        throw new EventAttachmentUploadError(
          'UPLOAD_RANGE_REJECTED',
          `Intermediate event attachment range returned HTTP ${response.status}; no automatic replay is allowed.`,
          { start, end, status: response.status }
        );
      }
      const progress = await readJson(response);
      const nextStart = parseNextExpectedStart(progress);
      if (nextStart !== end + 1) {
        throw new EventAttachmentUploadError(
          'NEXT_RANGE_MISMATCH',
          'Provider nextExpectedRanges did not prove acceptance of the exact intermediate range.',
          { start, end, expected: end + 1, actual: nextStart }
        );
      }
      offset += chunk.byteLength;
      chunksUploaded += 1;
      continue;
    }

    if (response.status !== 201) {
      throw new EventAttachmentUploadError(
        'FINAL_UPLOAD_RESPONSE_INVALID',
        `Final event attachment range returned HTTP ${response.status} instead of 201.`,
        { start, end, status: response.status }
      );
    }

    const preferredId = attachmentIdFromLocation(response.headers.get('location'));
    const reconciled = await reconcileCreatedAttachment(
      eventId,
      sourceHash,
      beforeIds,
      deps,
      preferredId
    );
    if (!reconciled) {
      throw new EventAttachmentUploadError(
        'UPLOAD_READBACK_UNPROVEN',
        'Provider returned final success but exact event attachment readback did not prove one unique byte-identical new attachment.',
        { preferredAttachmentIdPresent: preferredId !== null }
      );
    }

    offset += chunk.byteLength;
    chunksUploaded += 1;
    return {
      attachmentId: reconciled.attachmentId,
      name: input.source.name,
      sourceBytes: input.source.size,
      sha256: sourceHash,
      chunksUploaded,
      reconciledAfterUncertainFinal: false,
      ...(event.changeKey ? { eventChangeKeyBefore: event.changeKey } : {}),
      ...(event.lastModifiedDateTime
        ? { eventLastModifiedBefore: event.lastModifiedDateTime }
        : {}),
    };
  }

  if (offset !== input.source.size) {
    throw new EventAttachmentUploadError('SOURCE_TOO_SHORT', 'Byte source ended before declared size.');
  }

  throw new EventAttachmentUploadError(
    'UPLOAD_INCOMPLETE',
    'Attachment upload ended without a final provider response.'
  );
}
