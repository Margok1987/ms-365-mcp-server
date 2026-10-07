import { describe, expect, it, vi } from 'vitest';
import {
  EVENT_ATTACHMENT_CHUNK_BYTES,
  EventAttachmentUploadError,
  uploadLargeFileAttachmentToEvent,
} from './event-attachment-upload.js';

function source(size: number, fill = 7) {
  return {
    name: 'calendar-qa.bin',
    size,
    open: async function* () {
      const full = new Uint8Array(size);
      full.fill(fill);
      yield full;
    },
  };
}

function event(id = 'immutable-event') {
  return {
    id,
    type: 'singleInstance',
    isCancelled: false,
    isOrganizer: true,
    attendees: [],
    changeKey: 'before',
  };
}

async function* byteStream(bytes: Uint8Array) {
  yield bytes;
}

describe('uploadLargeFileAttachmentToEvent', () => {
  it('uses Outlook attachment range semantics and proves bytes by exact readback', async () => {
    const size = EVENT_ATTACHMENT_CHUNK_BYTES + 3;
    const body = new Uint8Array(size);
    body.fill(7);
    let attachments: Array<{ id: string }> = [];
    const fetchImpl = vi
      .fn()
      .mockResolvedValueOnce(
        new Response(JSON.stringify({ nextExpectedRanges: [`${EVENT_ATTACHMENT_CHUNK_BYTES}-`] }), {
          status: 200,
          headers: { 'content-type': 'application/json' },
        })
      )
      .mockImplementationOnce(async (_url: string, init?: RequestInit) => {
        expect(new Headers(init?.headers).has('authorization')).toBe(false);
        attachments = [{ id: 'att-1' }];
        return new Response(null, {
          status: 201,
          headers: {
            location:
              'https://outlook.office.com/api/v2.0/me/events/immutable-event/attachments/att-1',
          },
        });
      });

    const result = await uploadLargeFileAttachmentToEvent(
      { eventId: 'immutable-event', idKind: 'restImmutableEntryId', source: source(size) },
      {
        readEvent: vi.fn(async () => event()),
        listAttachments: vi.fn(async () => attachments),
        createUploadSession: vi.fn(async () => ({ uploadUrl: 'https://upload.example.test/a' })),
        readAttachmentBytes: vi.fn(async () => byteStream(body)),
        fetchImpl,
      }
    );

    expect(result.attachmentId).toBe('att-1');
    expect(result.reconciledAfterUncertainFinal).toBe(false);
    expect(fetchImpl).toHaveBeenCalledTimes(2);
    const firstHeaders = new Headers(fetchImpl.mock.calls[0][1]?.headers);
    const secondHeaders = new Headers(fetchImpl.mock.calls[1][1]?.headers);
    expect(firstHeaders.get('content-range')).toBe(
      `bytes 0-${EVENT_ATTACHMENT_CHUNK_BYTES - 1}/${size}`
    );
    expect(secondHeaders.get('content-range')).toBe(
      `bytes ${EVENT_ATTACHMENT_CHUNK_BYTES}-${size - 1}/${size}`
    );
    expect(firstHeaders.get('content-type')).toBe('application/octet-stream');
  });

  it('reconciles an uncertain final transport by new attachment id plus byte hash', async () => {
    const size = 3 * 1024 * 1024;
    const body = new Uint8Array(size);
    body.fill(7);
    let attachments: Array<{ id: string }> = [];
    const fetchImpl = vi.fn(async () => {
      attachments = [{ id: 'att-after-uncertain' }];
      throw new Error('connection reset after final PUT');
    });

    const result = await uploadLargeFileAttachmentToEvent(
      { eventId: 'immutable-event', idKind: 'restImmutableEntryId', source: source(size) },
      {
        readEvent: vi.fn(async () => event()),
        listAttachments: vi.fn(async () => attachments),
        createUploadSession: vi.fn(async () => ({ uploadUrl: 'https://upload.example.test/a' })),
        readAttachmentBytes: vi.fn(async () => byteStream(body)),
        fetchImpl,
      }
    );

    expect(result.attachmentId).toBe('att-after-uncertain');
    expect(result.reconciledAfterUncertainFinal).toBe(true);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('does not replay an uncertain intermediate range', async () => {
    const size = EVENT_ATTACHMENT_CHUNK_BYTES + 1;
    const fetchImpl = vi.fn(async () => {
      throw new Error('connection reset');
    });

    await expect(
      uploadLargeFileAttachmentToEvent(
        { eventId: 'immutable-event', idKind: 'restImmutableEntryId', source: source(size) },
        {
          readEvent: vi.fn(async () => event()),
          listAttachments: vi.fn(async () => []),
          createUploadSession: vi.fn(async () => ({ uploadUrl: 'https://upload.example.test/a' })),
          readAttachmentBytes: vi.fn(),
          fetchImpl,
        }
      )
    ).rejects.toMatchObject({ code: 'CHUNK_RESULT_UNCERTAIN' });
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('rejects any attendee-bearing event before session creation', async () => {
    const createUploadSession = vi.fn();
    await expect(
      uploadLargeFileAttachmentToEvent(
        { eventId: 'immutable-event', idKind: 'restImmutableEntryId', source: source(3 * 1024 * 1024) },
        {
          readEvent: vi.fn(async () => ({ ...event(), attendees: [{ emailAddress: { address: 'x@example.test' } }] })),
          listAttachments: vi.fn(async () => []),
          createUploadSession,
          readAttachmentBytes: vi.fn(),
        }
      )
    ).rejects.toMatchObject({ code: 'EVENT_HAS_ATTENDEES' });
    expect(createUploadSession).not.toHaveBeenCalled();
  });

  it('rejects a non-immutable id kind before provider calls', async () => {
    const readEvent = vi.fn();
    await expect(
      uploadLargeFileAttachmentToEvent(
        {
          eventId: 'rest-id',
          idKind: 'restId' as never,
          source: source(3 * 1024 * 1024),
        },
        {
          readEvent,
          listAttachments: vi.fn(),
          createUploadSession: vi.fn(),
          readAttachmentBytes: vi.fn(),
        }
      )
    ).rejects.toMatchObject({ code: 'INVALID_ID_KIND' });
    expect(readEvent).not.toHaveBeenCalled();
  });

  it('rejects exact-read identity drift', async () => {
    await expect(
      uploadLargeFileAttachmentToEvent(
        { eventId: 'immutable-event', idKind: 'restImmutableEntryId', source: source(3 * 1024 * 1024) },
        {
          readEvent: vi.fn(async () => event('different-id')),
          listAttachments: vi.fn(),
          createUploadSession: vi.fn(),
          readAttachmentBytes: vi.fn(),
        }
      )
    ).rejects.toBeInstanceOf(EventAttachmentUploadError);
  });
});
