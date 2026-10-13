import { describe, expect, it } from 'vitest';
import {
  attachDrivePdfToMailDraft,
  DrivePdfDraftError,
  MAX_DRAFT_PDF_BYTES,
  type DrivePdfDraftDependencies,
  type DrivePdfDraftInput,
} from './drive-pdf-mail-attachment.js';

const input: DrivePdfDraftInput = {
  driveId: 'drive-test',
  driveItemId: 'item-test',
  messageId: 'immutable-draft-test',
  expectedParentFolderId: 'draft-folder',
  expectedChangeKey: 'CK-before',
  expectedSourceETag: '"etag-original"',
};

const defaultPdf = Buffer.concat([Buffer.from('%PDF-1.7\n'), Buffer.alloc(349_760, 0x42)]);
const oldLogo = Buffer.from('synthetic inline PNG');

interface FixtureOptions {
  pdf?: Buffer;
  mimeType?: string;
  name?: string;
  sourceSize?: number;
  sourceETag?: string;
  parentChangeKey?: string;
  existingName?: string;
  nextLink?: boolean;
  mutateSourceBeforePost?: boolean;
  mutatePriorAfterPost?: boolean;
  mutateAudienceAfterPost?: boolean;
  mutateDraftDuringFinalRead?: boolean;
  mutateNewBytes?: boolean;
  throwPost?: boolean;
  badAcknowledgement?: boolean;
  badStatus?: boolean;
}

function fixture(options: FixtureOptions = {}) {
  const pdf = options.pdf ?? defaultPdf;
  const fileName = options.name ?? 'attendance.pdf';
  const oldName = options.existingName ?? 'logo.png';
  let applied = false;
  let posts = 0;
  let sourceReads = 0;
  let draftReads = 0;
  let actorReads = 0;
  let listReads = 0;
  const source = () => {
    sourceReads += 1;
    return {
      id: input.driveItemId,
      name: fileName,
      size: options.sourceSize ?? pdf.length,
      eTag:
        options.mutateSourceBeforePost && sourceReads > 1
          ? '"etag-mutated"'
          : (options.sourceETag ?? input.expectedSourceETag),
      file: { mimeType: options.mimeType ?? 'application/pdf' },
      parentReference: { driveId: input.driveId },
    };
  };
  const draft = () => {
    draftReads += 1;
    return {
      id: input.messageId,
      isDraft: true,
      changeKey: applied
        ? options.mutateDraftDuringFinalRead && draftReads > 3
          ? 'CK-third'
          : 'CK-after'
        : (options.parentChangeKey ?? input.expectedChangeKey),
      parentFolderId: input.expectedParentFolderId,
      subject: 'Synthetic QA draft only',
      body: { contentType: 'html', content: '<p>QA with CID logo</p><img src="cid:logo">' },
      toRecipients: applied && options.mutateAudienceAfterPost
        ? [{ emailAddress: { address: 'unexpected@example.test' } }]
        : [],
      ccRecipients: [],
      bccRecipients: [],
      replyTo: [],
      from: { emailAddress: { address: 'sender@example.test' } },
      sender: { emailAddress: { address: 'sender@example.test' } },
      conversationId: 'thread-test',
      internetMessageId: '<original@example.test>',
    };
  };
  const prior = {
    id: 'logo-child',
    name: oldName,
    contentType: 'image/png',
    isInline: true,
    contentId: 'logo',
    size: oldLogo.length,
  };
  const appended = {
    id: 'new-file-child',
    name: fileName,
    contentType: 'application/pdf',
    isInline: false,
    contentId: null,
    size: pdf.length + 300,
  };
  const listAttachments = () => {
    listReads += 1;
    const value = [
      { ...prior, ...(applied && options.mutatePriorAfterPost ? { name: 'logo-broken.png' } : {}) },
      ...(applied ? [appended] : []),
    ];
    return {
      value,
      ...(options.nextLink && !applied ? { '@odata.nextLink': 'https://next.invalid' } : {}),
    };
  };
  const deps: DrivePdfDraftDependencies = {
    readActor: async () => {
      actorReads += 1;
      return { id: 'actor-guid' };
    },
    readSource: async () => source(),
    readDraft: async () => draft(),
    listAttachments: async () => listAttachments(),
    readSourceBytes: async () => Buffer.from(pdf),
    readAttachmentBytes: async (id) => {
      if (id === prior.id) return options.mutatePriorAfterPost && applied
        ? Buffer.from('mutated inline PNG') : Buffer.from(oldLogo);
      if (id === appended.id && applied) return options.mutateNewBytes
        ? Buffer.from('%PDF-evil') : Buffer.from(pdf);
      throw new Error('unexpected attachment read');
    },
    appendFile: async (file) => {
      posts += 1;
      expect(file.name).toBe(fileName);
      expect(Buffer.from(file.contentBytes, 'base64')).toEqual(pdf);
      applied = true;
      if (options.throwPost) throw new Error('connection lost after possible Graph success');
      if (options.badAcknowledgement) return {
        status: 201,
        resource: { id: '', name: fileName, contentType: 'application/pdf', isInline: false },
      };
      return {
        status: options.badStatus ? 202 : 201,
        resource: {
          id: appended.id,
          name: fileName,
          contentType: 'application/pdf',
          isInline: false,
        },
      };
    },
  };
  return {
    deps,
    get posts() { return posts; },
    get applied() { return applied; },
    get reads() { return { sourceReads, draftReads, actorReads, listReads }; },
  };
}

async function errorCode(action: () => Promise<unknown>): Promise<string> {
  try {
    await action();
    throw new Error('expected rejection');
  } catch (error) {
    expect(error).toBeInstanceOf(DrivePdfDraftError);
    return (error as DrivePdfDraftError).code;
  }
}

describe('server-side OneDrive PDF to unsent Outlook draft', () => {
  it('attaches a real PDF and independently proves old logo, exact bytes and no send', async () => {
    const f = fixture();
    const result = await attachDrivePdfToMailDraft(input, f.deps);
    expect(result).toMatchObject({
      status: 'VERIFIED_DRAFT_ATTACHMENT',
      messageId: input.messageId,
      driveItemId: input.driveItemId,
      attachmentId: 'new-file-child',
      filename: 'attendance.pdf',
      decodedBytes: defaultPdf.length,
      retainedAttachments: 1,
      isDraft: true,
      sent: false,
    });
    expect(result.sha256).toMatch(/^[a-f0-9]{64}$/);
    expect(result.attemptId).toBeTruthy();
    expect(f.posts).toBe(1);
    expect(f.reads.sourceReads).toBe(3);
    expect(f.reads.draftReads).toBe(4);
  });

  it('accepts the hard 2,000,000-byte boundary', async () => {
    const f = fixture({ pdf: Buffer.concat([Buffer.from('%PDF-'), Buffer.alloc(MAX_DRAFT_PDF_BYTES - 5)]) });
    const result = await attachDrivePdfToMailDraft(input, f.deps);
    expect(result.decodedBytes).toBe(MAX_DRAFT_PDF_BYTES);
    expect(f.posts).toBe(1);
  });

  it('rejects dot-segment identities before any external read', async () => {
    for (const suspicious of ['.', '..']) {
      const f = fixture();
      const changed = { ...input, messageId: suspicious };
      expect(await errorCode(() => attachDrivePdfToMailDraft(changed, f.deps))).toBe(
        'INVALID_MESSAGEID'
      );
      expect(f.posts).toBe(0);
      expect(f.reads.sourceReads).toBe(0);
    }
  });

  it('rejects one byte beyond the limit without a write', async () => {
    const f = fixture({ sourceSize: MAX_DRAFT_PDF_BYTES + 1 });
    expect(await errorCode(() => attachDrivePdfToMailDraft(input, f.deps))).toBe('SOURCE_SIZE_CLASS');
    expect(f.posts).toBe(0);
  });

  it('rejects false PDF MIME and corrupt PDF signature before writing', async () => {
    const a = fixture({ mimeType: 'application/octet-stream' });
    expect(await errorCode(() => attachDrivePdfToMailDraft(input, a.deps))).toBe('SOURCE_MIME_NOT_PDF');
    expect(a.posts).toBe(0);
    const b = fixture({ pdf: Buffer.from('NOT-PDF-test-payload') });
    expect(await errorCode(() => attachDrivePdfToMailDraft(input, b.deps))).toBe(
      'SOURCE_PDF_HEADER_INVALID'
    );
    expect(b.posts).toBe(0);
  });

  it('rejects duplicate filename, even when an existing resource is inline', async () => {
    const f = fixture({ existingName: 'attendance.PDF' });
    expect(await errorCode(() => attachDrivePdfToMailDraft(input, f.deps))).toBe(
      'ATTACHMENT_NAME_ALREADY_EXISTS'
    );
    expect(f.posts).toBe(0);
  });

  it('rejects changed source revision or draft changeKey without writing', async () => {
    const a = fixture({ sourceETag: '"stale"' });
    expect(await errorCode(() => attachDrivePdfToMailDraft(input, a.deps))).toBe('SOURCE_ETAG_DRIFT');
    expect(a.posts).toBe(0);
    const b = fixture({ parentChangeKey: 'CK-elsewhere' });
    expect(await errorCode(() => attachDrivePdfToMailDraft(input, b.deps))).toBe('DRAFT_REVISION_DRIFT');
    expect(b.posts).toBe(0);
  });

  it('refuses incomplete paginated attachment inventory before dispatch', async () => {
    const f = fixture({ nextLink: true });
    expect(await errorCode(() => attachDrivePdfToMailDraft(input, f.deps))).toBe(
      'ATTACHMENT_PAGE_INCOMPLETE'
    );
    expect(f.posts).toBe(0);
  });

  it('detects OneDrive source mutation during preflight without dispatch', async () => {
    const f = fixture({ mutateSourceBeforePost: true });
    expect(await errorCode(() => attachDrivePdfToMailDraft(input, f.deps))).toBe('SOURCE_ETAG_DRIFT');
    expect(f.posts).toBe(0);
  });

  it('treats uncertain POST as possible effect and never retries', async () => {
    const f = fixture({ throwPost: true });
    let error: DrivePdfDraftError | undefined;
    try {
      await attachDrivePdfToMailDraft(input, f.deps);
    } catch (e) {
      error = e as DrivePdfDraftError;
    }
    expect(error?.code).toBe('POST_OUTCOME_UNKNOWN_NO_REPLAY');
    expect(error?.possibleWrite).toBe(true);
    expect(error?.attemptId).toBeTruthy();
    expect(f.posts).toBe(1);
  });

  it('does not claim success for missing provider acknowledgement or non-201', async () => {
    for (const options of [{ badAcknowledgement: true }, { badStatus: true }]) {
      const f = fixture(options);
      expect(await errorCode(() => attachDrivePdfToMailDraft(input, f.deps))).toBe(
        'POST_APPLIED_UNVERIFIED_NO_REPLAY'
      );
      expect(f.posts).toBe(1);
    }
  });

  it('rejects changed inline signature or recipient audience after POST', async () => {
    for (const options of [{ mutatePriorAfterPost: true }, { mutateAudienceAfterPost: true }]) {
      const f = fixture(options);
      expect(await errorCode(() => attachDrivePdfToMailDraft(input, f.deps))).toBe(
        'POST_APPLIED_UNVERIFIED_NO_REPLAY'
      );
      expect(f.posts).toBe(1);
    }
  });

  it('does not claim success for altered new attachment bytes', async () => {
    const f = fixture({ mutateNewBytes: true });
    expect(await errorCode(() => attachDrivePdfToMailDraft(input, f.deps))).toBe(
      'POST_APPLIED_UNVERIFIED_NO_REPLAY'
    );
    expect(f.posts).toBe(1);
  });

  it('rejects concurrent edit during final independent draft readback', async () => {
    const f = fixture({ mutateDraftDuringFinalRead: true });
    expect(await errorCode(() => attachDrivePdfToMailDraft(input, f.deps))).toBe(
      'POST_APPLIED_UNVERIFIED_NO_REPLAY'
    );
    expect(f.posts).toBe(1);
  });
});
