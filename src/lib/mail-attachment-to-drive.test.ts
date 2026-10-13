import { describe, expect, it } from 'vitest';
import {
  saveMailAttachmentToDrive,
  MailAttachmentToDriveError,
  MAX_MAIL_TO_DRIVE_BYTES,
  type MailAttachmentToDriveDeps,
  type MailAttachmentToDriveInput,
} from './mail-attachment-to-drive.js';

const input: MailAttachmentToDriveInput = {
  messageId: 'immutable-message-1',
  attachmentId: 'immutable-attachment-1',
  driveId: 'drive-1',
  parentItemId: 'folder-1',
  expectedMessageChangeKey: 'CK-before',
  expectedParentETag: '"etag-folder"',
  expectedAttachmentName: 'Bücher Liste 2026.pdf',
};
const originalPdf = Buffer.concat([Buffer.from('%PDF-1.7\n'), Buffer.alloc(1_172_933, 0x48)]);

interface Case {
  bytes?: Buffer;
  attachmentType?: string;
  name?: string;
  metadataSize?: number;
  inline?: boolean;
  preexisting?: boolean;
  newConflictBeforeWrite?: boolean;
  sourceChangeKey?: string;
  sourceChangeAfterUpload?: boolean;
  attachmentChangeBeforeUpload?: boolean;
  attachmentChangeAfterUpload?: boolean;
  folderETag?: string;
  folderIsFile?: boolean;
  parentDrive?: string;
  actorChanges?: boolean;
  ambiguousFinal?: boolean;
  sessionError?: boolean;
  uploadIncorrectSize?: boolean;
  uploadNoId?: boolean;
  postDifferentBytes?: boolean;
  sourceDifferentBytesAfter?: boolean;
  postWrongParent?: boolean;
  malformedFilename?: boolean;
}

function fixture(options: Case = {}) {
  const bytes = options.bytes ?? originalPdf;
  const name = options.name ?? input.expectedAttachmentName;
  let committed = false;
  let writes = 0;
  let sessions = 0;
  let sourceReads = 0;
  let actorReads = 0;
  let messageReads = 0;
  let destReads = 0;
  const newId = 'created-drive-item-1';
  const attachment = () => {
    sourceReads += 1;
    const changing =
      options.attachmentChangeBeforeUpload && sourceReads > 1 ||
      options.attachmentChangeAfterUpload && committed;
    return {
      id: input.attachmentId,
      '@odata.type': options.attachmentType ?? '#microsoft.graph.fileAttachment',
      name: changing ? 'renamed.pdf' : name,
      size: options.metadataSize ?? bytes.length + 400,
      contentType: 'application/pdf',
      isInline: options.inline ?? false,
    };
  };
  const message = () => {
    messageReads += 1;
    return {
      id: input.messageId,
      changeKey: options.sourceChangeKey ?? (
        committed && options.sourceChangeAfterUpload ? 'CK-other' : input.expectedMessageChangeKey
      ),
      parentFolderId: 'inbox-1',
      subject: 'Workshop – Rückmeldung',
      isDraft: false,
      isRead: true,
      body: { contentType: 'html', content: '<p>Beleg anbei</p>' },
      toRecipients: [{ emailAddress: { address: 'person@example.test' } }],
      ccRecipients: [],
      bccRecipients: [],
      conversationId: 'conv-1',
      hasAttachments: true,
    };
  };
  const parent = () => ({
    id: input.parentItemId,
    eTag: options.folderETag ?? (committed ? '"folder-version-2"' : input.expectedParentETag),
    folder: options.folderIsFile ? undefined : { childCount: committed ? 1 : 0 },
    parentReference: { driveId: options.parentDrive ?? input.driveId },
  });
  const destination = () => ({
    id: newId,
    name,
    size: bytes.length,
    file: { mimeType: 'application/pdf' },
    parentReference: {
      driveId: options.postWrongParent ? 'other-drive' : input.driveId,
      id: options.postWrongParent ? 'other-folder' : input.parentItemId,
    },
  });
  const deps: MailAttachmentToDriveDeps = {
    readActor: async () => {
      actorReads += 1;
      return { id: options.actorChanges && actorReads > 1 ? 'actor-other' : 'actor-1' };
    },
    readMessage: async () => message(),
    readAttachment: async () => attachment(),
    readAttachmentBytes: async () =>
      committed && options.sourceDifferentBytesAfter
        ? Buffer.from('wrong retained original bytes')
        : Buffer.from(bytes),
    readParent: async () => parent(),
    readDestination: async () => {
      destReads += 1;
      const exists =
        options.preexisting ||
        (options.newConflictBeforeWrite && destReads >= 2) ||
        committed;
      return exists ? destination() : null;
    },
    createUploadSession: async (filename) => {
      sessions += 1;
      expect(filename).toBe(name);
      if (options.sessionError) throw new Error('Graph refused to create upload session');
      return { uploadUrl: 'https://upload.example.test/session' };
    },
    uploadSession: async (_url, content) => {
      writes += 1;
      expect(content).toEqual(bytes);
      committed = true;
      if (options.ambiguousFinal) throw new Error('final byte-range reply lost');
      return {
        driveItem: {
          id: options.uploadNoId ? '' : newId,
          name,
          size: options.uploadIncorrectSize ? bytes.length + 1 : bytes.length,
        },
        chunksUploaded: Math.ceil(bytes.length / (5 * 1024 * 1024)),
      };
    },
    readDriveBytes: async () =>
      options.postDifferentBytes ? Buffer.from('corrupted destination') : Buffer.from(bytes),
  };
  return {
    deps,
    get writes() { return writes; },
    get sessions() { return sessions; },
    get sourceReads() { return sourceReads; },
    get committed() { return committed; },
    get messageReads() { return messageReads; },
  };
}

async function getError(f: () => Promise<unknown>): Promise<MailAttachmentToDriveError> {
  try {
    await f();
    throw new Error('expected MailAttachmentToDriveError');
  } catch (error) {
    expect(error).toBeInstanceOf(MailAttachmentToDriveError);
    return error as MailAttachmentToDriveError;
  }
}

describe('Outlook attachment to new OneDrive or SharePoint file', () => {
  it('creates one real PDF with spaces and umlauts and verifies all original bytes', async () => {
    const f = fixture();
    const result = await saveMailAttachmentToDrive(input, f.deps);
    expect(result).toMatchObject({
      status: 'VERIFIED_DRIVE_FILE',
      messageId: input.messageId,
      attachmentId: input.attachmentId,
      driveItemId: 'created-drive-item-1',
      filename: 'Bücher Liste 2026.pdf',
      originalBytes: originalPdf.length,
      originalMailUnchanged: true,
      originalAttachmentRetained: true,
    });
    expect(result.sha256).toMatch(/^[a-f0-9]{64}$/);
    expect(result.attemptId).toBeTruthy();
    expect(f.writes).toBe(1);
    expect(f.sessions).toBe(1);
    expect(f.messageReads).toBe(4);
  });

  it('supports the full 30 MiB original-content class in bounded chunks', async () => {
    const bytes = Buffer.alloc(MAX_MAIL_TO_DRIVE_BYTES, 0x42);
    const f = fixture({ bytes });
    const result = await saveMailAttachmentToDrive(input, f.deps);
    expect(result.originalBytes).toBe(MAX_MAIL_TO_DRIVE_BYTES);
    expect(result.chunksUploaded).toBe(6);
    expect(f.writes).toBe(1);
  });

  it('refuses an actual stream one byte too large before write', async () => {
    const f = fixture({ bytes: Buffer.alloc(MAX_MAIL_TO_DRIVE_BYTES + 1) });
    const error = await getError(() => saveMailAttachmentToDrive(input, f.deps));
    expect(error.code).toBe('SOURCE_BYTES_OUT_OF_BOUNDS');
    expect(error.possibleWrite).toBe(false);
    expect(f.sessions).toBe(0);
  });

  it('rejects source metadata size inconsistent with the original bytes', async () => {
    for (const metadataSize of [originalPdf.length - 1, originalPdf.length + 65_537]) {
      const f = fixture({ metadataSize });
      const error = await getError(() => saveMailAttachmentToDrive(input, f.deps));
      expect(error.code).toBe('SOURCE_METADATA_SIZE_MISMATCH');
      expect(error.possibleWrite).toBe(false);
      expect(f.sessions).toBe(0);
    }
  });

  it('does not export references, item attachments or inline CID resources', async () => {
    for (const opts of [
      { attachmentType: '#microsoft.graph.referenceAttachment' },
      { attachmentType: '#microsoft.graph.itemAttachment' },
      { inline: true },
    ]) {
      const f = fixture(opts);
      const error = await getError(() => saveMailAttachmentToDrive(input, f.deps));
      expect(['NOT_ORIGINAL_FILE_ATTACHMENT', 'INLINE_CID_NOT_EXPORTABLE']).toContain(error.code);
      expect(f.writes).toBe(0);
    }
  });

  it('rejects a changed original file name or invalid OneDrive filename', async () => {
    const a = fixture({ name: 'other.pdf' });
    expect((await getError(() => saveMailAttachmentToDrive(input, a.deps))).code).toBe(
      'SOURCE_ATTACHMENT_NAME_DRIFT'
    );
    const b = fixture({ name: 'a/b.pdf' });
    const changed = { ...input, expectedAttachmentName: 'a/b.pdf' };
    expect((await getError(() => saveMailAttachmentToDrive(changed, b.deps))).code).toBe(
      'SOURCE_FILENAME_INVALID'
    );
    expect(a.writes + b.writes).toBe(0);
  });

  it('does not silently overwrite an existing OneDrive child', async () => {
    const f = fixture({ preexisting: true });
    expect((await getError(() => saveMailAttachmentToDrive(input, f.deps))).code).toBe(
      'DESTINATION_EXISTS_NO_OVERWRITE'
    );
    expect(f.sessions).toBe(0);
  });

  it('rejects a new name collision immediately before opening an upload session', async () => {
    const f = fixture({ newConflictBeforeWrite: true });
    expect((await getError(() => saveMailAttachmentToDrive(input, f.deps))).code).toBe(
      'DESTINATION_RACE_NO_OVERWRITE'
    );
    expect(f.sessions).toBe(0);
  });

  it('fails closed for missing folder, folder drift or stale eTag', async () => {
    for (const opts of [
      { folderIsFile: true },
      { parentDrive: 'other-drive' },
      { folderETag: 'stale' },
    ]) {
      const f = fixture(opts);
      const error = await getError(() => saveMailAttachmentToDrive(input, f.deps));
      expect([
        'DESTINATION_NOT_A_FOLDER', 'DESTINATION_DRIVE_DRIFT',
        'DESTINATION_FOLDER_ETAG_DRIFT',
      ]).toContain(error.code);
      expect(f.writes).toBe(0);
    }
  });

  it('stops on Outlook message changeKey, source name or actor change', async () => {
    for (const opts of [
      { sourceChangeKey: 'CK-other' },
      { attachmentChangeBeforeUpload: true },
      { actorChanges: true },
    ]) {
      const f = fixture(opts);
      const error = await getError(() => saveMailAttachmentToDrive(input, f.deps));
      expect([
        'SOURCE_MESSAGE_CHANGEKEY_DRIFT', 'SOURCE_ATTACHMENT_NAME_DRIFT',
        'ACTOR_ID_DRIFT',
      ]).toContain(error.code);
      expect(f.sessions).toBe(0);
    }
  });

  it('reports a server/session creation failure without claiming success', async () => {
    const f = fixture({ sessionError: true });
    const error = await getError(() => saveMailAttachmentToDrive(input, f.deps));
    expect(error.code).toBe('UPLOAD_OUTCOME_UNKNOWN_NO_REPLAY');
    expect(error.possibleWrite).toBe(true);
    expect(error.attemptId).toBeTruthy();
    expect(f.sessions).toBe(1);
    expect(f.writes).toBe(0);
  });

  it('does not retry an ambiguous final upload', async () => {
    const f = fixture({ ambiguousFinal: true });
    const error = await getError(() => saveMailAttachmentToDrive(input, f.deps));
    expect(error.code).toBe('UPLOAD_OUTCOME_UNKNOWN_NO_REPLAY');
    expect(error.possibleWrite).toBe(true);
    expect(error.attemptId).toBeTruthy();
    expect(f.writes).toBe(1);
  });

  it('does not claim success from an invalid completion ID or size', async () => {
    for (const opts of [{ uploadNoId: true }, { uploadIncorrectSize: true }]) {
      const f = fixture(opts);
      const error = await getError(() => saveMailAttachmentToDrive(input, f.deps));
      expect(error.code).toBe('UPLOAD_APPLIED_UNVERIFIED_NO_REPLAY');
      expect(error.possibleWrite).toBe(true);
      expect(f.writes).toBe(1);
    }
  });

  it('blocks a destination file with a mismatched byte hash or folder', async () => {
    for (const opts of [{ postDifferentBytes: true }, { postWrongParent: true }]) {
      const f = fixture(opts);
      const error = await getError(() => saveMailAttachmentToDrive(input, f.deps));
      expect(error.code).toBe('UPLOAD_APPLIED_UNVERIFIED_NO_REPLAY');
      expect(f.writes).toBe(1);
    }
  });

  it('detects post-write Outlook message edits and source attachment mutations', async () => {
    for (const opts of [
      { sourceChangeAfterUpload: true },
      { attachmentChangeAfterUpload: true },
      { sourceDifferentBytesAfter: true },
    ]) {
      const f = fixture(opts);
      const error = await getError(() => saveMailAttachmentToDrive(input, f.deps));
      expect(error.code).toBe('UPLOAD_APPLIED_UNVERIFIED_NO_REPLAY');
      expect(error.possibleWrite).toBe(true);
      expect(f.writes).toBe(1);
    }
  });

  it('rejects invalid resource IDs before Graph reads', async () => {
    for (const suspicious of ['.', '..', '', 'id?x=3', 'id\nX']) {
      const f = fixture();
      const error = await getError(() =>
        saveMailAttachmentToDrive({ ...input, messageId: suspicious }, f.deps)
      );
      expect(error.code).toBe('INVALID_MESSAGEID');
      expect(f.sourceReads).toBe(0);
      expect(f.sessions).toBe(0);
    }
  });
});
