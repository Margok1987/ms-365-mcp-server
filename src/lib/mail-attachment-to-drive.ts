import { createHash, randomUUID } from 'crypto';
import { validateDriveUploadFileName } from './drive-large-upload.js';

export const MAX_MAIL_TO_DRIVE_BYTES = 30 * 1024 * 1024;
const MAX_METADATA_OVERHEAD = 64 * 1024;

type Json = Record<string, unknown>;

export interface MailAttachmentToDriveInput {
  messageId: string;
  attachmentId: string;
  expectedMessageChangeKey: string;
  expectedAttachmentName: string;
  driveId: string;
  parentItemId: string;
  expectedParentETag: string;
}

export interface MailAttachmentToDriveDeps {
  readActor(): Promise<Json>;
  readMessage(): Promise<Json>;
  readAttachment(): Promise<Json>;
  readAttachmentBytes(maxBytes: number): Promise<Buffer>;
  readParent(): Promise<Json>;
  readDestination(filename: string): Promise<Json | null>;
  createUploadSession(filename: string): Promise<{ uploadUrl: string }>;
  uploadSession(
    uploadUrl: string,
    bytes: Buffer
  ): Promise<{ driveItem: Json; chunksUploaded: number }>;
  readDriveBytes(itemId: string, maxBytes: number): Promise<Buffer>;
}

export interface MailAttachmentToDriveResult {
  status: 'VERIFIED_DRIVE_FILE';
  attemptId: string;
  messageId: string;
  attachmentId: string;
  driveId: string;
  parentItemId: string;
  driveItemId: string;
  filename: string;
  originalBytes: number;
  sha256: string;
  chunksUploaded: number;
  originalMailUnchanged: true;
  originalAttachmentRetained: true;
}

export class MailAttachmentToDriveError extends Error {
  constructor(
    public readonly code: string,
    public readonly possibleWrite = false,
    public readonly attemptId?: string
  ) {
    super(code);
    this.name = 'MailAttachmentToDriveError';
  }
}

function requireFact(condition: unknown, code: string): asserts condition {
  if (!condition) throw new MailAttachmentToDriveError(code);
}

function requireId(value: unknown, field: string): asserts value is string {
  requireFact(
    typeof value === 'string' &&
      value.length > 0 &&
      value.length <= 2048 &&
      value === value.trim() &&
      value !== '.' &&
      value !== '..' &&
      !/[\x00-\x20\x7f?#]/.test(value),
    'INVALID_' + field.toUpperCase()
  );
}

function record(value: unknown, code: string): asserts value is Json {
  requireFact(value !== null && typeof value === 'object' && !Array.isArray(value), code);
}

function sha256(bytes: Buffer): string {
  return createHash('sha256').update(bytes).digest('hex');
}

function checkMessage(current: Json, input: MailAttachmentToDriveInput, before?: Json): void {
  requireFact(current.id === input.messageId, 'SOURCE_MESSAGE_ID_DRIFT');
  requireFact(
    typeof current.changeKey === 'string' && !!current.changeKey,
    'SOURCE_MESSAGE_CHANGEKEY_MISSING'
  );
  requireFact(
    current.changeKey === input.expectedMessageChangeKey,
    'SOURCE_MESSAGE_CHANGEKEY_DRIFT'
  );
  if (before) {
    // A file export may not change or mark an Outlook message as read.
    for (const key of [
      'id',
      'changeKey',
      'parentFolderId',
      'subject',
      'isDraft',
      'isRead',
      'body',
      'toRecipients',
      'ccRecipients',
      'bccRecipients',
      'conversationId',
      'hasAttachments',
    ]) {
      requireFact(
        JSON.stringify(current[key] ?? null) === JSON.stringify(before[key] ?? null),
        'SOURCE_MESSAGE_MUTATED_' + key
      );
    }
  }
}

function checkAttachment(attachment: Json, input: MailAttachmentToDriveInput): string {
  requireFact(attachment.id === input.attachmentId, 'SOURCE_ATTACHMENT_ID_DRIFT');
  requireFact(
    attachment['@odata.type'] === '#microsoft.graph.fileAttachment',
    'NOT_ORIGINAL_FILE_ATTACHMENT'
  );
  requireFact(attachment.isInline === false, 'INLINE_CID_NOT_EXPORTABLE');
  requireFact(
    typeof attachment.contentType === 'string' &&
      attachment.contentType.length > 0 &&
      attachment.contentType.length <= 200,
    'SOURCE_CONTENT_TYPE_INVALID'
  );
  requireFact(
    typeof attachment.size === 'number' &&
      Number.isSafeInteger(attachment.size) &&
      attachment.size >= 1 &&
      attachment.size <= MAX_MAIL_TO_DRIVE_BYTES + MAX_METADATA_OVERHEAD,
    'SOURCE_SIZE_CLASS'
  );
  requireFact(
    attachment.name === input.expectedAttachmentName,
    'SOURCE_ATTACHMENT_NAME_DRIFT'
  );
  requireFact(
    typeof attachment.name === 'string' &&
      attachment.name.length <= 255 &&
      attachment.name === attachment.name.trim(),
    'SOURCE_FILENAME_INVALID'
  );
  const filename = attachment.name as string;
  try {
    return validateDriveUploadFileName(filename);
  } catch {
    throw new MailAttachmentToDriveError('SOURCE_FILENAME_INVALID');
  }
}

function checkParent(parent: Json, input: MailAttachmentToDriveInput, before = false): void {
  requireFact(parent.id === input.parentItemId, 'DESTINATION_FOLDER_ID_DRIFT');
  record(parent.folder, 'DESTINATION_NOT_A_FOLDER');
  if (!before) {
    requireFact(parent.eTag === input.expectedParentETag, 'DESTINATION_FOLDER_ETAG_DRIFT');
  }
  const reference = parent.parentReference;
  if (reference !== undefined && reference !== null) {
    record(reference, 'DESTINATION_PARENT_REFERENCE_INVALID');
    if (reference.driveId !== undefined) {
      requireFact(reference.driveId === input.driveId, 'DESTINATION_DRIVE_DRIFT');
    }
  }
}

function checkDestination(
  destination: Json | null,
  input: MailAttachmentToDriveInput,
  filename: string,
  size: number,
  completedId: string
): Json {
  requireFact(destination !== null, 'DESTINATION_READBACK_MISSING');
  record(destination, 'DESTINATION_READBACK_INVALID');
  requireFact(
    destination.id === completedId &&
      destination.name === filename &&
      destination.size === size,
    'DESTINATION_ID_NAME_SIZE_MISMATCH'
  );
  record(destination.file, 'DESTINATION_NOT_A_FILE');
  record(destination.parentReference, 'DESTINATION_PARENT_REFERENCE_MISSING');
  requireFact(
    destination.parentReference.id === input.parentItemId &&
      destination.parentReference.driveId === input.driveId,
    'DESTINATION_PARENT_DRIVE_MISMATCH'
  );
  return destination;
}

function sameAttachment(a: Json, b: Json): boolean {
  for (const key of ['id', 'name', 'contentType', 'size', 'isInline', '@odata.type']) {
    if (JSON.stringify(a[key] ?? null) !== JSON.stringify(b[key] ?? null)) return false;
  }
  return true;
}

/**
 * Server-side source export: exactly one real Outlook fileAttachment -> new Drive item.
 * Never overwrites, deletes, forwards, sends, or transmits original bytes to the model.
 * An uncertain final upload may have created a file: no blind replay.
 */
export async function saveMailAttachmentToDrive(
  input: MailAttachmentToDriveInput,
  deps: MailAttachmentToDriveDeps
): Promise<MailAttachmentToDriveResult> {
  for (const key of ['messageId', 'attachmentId', 'driveId', 'parentItemId'] as const) {
    requireId(input[key], key);
  }
  for (const key of ['expectedMessageChangeKey', 'expectedParentETag'] as const) {
    requireFact(
      typeof input[key] === 'string' && input[key].length > 0 && !/[\r\n]/.test(input[key]),
      'INVALID_' + key.toUpperCase()
    );
  }
  requireFact(
    typeof input.expectedAttachmentName === 'string' &&
      input.expectedAttachmentName.length > 0,
    'INVALID_EXPECTED_ATTACHMENT_NAME'
  );
  const attemptId = randomUUID();
  const [actor, originalMessage, originalAttachment, folder, preexisting] = await Promise.all([
    deps.readActor(),
    deps.readMessage(),
    deps.readAttachment(),
    deps.readParent(),
    deps.readDestination(input.expectedAttachmentName),
  ]);
  requireId(actor.id, 'actor_id');
  checkMessage(originalMessage, input);
  const filename = checkAttachment(originalAttachment, input);
  checkParent(folder, input);
  requireFact(preexisting === null, 'DESTINATION_EXISTS_NO_OVERWRITE');

  const bytes = await deps.readAttachmentBytes(MAX_MAIL_TO_DRIVE_BYTES);
  requireFact(Buffer.isBuffer(bytes), 'SOURCE_BYTES_INVALID');
  requireFact(
    bytes.length > 0 && bytes.length <= MAX_MAIL_TO_DRIVE_BYTES,
    'SOURCE_BYTES_OUT_OF_BOUNDS'
  );
  const originalMetadataSize = originalAttachment.size;
  requireFact(
    typeof originalMetadataSize === 'number' &&
      originalMetadataSize >= bytes.length &&
      originalMetadataSize - bytes.length <= MAX_METADATA_OVERHEAD,
    'SOURCE_METADATA_SIZE_MISMATCH'
  );
  const expectedHash = sha256(bytes);
  const [actor2, message2, attachment2, parent2, exists2] = await Promise.all([
    deps.readActor(),
    deps.readMessage(),
    deps.readAttachment(),
    deps.readParent(),
    deps.readDestination(filename),
  ]);
  requireFact(actor2.id === actor.id, 'ACTOR_ID_DRIFT');
  checkMessage(message2, input, originalMessage);
  checkAttachment(attachment2, input);
  requireFact(sameAttachment(attachment2, originalAttachment), 'SOURCE_ATTACHMENT_DRIFT');
  checkParent(parent2, input);
  requireFact(exists2 === null, 'DESTINATION_RACE_NO_OVERWRITE');

  // Graph upload sessions with conflictBehavior=fail create a new DriveItem only.
  // Once an upload has begun, a failed/uncertain final chunk is never replayed.
  let committed: { driveItem: Json; chunksUploaded: number };
  try {
    const session = await deps.createUploadSession(filename);
    requireFact(
      typeof session.uploadUrl === 'string' && session.uploadUrl.startsWith('https://'),
      'UPLOAD_SESSION_MISSING'
    );
    committed = await deps.uploadSession(session.uploadUrl, bytes);
  } catch {
    throw new MailAttachmentToDriveError(
      'UPLOAD_OUTCOME_UNKNOWN_NO_REPLAY',
      true,
      attemptId
    );
  }

  try {
    record(committed.driveItem, 'UPLOAD_COMPLETION_INVALID');
    requireId(committed.driveItem.id, 'new_drive_item_id');
    requireFact(
      committed.driveItem.name === filename &&
        committed.driveItem.size === bytes.length,
      'UPLOAD_COMPLETION_DRIFT'
    );
    const newId = committed.driveItem.id;
    const [saved, afterMessage, afterAttachment, afterActor, parentAfter] =
      await Promise.all([
        deps.readDestination(filename),
        deps.readMessage(),
        deps.readAttachment(),
        deps.readActor(),
        deps.readParent(),
      ]);
    checkDestination(saved, input, filename, bytes.length, newId);
    checkMessage(afterMessage, input, originalMessage);
    checkAttachment(afterAttachment, input);
    requireFact(
      sameAttachment(afterAttachment, originalAttachment),
      'ORIGINAL_ATTACHMENT_METADATA_DRIFT'
    );
    requireFact(afterActor.id === actor.id, 'ACTOR_ID_DRIFT');
    checkParent(parentAfter, input, true);

    const [destBytes, retainedBytes] = await Promise.all([
      deps.readDriveBytes(newId, MAX_MAIL_TO_DRIVE_BYTES),
      deps.readAttachmentBytes(MAX_MAIL_TO_DRIVE_BYTES),
    ]);
    requireFact(
      Buffer.isBuffer(destBytes) &&
        destBytes.length === bytes.length &&
        sha256(destBytes) === expectedHash,
      'DESTINATION_BYTE_HASH_MISMATCH'
    );
    requireFact(
      Buffer.isBuffer(retainedBytes) &&
        retainedBytes.length === bytes.length &&
        sha256(retainedBytes) === expectedHash,
      'ORIGINAL_ATTACHMENT_BYTE_HASH_DRIFT'
    );

    const [finalSaved, finalMessage, finalAttachment] = await Promise.all([
      deps.readDestination(filename),
      deps.readMessage(),
      deps.readAttachment(),
    ]);
    checkDestination(finalSaved, input, filename, bytes.length, newId);
    checkMessage(finalMessage, input, originalMessage);
    checkAttachment(finalAttachment, input);
    requireFact(
      sameAttachment(finalAttachment, originalAttachment),
      'ORIGINAL_ATTACHMENT_FINAL_DRIFT'
    );
    return {
      status: 'VERIFIED_DRIVE_FILE',
      attemptId,
      messageId: input.messageId,
      attachmentId: input.attachmentId,
      driveId: input.driveId,
      parentItemId: input.parentItemId,
      driveItemId: newId,
      filename,
      originalBytes: bytes.length,
      sha256: expectedHash,
      chunksUploaded: committed.chunksUploaded,
      originalMailUnchanged: true,
      originalAttachmentRetained: true,
    };
  } catch {
    throw new MailAttachmentToDriveError(
      'UPLOAD_APPLIED_UNVERIFIED_NO_REPLAY',
      true,
      attemptId
    );
  }
}
