import { createHash, randomUUID } from 'crypto';

export const MAX_DRAFT_PDF_BYTES = 2_000_000;
const MAX_PRIOR_ATTACHMENTS = 12;
const MAX_PRIOR_BYTES = 3_000_000;
type Json = Record<string, unknown>;

export interface DrivePdfDraftInput {
  driveId: string;
  driveItemId: string;
  messageId: string;
  expectedParentFolderId: string;
  expectedChangeKey: string;
  expectedSourceETag: string;
}

export interface DrivePdfDraftDependencies {
  readActor(): Promise<Json>;
  readSource(): Promise<Json>;
  readDraft(): Promise<Json>;
  listAttachments(): Promise<Json>;
  readSourceBytes(maxBytes: number, etag: string): Promise<Buffer>;
  readAttachmentBytes(attachmentId: string, maxBytes: number): Promise<Buffer>;
  appendFile(file: { name: string; contentBytes: string }): Promise<{
    status: number;
    resource: Json;
  }>;
}

export interface DrivePdfDraftResult {
  status: 'VERIFIED_DRAFT_ATTACHMENT';
  attemptId: string;
  messageId: string;
  driveItemId: string;
  attachmentId: string;
  filename: string;
  decodedBytes: number;
  sha256: string;
  retainedAttachments: number;
  isDraft: true;
  sent: false;
}

export class DrivePdfDraftError extends Error {
  constructor(
    public readonly code: string,
    public readonly possibleWrite = false,
    public readonly attemptId?: string
  ) {
    super(code);
    this.name = 'DrivePdfDraftError';
  }
}

function requireValue(condition: unknown, code: string): asserts condition {
  if (!condition) throw new DrivePdfDraftError(code);
}

function requireId(value: unknown, field: string): asserts value is string {
  requireValue(
    typeof value === 'string' &&
      value.length > 0 &&
      value.length <= 4096 &&
      value.trim() === value &&
      value !== '.' &&
      value !== '..' &&
      // Intentional exclusion of control bytes from opaque Graph identifiers.
      // eslint-disable-next-line no-control-regex
      !/[\x00-\x20\x7f?#]/.test(value),
    'INVALID_' + field.toUpperCase()
  );
}

function requireRecord(value: unknown, code: string): asserts value is Json {
  requireValue(value !== null && typeof value === 'object' && !Array.isArray(value), code);
}

function digest(bytes: Buffer): string {
  return createHash('sha256').update(bytes).digest('hex');
}

function sameJson(a: unknown, b: unknown): boolean {
  return JSON.stringify(a ?? null) === JSON.stringify(b ?? null);
}

const frozenFields = [
  'subject',
  'body',
  'toRecipients',
  'ccRecipients',
  'bccRecipients',
  'replyTo',
  'from',
  'sender',
  'conversationId',
  'internetMessageId',
] as const;

function checkDraft(draft: Json, input: DrivePdfDraftInput, before?: Json): void {
  requireValue(draft.id === input.messageId, 'DRAFT_ID_DRIFT');
  requireValue(draft.isDraft === true, 'NOT_UNSENT_DRAFT');
  requireValue(draft.parentFolderId === input.expectedParentFolderId, 'DRAFT_FOLDER_DRIFT');
  requireValue(typeof draft.changeKey === 'string' && !!draft.changeKey, 'DRAFT_REVISION_MISSING');
  requireValue(typeof draft.subject === 'string' && !!draft.subject.trim(), 'DRAFT_SUBJECT_MISSING');
  requireRecord(draft.body, 'DRAFT_BODY_MISSING');
  requireValue(
    typeof draft.body.content === 'string' &&
      (draft.body.contentType === 'html' || draft.body.contentType === 'text'),
    'DRAFT_BODY_INVALID'
  );
  for (const key of ['toRecipients', 'ccRecipients', 'bccRecipients'] as const) {
    requireValue(Array.isArray(draft[key]), 'DRAFT_AUDIENCE_INCOMPLETE');
  }
  if (!before) {
    requireValue(draft.changeKey === input.expectedChangeKey, 'DRAFT_REVISION_DRIFT');
  } else {
    requireValue(draft.changeKey !== before.changeKey, 'DRAFT_REVISION_UNCHANGED');
    for (const key of frozenFields) {
      requireValue(sameJson(draft[key], before[key]), 'DRAFT_CONTENT_DRIFT_' + key);
    }
  }
}

function sourceFile(source: Json, input: DrivePdfDraftInput): string {
  requireValue(source.id === input.driveItemId, 'SOURCE_ID_DRIFT');
  const parent = source.parentReference;
  requireRecord(parent, 'SOURCE_PARENT_MISSING');
  requireValue(parent.driveId === input.driveId, 'SOURCE_DRIVE_DRIFT');
  requireValue(source.eTag === input.expectedSourceETag, 'SOURCE_ETAG_DRIFT');
  requireValue(
    typeof source.size === 'number' &&
      Number.isInteger(source.size) &&
      source.size >= 15 &&
      source.size <= MAX_DRAFT_PDF_BYTES,
    'SOURCE_SIZE_CLASS'
  );
  const file = source.file;
  requireRecord(file, 'SOURCE_NOT_A_FILE');
  requireValue(file.mimeType === 'application/pdf', 'SOURCE_MIME_NOT_PDF');
  const name = source.name;
  requireValue(
    typeof name === 'string' &&
      name.length > 4 &&
      name.length <= 255 &&
      name.toLowerCase().endsWith('.pdf') &&
      // Graph drive item names may not contain path separators or control bytes.
      // eslint-disable-next-line no-control-regex
      !/[\\/\x00-\x1f\x7f]/.test(name),
    'SOURCE_FILENAME_INVALID'
  );
  return name as string;
}

interface PriorChild {
  id: string;
  name: string;
  contentType: string;
  isInline: boolean;
  contentId: string | null;
  size: number;
}

function children(response: Json): PriorChild[] {
  requireValue(!response['@odata.nextLink'], 'ATTACHMENT_PAGE_INCOMPLETE');
  requireValue(Array.isArray(response.value), 'ATTACHMENT_LIST_MISSING');
  requireValue(response.value.length <= MAX_PRIOR_ATTACHMENTS + 1, 'ATTACHMENT_LIMIT');
  const ids = new Set<string>();
  return response.value.map((value: unknown) => {
    requireRecord(value, 'ATTACHMENT_INVALID');
    requireId(value.id, 'attachment_id');
    requireValue(!ids.has(value.id), 'ATTACHMENT_ID_DUPLICATE');
    ids.add(value.id);
    requireValue(
      typeof value.name === 'string' &&
        typeof value.contentType === 'string' &&
        typeof value.isInline === 'boolean' &&
        typeof value.size === 'number' &&
        Number.isInteger(value.size) &&
        value.size >= 0 &&
        value.size <= MAX_PRIOR_BYTES + 2048,
      'ATTACHMENT_METADATA_INVALID'
    );
    requireValue(
      value.contentId === undefined ||
        value.contentId === null ||
        typeof value.contentId === 'string',
      'ATTACHMENT_CONTENT_ID_INVALID'
    );
    return {
      id: value.id,
      name: value.name,
      contentType: value.contentType,
      isInline: value.isInline,
      contentId: typeof value.contentId === 'string' ? value.contentId : null,
      size: value.size,
    };
  });
}

function fingerprint(child: PriorChild): string {
  return JSON.stringify(child);
}

function ensureBaseline(next: PriorChild[], baseline: PriorChild[], newId: string): void {
  requireValue(next.length === baseline.length + 1, 'ATTACHMENT_SET_DRIFT');
  const existing = new Map(next.map((child) => [child.id, child]));
  for (const child of baseline) {
    requireValue(existing.has(child.id), 'PRIOR_ATTACHMENT_LOST');
    requireValue(
      fingerprint(existing.get(child.id) as PriorChild) === fingerprint(child),
      'PRIOR_ATTACHMENT_CHANGED'
    );
  }
  requireValue(next.filter((child) => child.id === newId).length === 1, 'NEW_ATTACHMENT_MISSING');
}

async function readExisting(
  deps: DrivePdfDraftDependencies,
  items: PriorChild[]
): Promise<Map<string, string>> {
  const hashes = new Map<string, string>();
  for (const child of items) {
    const bytes = await deps.readAttachmentBytes(child.id, MAX_PRIOR_BYTES);
    requireValue(Buffer.isBuffer(bytes) && bytes.length <= MAX_PRIOR_BYTES, 'PRIOR_BYTES_INVALID');
    hashes.set(child.id, digest(bytes));
  }
  return hashes;
}

/**
 * One first-class operation: server-side Drive content -> real Outlook fileAttachment.
 * Neither original bytes nor Graph tokens are returned in the model-visible result.
 * No send, no destructive compensation, no retry after possible POST.
 */
export async function attachDrivePdfToMailDraft(
  input: DrivePdfDraftInput,
  deps: DrivePdfDraftDependencies
): Promise<DrivePdfDraftResult> {
  for (const [key, value] of Object.entries(input)) requireId(value, key);
  const attemptId = randomUUID();
  const actor = await deps.readActor();
  requireId(actor.id, 'actor');
  const [source, draft, initialList] = await Promise.all([
    deps.readSource(),
    deps.readDraft(),
    deps.listAttachments(),
  ]);
  const name = sourceFile(source, input);
  checkDraft(draft, input);
  const initial = children(initialList);
  requireValue(initial.length <= MAX_PRIOR_ATTACHMENTS, 'ATTACHMENT_LIMIT');
  requireValue(
    !initial.some((child) => child.name.toLowerCase() === name.toLowerCase()),
    'ATTACHMENT_NAME_ALREADY_EXISTS'
  );
  const priorHashes = await readExisting(deps, initial);
  const bytes = await deps.readSourceBytes(MAX_DRAFT_PDF_BYTES, input.expectedSourceETag);
  requireValue(Buffer.isBuffer(bytes), 'SOURCE_BYTES_INVALID');
  requireValue(bytes.length === source.size, 'SOURCE_BYTES_LENGTH_DRIFT');
  requireValue(bytes.subarray(0, 5).toString('ascii') === '%PDF-', 'SOURCE_PDF_HEADER_INVALID');
  const hash = digest(bytes);

  // Pre-dispatch rechecks: source version and draft revision may change during file reads.
  const [againSource, againDraft, againList, againActor] = await Promise.all([
    deps.readSource(), deps.readDraft(), deps.listAttachments(), deps.readActor(),
  ]);
  sourceFile(againSource, input);
  checkDraft(againDraft, input);
  requireValue(againActor.id === actor.id, 'ACTOR_ID_DRIFT');
  const prepost = children(againList);
  requireValue(
    prepost.length === initial.length &&
      initial.every((item, i) => fingerprint(item) === fingerprint(prepost[i])),
    'ATTACHMENT_PREPOST_DRIFT'
  );

  let acknowledgment: { status: number; resource: Json };
  try {
    acknowledgment = await deps.appendFile({ name, contentBytes: bytes.toString('base64') });
  } catch {
    throw new DrivePdfDraftError('POST_OUTCOME_UNKNOWN_NO_REPLAY', true, attemptId);
  }

  // Once a POST was attempted, every unverified result is a possible effect.
  try {
    requireValue(acknowledgment.status === 201, 'POST_NOT_ACKNOWLEDGED');
    const ack = acknowledgment.resource;
    requireRecord(ack, 'POST_ACK_INVALID');
    requireId(ack.id, 'new_attachment_id');
    requireValue(
      ack.name === name &&
        ack.contentType === 'application/pdf' &&
        ack.isInline === false,
      'POST_ACK_DRIFT'
    );
    const [afterDraft, afterList, afterSource, afterActor] = await Promise.all([
      deps.readDraft(), deps.listAttachments(), deps.readSource(), deps.readActor(),
    ]);
    checkDraft(afterDraft, input, draft);
    sourceFile(afterSource, input);
    requireValue(afterActor.id === actor.id, 'ACTOR_ID_DRIFT');
    const after = children(afterList);
    ensureBaseline(after, initial, ack.id);
    const added = after.find((child) => child.id === ack.id) as PriorChild;
    requireValue(
      added.name === name && added.contentType === 'application/pdf' && !added.isInline,
      'NEW_ATTACHMENT_METADATA_DRIFT'
    );
    const newBytes = await deps.readAttachmentBytes(ack.id, MAX_DRAFT_PDF_BYTES);
    requireValue(
      Buffer.isBuffer(newBytes) &&
        newBytes.length === bytes.length &&
        digest(newBytes) === hash,
      'NEW_ATTACHMENT_BYTES_DRIFT'
    );
    const finalHashes = await readExisting(deps, initial);
    for (const child of initial) {
      requireValue(finalHashes.get(child.id) === priorHashes.get(child.id), 'PRIOR_BYTES_DRIFT');
    }
    const [finalDraft, finalList] = await Promise.all([deps.readDraft(), deps.listAttachments()]);
    checkDraft(finalDraft, input, draft);
    requireValue(
      finalDraft.changeKey === afterDraft.changeKey,
      'DRAFT_CHANGED_DURING_READBACK'
    );
    ensureBaseline(children(finalList), initial, ack.id);
    return {
      status: 'VERIFIED_DRAFT_ATTACHMENT',
      attemptId,
      messageId: input.messageId,
      driveItemId: input.driveItemId,
      attachmentId: ack.id,
      filename: name,
      decodedBytes: bytes.length,
      sha256: hash,
      retainedAttachments: initial.length,
      isDraft: true,
      sent: false,
    };
  } catch {
    throw new DrivePdfDraftError('POST_APPLIED_UNVERIFIED_NO_REPLAY', true, attemptId);
  }
}
