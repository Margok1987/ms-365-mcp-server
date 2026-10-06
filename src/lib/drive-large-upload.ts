import type { LookupAll, OpenAIFileParam } from './openai-file-source.js';
import { stageOpenAIFile } from './openai-file-source.js';
import {
  UploadSessionTransportError,
  uploadChunksToSession,
  type FetchLike,
} from './upload-session-transport.js';

export interface DriveItemReadback {
  id?: string;
  name?: string;
  size?: number;
  parentReference?: { driveId?: string; id?: string };
  [key: string]: unknown;
}

export interface LargeDriveUploadDependencies {
  createUploadSession(input: {
    driveId: string;
    parentItemId: string;
    fileName: string;
  }): Promise<{ uploadUrl: string }>;
  readDestination(input: {
    driveId: string;
    parentItemId: string;
    fileName: string;
  }): Promise<DriveItemReadback | null>;
  fetchImpl?: FetchLike;
  lookupAll?: LookupAll;
  tempRoot?: string;
}

export interface LargeDriveUploadInput {
  file: OpenAIFileParam;
  driveId: string;
  parentItemId: string;
  fileName?: string;
}

export interface LargeDriveUploadResult {
  id: string;
  name: string;
  size: number;
  driveId: string;
  parentItemId: string;
  chunksUploaded: number;
  reconciledAfterUncertainFinal: boolean;
}

export class LargeDriveUploadError extends Error {
  constructor(
    public readonly code: string,
    message: string,
    public readonly details: Record<string, unknown> = {}
  ) {
    super(message);
    this.name = 'LargeDriveUploadError';
  }
}

export function validateDriveUploadFileName(name: string): string {
  const value = name.trim();
  if (!value || value === '.' || value === '..') {
    throw new LargeDriveUploadError('INVALID_FILE_NAME', 'Destination file name is invalid.');
  }
  if (/[\\/:*?"<>|]/.test(value) || value.endsWith('.')) {
    throw new LargeDriveUploadError(
      'INVALID_FILE_NAME',
      'Destination file name contains characters not accepted by OneDrive path addressing.'
    );
  }
  return value;
}

function requiredString(value: unknown, field: string): string {
  if (typeof value !== 'string' || value.trim().length === 0) {
    throw new LargeDriveUploadError('INVALID_INPUT', `${field} must be a non-empty string.`);
  }
  return value.trim();
}

function verifyReadback(
  item: DriveItemReadback | null,
  expected: {
    driveId: string;
    parentItemId: string;
    fileName: string;
    size: number;
    finalItemId?: string;
  }
): DriveItemReadback {
  if (!item) {
    throw new LargeDriveUploadError(
      'UPLOAD_READBACK_MISSING',
      'Upload completed but exact destination readback is missing.'
    );
  }
  if (typeof item.id !== 'string' || item.id.length === 0) {
    throw new LargeDriveUploadError('UPLOAD_READBACK_INVALID', 'Readback has no driveItem id.');
  }
  if (expected.finalItemId && item.id !== expected.finalItemId) {
    throw new LargeDriveUploadError(
      'UPLOAD_ID_MISMATCH',
      'Final upload response and exact destination readback refer to different driveItems.',
      { finalItemId: expected.finalItemId, readbackItemId: item.id }
    );
  }
  if (item.name !== expected.fileName) {
    throw new LargeDriveUploadError(
      'UPLOAD_NAME_MISMATCH',
      'Exact destination readback has an unexpected name.',
      { expected: expected.fileName, actual: item.name }
    );
  }
  if (item.size !== expected.size) {
    throw new LargeDriveUploadError(
      'UPLOAD_SIZE_MISMATCH',
      'Exact destination readback size does not match the staged ChatGPT file.',
      { expected: expected.size, actual: item.size }
    );
  }
  if (
    item.parentReference?.driveId !== undefined &&
    item.parentReference.driveId !== expected.driveId
  ) {
    throw new LargeDriveUploadError(
      'UPLOAD_DRIVE_MISMATCH',
      'Exact destination readback belongs to a different drive.'
    );
  }
  if (
    item.parentReference?.id !== undefined &&
    item.parentReference.id !== expected.parentItemId
  ) {
    throw new LargeDriveUploadError(
      'UPLOAD_PARENT_MISMATCH',
      'Exact destination readback belongs to a different parent folder.'
    );
  }
  return item;
}

export async function uploadOpenAIFileToDrive(
  input: LargeDriveUploadInput,
  deps: LargeDriveUploadDependencies
): Promise<LargeDriveUploadResult> {
  const driveId = requiredString(input.driveId, 'driveId');
  const parentItemId = requiredString(input.parentItemId, 'parentItemId');
  const fileName = validateDriveUploadFileName(input.fileName ?? input.file.file_name ?? '');

  const staged = await stageOpenAIFile(input.file, {
    fetchImpl: deps.fetchImpl as typeof fetch | undefined,
    lookupAll: deps.lookupAll,
    tempRoot: deps.tempRoot,
  });

  try {
    const preState = await deps.readDestination({ driveId, parentItemId, fileName });
    if (preState) {
      throw new LargeDriveUploadError(
        'DESTINATION_EXISTS',
        'Destination already exists; large upload uses conflictBehavior=fail and will not overwrite it.',
        { itemId: preState.id ?? null }
      );
    }

    const session = await deps.createUploadSession({ driveId, parentItemId, fileName });
    if (!session || typeof session.uploadUrl !== 'string' || session.uploadUrl.length === 0) {
      throw new LargeDriveUploadError(
        'UPLOAD_SESSION_INVALID',
        'Provider did not return an uploadUrl.'
      );
    }

    try {
      const uploaded = await uploadChunksToSession({
        uploadUrl: session.uploadUrl,
        totalBytes: staged.size,
        source: staged.open(),
        fetchImpl: deps.fetchImpl,
      });
      const finalId =
        typeof uploaded.driveItem.id === 'string' ? uploaded.driveItem.id : undefined;
      const readback = verifyReadback(
        await deps.readDestination({ driveId, parentItemId, fileName }),
        { driveId, parentItemId, fileName, size: staged.size, finalItemId: finalId }
      );
      return {
        id: readback.id as string,
        name: fileName,
        size: staged.size,
        driveId,
        parentItemId,
        chunksUploaded: uploaded.chunksUploaded,
        reconciledAfterUncertainFinal: false,
      };
    } catch (error) {
      if (
        error instanceof UploadSessionTransportError &&
        error.code === 'FINAL_CHUNK_RESULT_UNCERTAIN'
      ) {
        const readback = await deps.readDestination({ driveId, parentItemId, fileName });
        try {
          const verified = verifyReadback(readback, {
            driveId,
            parentItemId,
            fileName,
            size: staged.size,
          });
          return {
            id: verified.id as string,
            name: fileName,
            size: staged.size,
            driveId,
            parentItemId,
            chunksUploaded: 0,
            reconciledAfterUncertainFinal: true,
          };
        } catch {
          throw error;
        }
      }
      throw error;
    }
  } finally {
    await staged.cleanup();
  }
}
