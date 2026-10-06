import { mkdtemp } from 'fs/promises';
import { tmpdir } from 'os';
import path from 'path';
import { describe, expect, it, vi } from 'vitest';
import { uploadOpenAIFileToDrive, validateDriveUploadFileName } from './drive-large-upload.js';

function fileFetch(bytes: Uint8Array) {
  return vi.fn(async (_url: string, init?: RequestInit) => {
    if (!init?.method || init.method === 'GET') return new Response(Buffer.from(bytes));
    throw new Error('unexpected source request');
  });
}

describe('uploadOpenAIFileToDrive', () => {
  it('qualifies a new-file upload with exact destination readback', async () => {
    const tempRoot = await mkdtemp(path.join(tmpdir(), 'drive-large-upload-'));
    let exists = false;
    const sourceFetch = fileFetch(new Uint8Array([1, 2, 3, 4]));
    const fetchImpl = vi.fn(async (url: string, init?: RequestInit) => {
      if (String(url) === 'https://files.example.test/file') {
        return sourceFetch(String(url), init);
      }
      if (String(url) === 'https://upload.example.test/session' && init?.method === 'PUT') {
        exists = true;
        return new Response(JSON.stringify({ id: 'item-1', name: 'report.pdf', size: 4 }), {
          status: 201,
          headers: { 'content-type': 'application/json' },
        });
      }
      throw new Error(`unexpected request ${url} ${init?.method}`);
    });

    const result = await uploadOpenAIFileToDrive(
      {
        file: {
          download_url: 'https://files.example.test/file',
          file_id: 'file_1',
          file_name: 'report.pdf',
        },
        driveId: 'drive-1',
        parentItemId: 'parent-1',
      },
      {
        tempRoot,
        fetchImpl,
        createUploadSession: vi.fn(async () => ({
          uploadUrl: 'https://upload.example.test/session',
        })),
        readDestination: vi.fn(async () =>
          exists
            ? {
                id: 'item-1',
                name: 'report.pdf',
                size: 4,
                parentReference: { driveId: 'drive-1', id: 'parent-1' },
              }
            : null
        ),
      }
    );

    expect(result).toMatchObject({
      id: 'item-1',
      name: 'report.pdf',
      size: 4,
      driveId: 'drive-1',
      parentItemId: 'parent-1',
      reconciledAfterUncertainFinal: false,
    });
  });

  it('fails before session creation when the destination exists', async () => {
    const tempRoot = await mkdtemp(path.join(tmpdir(), 'drive-large-upload-'));
    const createUploadSession = vi.fn();
    await expect(
      uploadOpenAIFileToDrive(
        {
          file: {
            download_url: 'https://files.example.test/file',
            file_id: 'file_2',
            file_name: 'exists.pdf',
          },
          driveId: 'drive-1',
          parentItemId: 'parent-1',
        },
        {
          tempRoot,
          fetchImpl: fileFetch(new Uint8Array([1])) as unknown as typeof fetch,
          createUploadSession,
          readDestination: vi.fn(async () => ({
            id: 'existing',
            name: 'exists.pdf',
            size: 99,
          })),
        }
      )
    ).rejects.toMatchObject({ code: 'DESTINATION_EXISTS' });
    expect(createUploadSession).not.toHaveBeenCalled();
  });

  it('reconciles an uncertain final PUT by exact destination path without replay', async () => {
    const tempRoot = await mkdtemp(path.join(tmpdir(), 'drive-large-upload-'));
    let afterPut = false;
    let uploadPuts = 0;
    const fetchImpl = vi.fn(async (url: string, init?: RequestInit) => {
      if (String(url) === 'https://files.example.test/file') {
        return new Response(new Uint8Array([7, 8, 9]));
      }
      if (String(url) === 'https://upload.example.test/session' && init?.method === 'PUT') {
        uploadPuts += 1;
        afterPut = true;
        throw new Error('connection reset after provider accepted final range');
      }
      throw new Error('unexpected request');
    });

    const result = await uploadOpenAIFileToDrive(
      {
        file: {
          download_url: 'https://files.example.test/file',
          file_id: 'file_3',
          file_name: 'uncertain.pdf',
        },
        driveId: 'drive-1',
        parentItemId: 'parent-1',
      },
      {
        tempRoot,
        fetchImpl,
        createUploadSession: vi.fn(async () => ({
          uploadUrl: 'https://upload.example.test/session',
        })),
        readDestination: vi.fn(async () =>
          afterPut
            ? {
                id: 'item-final',
                name: 'uncertain.pdf',
                size: 3,
                parentReference: { driveId: 'drive-1', id: 'parent-1' },
              }
            : null
        ),
      }
    );

    expect(result.reconciledAfterUncertainFinal).toBe(true);
    expect(result.id).toBe('item-final');
    expect(uploadPuts).toBe(1);
  });

  it('rejects filename shapes that cannot be addressed as one OneDrive child', () => {
    for (const name of ['', '.', '..', 'a/b.pdf', 'a\\b.pdf', 'bad?.pdf', 'trailing.']) {
      expect(() => validateDriveUploadFileName(name)).toThrow();
    }
    expect(validateDriveUploadFileName('valid report.pdf')).toBe('valid report.pdf');
  });
});
