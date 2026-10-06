import { createReadStream, createWriteStream } from 'fs';
import { mkdtemp, rm, stat } from 'fs/promises';
import { isIP } from 'net';
import { tmpdir } from 'os';
import path from 'path';
import { Readable } from 'stream';
import { pipeline } from 'stream/promises';

export interface OpenAIFileParam {
  download_url: string;
  file_id: string;
  mime_type?: string;
  file_name?: string;
}

export interface StagedOpenAIFile {
  path: string;
  size: number;
  mimeType?: string;
  fileName?: string;
  open(): AsyncIterable<Uint8Array>;
  cleanup(): Promise<void>;
}

export class OpenAIFileSourceError extends Error {
  constructor(
    public readonly code: string,
    message: string,
    public readonly details: Record<string, unknown> = {}
  ) {
    super(message);
    this.name = 'OpenAIFileSourceError';
  }
}

function isPrivateIpv4(host: string): boolean {
  const octets = host.split('.').map(Number);
  if (octets.length !== 4 || octets.some((n) => !Number.isInteger(n) || n < 0 || n > 255)) {
    return false;
  }
  const [a, b] = octets;
  return (
    a === 10 ||
    a === 127 ||
    a === 0 ||
    (a === 169 && b === 254) ||
    (a === 172 && b >= 16 && b <= 31) ||
    (a === 192 && b === 168)
  );
}

function isPrivateIpv6(host: string): boolean {
  const h = host.toLowerCase();
  return (
    h === '::1' ||
    h === '::' ||
    h.startsWith('fc') ||
    h.startsWith('fd') ||
    /^fe[89ab]/.test(h)
  );
}

export function assertSafeFileDownloadUrl(raw: string): URL {
  let url: URL;
  try {
    url = new URL(raw);
  } catch {
    throw new OpenAIFileSourceError('INVALID_FILE_URL', 'File download URL is invalid.');
  }
  if (url.protocol !== 'https:') {
    throw new OpenAIFileSourceError('INVALID_FILE_URL', 'File download URL must use HTTPS.');
  }
  if (url.username || url.password) {
    throw new OpenAIFileSourceError(
      'INVALID_FILE_URL',
      'File download URL must not contain URL credentials.'
    );
  }
  const host = url.hostname.toLowerCase().replace(/^\[|\]$/g, '');
  if (host === 'localhost' || host.endsWith('.localhost')) {
    throw new OpenAIFileSourceError('UNSAFE_FILE_URL', 'Localhost file URLs are not allowed.');
  }
  const ipKind = isIP(host);
  if ((ipKind === 4 && isPrivateIpv4(host)) || (ipKind === 6 && isPrivateIpv6(host))) {
    throw new OpenAIFileSourceError(
      'UNSAFE_FILE_URL',
      'Private, loopback, link-local, or unspecified IP file URLs are not allowed.'
    );
  }
  return url;
}

async function fetchWithSafeRedirects(
  initial: URL,
  fetchImpl: typeof fetch,
  maxRedirects = 5
): Promise<Response> {
  let current = initial;
  for (let redirect = 0; redirect <= maxRedirects; redirect += 1) {
    const response = await fetchImpl(current, { method: 'GET', redirect: 'manual' });
    if (response.status >= 300 && response.status < 400) {
      if (redirect === maxRedirects) {
        throw new OpenAIFileSourceError(
          'TOO_MANY_FILE_REDIRECTS',
          'File download exceeded the redirect limit.'
        );
      }
      const location = response.headers.get('location');
      if (!location) {
        throw new OpenAIFileSourceError(
          'INVALID_FILE_REDIRECT',
          'File download redirect omitted Location.'
        );
      }
      current = assertSafeFileDownloadUrl(new URL(location, current).toString());
      continue;
    }
    if (!response.ok) {
      throw new OpenAIFileSourceError(
        'FILE_DOWNLOAD_REJECTED',
        `File download failed with HTTP ${response.status}.`,
        { status: response.status }
      );
    }
    return response;
  }
  throw new OpenAIFileSourceError('TOO_MANY_FILE_REDIRECTS', 'Redirect limit exceeded.');
}

function validateFileParam(file: OpenAIFileParam): void {
  if (!file || typeof file !== 'object') {
    throw new OpenAIFileSourceError('INVALID_FILE_PARAM', 'file must be an object.');
  }
  if (typeof file.download_url !== 'string' || file.download_url.length === 0) {
    throw new OpenAIFileSourceError(
      'INVALID_FILE_PARAM',
      'file.download_url must be a non-empty string.'
    );
  }
  if (typeof file.file_id !== 'string' || file.file_id.length === 0) {
    throw new OpenAIFileSourceError(
      'INVALID_FILE_PARAM',
      'file.file_id must be a non-empty string.'
    );
  }
}

export async function stageOpenAIFile(
  file: OpenAIFileParam,
  options: { fetchImpl?: typeof fetch; tempRoot?: string } = {}
): Promise<StagedOpenAIFile> {
  validateFileParam(file);
  const fetchImpl = options.fetchImpl ?? fetch;
  const url = assertSafeFileDownloadUrl(file.download_url);
  const response = await fetchWithSafeRedirects(url, fetchImpl);
  if (!response.body) {
    throw new OpenAIFileSourceError('EMPTY_FILE_RESPONSE', 'File download response has no body.');
  }

  const root = options.tempRoot ?? tmpdir();
  const dir = await mkdtemp(path.join(root, 'ms365-file-upload-'));
  const filePath = path.join(dir, 'source.bin');

  try {
    const nodeStream = Readable.fromWeb(
      response.body as unknown as import('stream/web').ReadableStream<Uint8Array>
    );
    await pipeline(
      nodeStream,
      createWriteStream(filePath, { flags: 'wx', mode: 0o600 })
    );
    const info = await stat(filePath);
    if (info.size <= 0) {
      throw new OpenAIFileSourceError('EMPTY_FILE', 'Zero-byte uploads are not supported.');
    }

    return {
      path: filePath,
      size: info.size,
      mimeType: file.mime_type,
      fileName: file.file_name,
      open: () => createReadStream(filePath),
      cleanup: async () => {
        await rm(dir, { recursive: true, force: true });
      },
    };
  } catch (error) {
    await rm(dir, { recursive: true, force: true });
    throw error;
  }
}
