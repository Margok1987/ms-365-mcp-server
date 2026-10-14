import { lookup } from 'node:dns/promises';
import type { IncomingMessage } from 'node:http';
import { mkdtemp, open, rm, type FileHandle } from 'node:fs/promises';
import { request as httpsRequest } from 'node:https';
import { isIP } from 'node:net';
import { tmpdir } from 'node:os';
import path from 'node:path';
import { Readable } from 'node:stream';

export interface HostFileParam {
  download_url: string;
  file_id: string;
  mime_type?: string;
  file_name?: string;
}

export interface DownloadedHostFile {
  handle: FileHandle;
  size: number;
  contentType?: string;
  cleanup: () => Promise<void>;
}

interface DownloadDeps {
  fetchImpl?: typeof fetch;
  lookupAll?: (hostname: string) => Promise<string[]>;
}

const MAX_REDIRECTS = 5;

function ipv4Parts(address: string): number[] | null {
  const parts = address.split('.').map((part) => Number.parseInt(part, 10));
  if (
    parts.length !== 4 ||
    parts.some((part) => !Number.isInteger(part) || part < 0 || part > 255)
  ) {
    return null;
  }
  return parts;
}

function isNonPublicIpv4(address: string): boolean {
  const parts = ipv4Parts(address);
  if (!parts) return true;
  const [a, b, c] = parts;
  return (
    a === 0 ||
    a === 10 ||
    a === 127 ||
    (a === 100 && b >= 64 && b <= 127) ||
    (a === 169 && b === 254) ||
    (a === 172 && b >= 16 && b <= 31) ||
    (a === 192 && b === 0 && (c === 0 || c === 2)) ||
    (a === 192 && b === 88 && c === 99) ||
    (a === 192 && b === 168) ||
    (a === 198 && (b === 18 || b === 19)) ||
    (a === 198 && b === 51 && c === 100) ||
    (a === 203 && b === 0 && c === 113) ||
    a >= 224
  );
}

function isNonPublicIpv6(address: string): boolean {
  const normalized = address.toLowerCase().split('%')[0];
  if (normalized === '::' || normalized === '::1') return true;
  if (normalized.startsWith('fc') || normalized.startsWith('fd')) return true;
  if (/^fe[89ab]/.test(normalized)) return true;
  if (normalized.startsWith('ff')) return true;
  if (normalized.startsWith('2001:db8:')) return true;
  // Reject the entire IPv4-mapped block rather than trying to normalize every textual form.
  if (normalized.startsWith('::ffff:')) return true;
  return false;
}

export function isPublicAddress(address: string): boolean {
  const family = isIP(address);
  if (family === 4) return !isNonPublicIpv4(address);
  if (family === 6) return !isNonPublicIpv6(address);
  return false;
}

async function defaultLookupAll(hostname: string): Promise<string[]> {
  return (await lookup(hostname, { all: true, verbatim: true })).map((entry) => entry.address);
}

async function validateDownloadUrl(
  rawUrl: string,
  lookupAll: (hostname: string) => Promise<string[]>
): Promise<{ url: URL; addresses: string[] }> {
  let url: URL;
  try {
    url = new URL(rawUrl);
  } catch {
    throw new Error('file.download_url must be a valid HTTPS URL.');
  }
  if (url.protocol !== 'https:') {
    throw new Error('file.download_url must use HTTPS.');
  }
  if (url.username || url.password) {
    throw new Error('file.download_url must not contain URL credentials.');
  }
  if (url.port && url.port !== '443') {
    throw new Error('file.download_url must use the default HTTPS port.');
  }

  const urlHostname = url.hostname.replace(/\.$/, '').toLowerCase();
  const hostname =
    urlHostname.startsWith('[') && urlHostname.endsWith(']')
      ? urlHostname.slice(1, -1)
      : urlHostname;
  if (hostname === 'localhost' || hostname.endsWith('.localhost') || hostname.endsWith('.local')) {
    throw new Error('file.download_url resolves to a local address and was refused.');
  }

  const addresses = isIP(hostname) ? [hostname] : await lookupAll(hostname);
  if (addresses.length === 0 || addresses.some((address) => !isPublicAddress(address))) {
    throw new Error('file.download_url resolves to a non-public address and was refused.');
  }
  return { url, addresses };
}

function responseHeaders(incoming: IncomingMessage): Headers {
  const headers = new Headers();
  for (const [name, value] of Object.entries(incoming.headers)) {
    if (Array.isArray(value)) {
      for (const item of value) headers.append(name, item);
    } else if (value !== undefined) {
      headers.set(name, value);
    }
  }
  return headers;
}

async function fetchPinnedHttps(url: URL, addresses: string[]): Promise<Response> {
  let lastError: unknown;
  for (const address of addresses) {
    try {
      return await new Promise<Response>((resolve, reject) => {
        const family = isIP(address);
        if (family !== 4 && family !== 6) {
          reject(new Error('file.download_url resolved to an invalid IP address.'));
          return;
        }
        const request = httpsRequest(
          url,
          {
            method: 'GET',
            agent: false,
            headers: { 'Accept-Encoding': 'identity' },
            lookup: (_hostname, _options, callback) => callback(null, address, family),
          },
          (incoming) => {
            const status = incoming.statusCode ?? 502;
            const bodyAllowed = ![204, 205, 304].includes(status);
            const body = bodyAllowed
              ? (Readable.toWeb(incoming) as ReadableStream<Uint8Array>)
              : null;
            resolve(
              new Response(body, {
                status,
                statusText: incoming.statusMessage,
                headers: responseHeaders(incoming),
              })
            );
          }
        );
        request.setTimeout(60_000, () => {
          request.destroy(new Error('file.download_url timed out.'));
        });
        request.once('error', reject);
        request.end();
      });
    } catch (error) {
      lastError = error;
    }
  }
  throw lastError instanceof Error
    ? lastError
    : new Error('file.download_url could not be reached.');
}

async function fetchValidated(rawUrl: string, deps: DownloadDeps): Promise<Response> {
  const lookupAll = deps.lookupAll ?? defaultLookupAll;
  let current = rawUrl;

  for (let redirects = 0; redirects <= MAX_REDIRECTS; redirects++) {
    const { url, addresses } = await validateDownloadUrl(current, lookupAll);
    const response = deps.fetchImpl
      ? await deps.fetchImpl(url, {
          method: 'GET',
          redirect: 'manual',
          headers: { 'Accept-Encoding': 'identity' },
        })
      : await fetchPinnedHttps(url, addresses);

    if ([301, 302, 303, 307, 308].includes(response.status)) {
      const location = response.headers.get('location');
      await response.body?.cancel().catch(() => undefined);
      if (!location) throw new Error('file.download_url redirect had no Location header.');
      if (redirects === MAX_REDIRECTS) {
        throw new Error(`file.download_url exceeded ${MAX_REDIRECTS} redirects.`);
      }
      current = new URL(location, url).toString();
      continue;
    }

    if (!response.ok) {
      await response.body?.cancel().catch(() => undefined);
      throw new Error(`file.download_url answered HTTP ${response.status}.`);
    }
    if (!response.body) throw new Error('file.download_url returned no response body.');
    return response;
  }

  throw new Error(`file.download_url exceeded ${MAX_REDIRECTS} redirects.`);
}

function contentLength(response: Response): number | undefined {
  const raw = response.headers.get('content-length');
  if (raw === null) return undefined;
  if (!/^\d+$/.test(raw)) throw new Error('file.download_url returned an invalid Content-Length.');
  const parsed = Number(raw);
  if (!Number.isSafeInteger(parsed)) {
    throw new Error('file.download_url returned an invalid Content-Length.');
  }
  return parsed;
}

/**
 * Materialize a host-provided file URL into a private temporary file.
 *
 * The MCP file-param object contains a temporary URL but not a byte length. Graph
 * attachment upload sessions require the exact length before the first chunk, so
 * the download is staged once, counted with a hard limit, then streamed to Graph.
 * Every URL and redirect is restricted to public HTTPS destinations.
 */
export async function downloadHostFile(
  file: HostFileParam,
  maxBytes: number,
  deps: DownloadDeps = {}
): Promise<DownloadedHostFile> {
  if (!file.file_id || typeof file.file_id !== 'string') {
    throw new Error('file.file_id is required.');
  }
  if (!file.download_url || typeof file.download_url !== 'string') {
    throw new Error('file.download_url is required.');
  }

  const response = await fetchValidated(file.download_url, deps);
  const declaredLength = contentLength(response);
  if (declaredLength !== undefined) {
    if (declaredLength === 0) {
      await response.body?.cancel().catch(() => undefined);
      throw new Error('The file is empty.');
    }
    if (declaredLength > maxBytes) {
      await response.body?.cancel().catch(() => undefined);
      throw new Error(
        `The file is ${declaredLength} bytes; the largest attachment is ${maxBytes} bytes.`
      );
    }
  }

  const dir = await mkdtemp(path.join(tmpdir(), 'ms365-mcp-host-file-'));
  const filePath = path.join(dir, 'payload');
  let writeHandle: FileHandle | undefined;
  let readHandle: FileHandle | undefined;

  try {
    writeHandle = await open(filePath, 'wx', 0o600);
    let size = 0;
    for await (const chunk of response.body as AsyncIterable<Uint8Array>) {
      const buffer = Buffer.from(chunk);
      size += buffer.length;
      if (size > maxBytes) {
        throw new Error(`The file exceeds the largest attachment size of ${maxBytes} bytes.`);
      }
      await writeHandle.write(buffer);
    }
    await writeHandle.close();
    writeHandle = undefined;

    if (size === 0) throw new Error('The file is empty.');
    if (declaredLength !== undefined && declaredLength !== size) {
      throw new Error(
        `file.download_url declared ${declaredLength} bytes but returned ${size} bytes.`
      );
    }

    readHandle = await open(filePath, 'r');
    const contentType = response.headers.get('content-type')?.split(';', 1)[0]?.trim() || undefined;
    const handle = readHandle;
    return {
      handle,
      size,
      ...(contentType ? { contentType } : {}),
      cleanup: async () => {
        await handle.close().catch(() => undefined);
        await rm(dir, { recursive: true, force: true });
      },
    };
  } catch (error) {
    await writeHandle?.close().catch(() => undefined);
    await readHandle?.close().catch(() => undefined);
    await rm(dir, { recursive: true, force: true });
    throw error;
  }
}
