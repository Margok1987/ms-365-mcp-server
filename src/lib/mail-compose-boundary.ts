/** Runtime-owned public dispatch boundary, never a caller-supplied approval.
 * Enabled instances require a separately bound private composer. This module
 * deliberately has no caller-visible/internal=true escape hatch.
 */
export class MailComposeBoundaryError extends Error {
  constructor() {
    super('BOUND_MAIL_COMPOSE_REQUIRED');
    this.name = 'MailComposeBoundaryError';
  }
}

function deny(): never {
  throw new MailComposeBoundaryError();
}

function record(value: unknown): value is Record<string, unknown> {
  return value !== null && typeof value === 'object' && !Array.isArray(value);
}

function configured(): boolean {
  const value = process.env.MS365_MCP_REQUIRE_BOUND_MAIL_COMPOSE;
  if (value === undefined || value === '0' || value === 'false') return false;
  if (value === '1' || value === 'true') return true;
  // A misspelled explicit policy must never silently remove the boundary.
  return deny();
}

function invalidPathCharacters(value: string): boolean {
  return [...value].some((c) => c.charCodeAt(0) <= 32 || c.charCodeAt(0) === 127 || c === '\\');
}

function pathOf(endpoint: string): string {
  if (typeof endpoint !== 'string' || endpoint.length > 16384 || invalidPathCharacters(endpoint))
    deny();
  let path = endpoint.split('?')[0];
  if (endpoint.includes('#')) deny();
  if (/^https?:\/\//i.test(endpoint)) {
    const u = new URL(endpoint);
    if (
      u.protocol !== 'https:' ||
      u.host !== 'graph.microsoft.com' ||
      u.username ||
      u.password ||
      u.hash ||
      !/^https:\/\/graph\.microsoft\.com\//i.test(endpoint)
    )
      deny();
    // Check the unnormalised text first: URL would erase dot segments.
    path = endpoint.replace(/^https:\/\/graph\.microsoft\.com/i, '').split('?')[0];
  }
  if (!path.startsWith('/')) path = '/' + path;
  try {
    path = decodeURIComponent(path);
  } catch {
    return deny();
  }
  if (
    /[%#?]/.test(path) ||
    invalidPathCharacters(path) ||
    path.includes('//') ||
    path.split('/').some((x) => x === '.' || x === '..')
  )
    deny();
  path = path.replace(/^\/(?:v1\.0|beta)(?=\/)/i, '');
  return path.toLowerCase();
}

function parsedBody(body: unknown): Record<string, unknown> {
  if (typeof body === 'string') {
    if (Buffer.byteLength(body, 'utf8') > 2 * 1024 * 1024) deny();
    try {
      const parsed: unknown = JSON.parse(body);
      const stack: Array<{ keys: Set<string>; key: boolean } | null> = [];
      for (const token of body.match(/"(?:\\.|[^"\\])*"|[{}[\],:]|[^{}[\],:\s]+/g) ?? []) {
        if (token === '{') stack.push({ keys: new Set(), key: true });
        else if (token === '[') stack.push(null);
        else if (token === '}' || token === ']') stack.pop();
        else {
          const top = stack[stack.length - 1];
          if (token === ',' && top) top.key = true;
          else if (token.startsWith('"') && top?.key) {
            const name: string = JSON.parse(token);
            if (top.keys.has(name)) deny();
            top.keys.add(name);
            top.key = false;
          }
        }
      }
      body = parsed;
    } catch {
      return deny();
    }
  }
  if (!record(body)) deny();
  return body;
}

function checkHeaders(headers: unknown): void {
  if (headers === undefined) return;
  if (!record(headers)) deny();
  const seen = new Set<string>();
  for (const [key, value] of Object.entries(headers)) {
    const name = key.toLowerCase();
    if (
      seen.has(name) ||
      typeof value !== 'string' ||
      !key ||
      invalidPathCharacters(key) ||
      [...value].some((c) => c.charCodeAt(0) < 32 || c.charCodeAt(0) === 127)
    )
      deny();
    seen.add(name);
    if (
      [
        'x-http-method',
        'x-http-method-override',
        'x-method-override',
        'authorization',
        'host',
      ].includes(name)
    )
      deny();
  }
}

function inspect(
  method: string,
  endpoint: string,
  body: unknown,
  depth = 0,
  headers?: unknown
): void {
  checkHeaders(headers);
  if (typeof method !== 'string' || !/^(GET|HEAD|POST|PATCH|PUT|DELETE)$/i.test(method)) deny();
  const verb = method.toUpperCase();
  const path = pathOf(endpoint);
  if (path === '/$batch' || path === '/$batch/') {
    if (verb !== 'POST' || depth !== 0) deny();
    const b = parsedBody(body);
    if (
      Object.keys(b).some((k) => k !== 'requests') ||
      !Array.isArray(b.requests) ||
      b.requests.length < 1 ||
      b.requests.length > 20
    )
      deny();
    const ids = new Set<string>();
    for (const r of b.requests) {
      if (
        !record(r) ||
        Object.keys(r).some(
          (k) => !['id', 'method', 'url', 'headers', 'body', 'dependsOn'].includes(k)
        ) ||
        typeof r.id !== 'string' ||
        !r.id ||
        ids.has(r.id) ||
        typeof r.method !== 'string' ||
        typeof r.url !== 'string'
      )
        deny();
      ids.add(r.id);
      inspect(r.method, r.url, r.body, depth + 1, r.headers);
    }
    return;
  }
  if (verb === 'GET' || verb === 'HEAD') return;
  if (/^\/(?:me|users(?:\/[^/]+|\([^/]+\)))\/sendmail(?:\/|$)/.test(path)) deny();
  // Covers folder-scoped messages, shared users, OData key selectors and child
  // attachments/actions. No send/delete/child-write authority is inferred.
  const mail =
    /^\/(?:me|users(?:\/[^/]+|\([^/]+\)))\/(?:mailfolders(?:\/[^/]+|\([^/]+\))\/)?messages(?:\/|\(|$)/.test(
      path
    );
  if (!mail) return;
  if (
    verb !== 'PATCH' ||
    !/^\/(?:me|users\/[^/]+)\/(?:mailfolders\/[^/]+\/)?messages\/[^/]+$/.test(path)
  )
    deny();
  const b = parsedBody(body);
  const keys = Object.keys(b);
  // Existing closed metadata writers may continue. Content/recipient/provider
  // extensions, unknown writes and any child mutation require private binding.
  if (
    !keys.length ||
    keys.some(
      (k) => !['isRead', 'importance', 'categories', 'flag', 'inferenceClassification'].includes(k)
    )
  )
    deny();
  if ('isRead' in b && typeof b.isRead !== 'boolean') deny();
  if ('importance' in b && !['low', 'normal', 'high'].includes(b.importance as string)) deny();
  if (
    'inferenceClassification' in b &&
    !['focused', 'other'].includes(b.inferenceClassification as string)
  )
    deny();
  if (
    'categories' in b &&
    (!Array.isArray(b.categories) || b.categories.some((x) => typeof x !== 'string'))
  )
    deny();
  if (
    'flag' in b &&
    (!record(b.flag) ||
      Object.keys(b.flag).some(
        (k) => !['flagStatus', 'startDateTime', 'dueDateTime', 'completedDateTime'].includes(k)
      ))
  )
    deny();
}

export function assertPublicMailComposeBoundary(
  method: string,
  endpoint: string,
  body?: unknown,
  headers?: unknown
): void {
  if (configured()) inspect(method, endpoint, body, 0, headers);
}
