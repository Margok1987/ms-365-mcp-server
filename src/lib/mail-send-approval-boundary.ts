/**
 * Fail-closed outbound Mail Send choke-point for deployments that intend to
 * expose send tool names BEFORE installing the trusted proof issuer.
 *
 * IMPORTANT: This module provides a DENY-UNTIL-BOUND mode, not approval itself.
 * The caller cannot supply an `approved`, `confirm`, or `internal` escape hatch.
 * The future positive route MUST verify independently authenticated human proof
 * and the exact native draft snapshot before allowing any dispatch.
 */
export class MailSendProofRequiredError extends Error {
  constructor() {
    super('MAIL_SEND_TRUSTED_PROOF_REQUIRED');
    this.name = 'MailSendProofRequiredError';
  }
}

function deny(): never {
  throw new MailSendProofRequiredError();
}

function enabled(): boolean {
  const value = process.env.MS365_MCP_REQUIRE_APPROVED_SEND;
  if (value === undefined || value === '0' || value === 'false') return false;
  if (value === '1' || value === 'true') return true;
  return deny();
}

function isRecord(value: unknown): value is Record<string, unknown> {
  return !!value && typeof value === 'object' && !Array.isArray(value);
}

function cleanPath(endpoint: string): string {
  if (typeof endpoint !== 'string' || endpoint.length > 8192 || /[\x00-\x20\x7f\\#]/.test(endpoint))
    return deny();
  let raw = endpoint.split('?')[0];
  if (/^https?:\/\//i.test(raw)) {
    const url = new URL(raw);
    if (url.protocol !== 'https:' || url.host !== 'graph.microsoft.com' ||
        url.username || url.password || url.hash ||
        !/^https:\/\/graph\.microsoft\.com\//i.test(raw)) return deny();
    raw = raw.substring('https://graph.microsoft.com'.length);
  }
  if (!raw.startsWith('/')) raw = '/' + raw;
  try { raw = decodeURIComponent(raw); } catch { return deny(); }
  if (/[%?#\x00-\x20\x7f\\]/.test(raw) || raw.includes('//') ||
      raw.split('/').some(x => x === '.' || x === '..')) return deny();
  return raw.replace(/^\/(?:v1\.0|beta)(?=\/)/i, '').toLowerCase();
}

function checkHeaders(headers: unknown): void {
  if (headers === undefined) return;
  if (!isRecord(headers)) deny();
  const seen = new Set<string>();
  for (const [key, value] of Object.entries(headers)) {
    const name = key.toLowerCase();
    if (seen.has(name) || typeof value !== 'string' || /[\x00-\x20\x7f\\]/.test(key) ||
        /[\x00-\x1f\x7f]/.test(value) ||
        ['x-http-method', 'x-http-method-override', 'x-method-override', 'authorization', 'host'].includes(name)) deny();
    seen.add(name);
  }
}

function batchBody(body: unknown): Record<string, unknown> {
  if (typeof body === 'string') {
    if (Buffer.byteLength(body, 'utf8') > 2 * 1024 * 1024) deny();
    // Reject duplicate JSON member names, including nested request objects.
    try {
      const stack: Array<{ keys: Set<string>; expectKey: boolean } | null> = [];
      const tokens = body.match(/"(?:\\.|[^"\\])*"|[{}\[\],:]|[^{}\[\],:\s]+/g) ?? [];
      for (const token of tokens) {
        if (token === '{') stack.push({ keys: new Set(), expectKey: true });
        else if (token === '[') stack.push(null);
        else if (token === '}' || token === ']') stack.pop();
        else {
          const top = stack[stack.length - 1];
          if (top && token === ',') top.expectKey = true;
          else if (top && token.startsWith('"') && top.expectKey) {
            const name: string = JSON.parse(token);
            if (top.keys.has(name)) deny();
            top.keys.add(name);
            top.expectKey = false;
          }
        }
      }
      body = JSON.parse(body);
    } catch { return deny(); }
  }
  if (!isRecord(body)) deny();
  return body;
}

function isMailSendPath(path: string): boolean {
  const owner = String.raw`(?:me|users(?:/[^/]+|\([^/]+\)))`;
  const folder = String.raw`(?:/mailfolders(?:/[^/]+|\([^/]+\)))?`;
  const message = String.raw`messages(?:/[^/]+|\([^/]+\))`;
  return new RegExp(`^/${owner}/sendmail/?$`).test(path) ||
    new RegExp(`^/${owner}${folder}/${message}/(?:send|reply|replyall|forward)/?$`).test(path);
}

function inspect(method: string, endpoint: string, body: unknown, headers?: unknown, depth = 0): void {
  checkHeaders(headers);
  if (typeof method !== 'string' || !/^(GET|HEAD|POST|PATCH|PUT|DELETE)$/i.test(method)) deny();
  const verb = method.toUpperCase();
  const path = cleanPath(endpoint);
  if (path === '/$batch' || path === '/$batch/') {
    if (verb !== 'POST' || depth !== 0) deny();
    const b = batchBody(body);
    if (Object.keys(b).some(k => k !== 'requests') || !Array.isArray(b.requests) ||
        b.requests.length < 1 || b.requests.length > 20) deny();
    const ids = new Set<string>();
    for (const request of b.requests) {
      if (!isRecord(request) ||
          Object.keys(request).some(k => !['id','method','url','headers','body','dependsOn'].includes(k)) ||
          typeof request.id !== 'string' || !request.id || ids.has(request.id) ||
          typeof request.method !== 'string' || typeof request.url !== 'string') deny();
      ids.add(request.id);
      inspect(request.method, request.url, request.body, request.headers, depth + 1);
    }
    return;
  }
  if (verb === 'POST' && isMailSendPath(path)) deny();
}

/**
 * Called at the LAST provider transport boundary before fetch, independent of
 * tool aliases and `confirm`. When enabled this cannot yet admit sends: no
 * trusted proof issuer exists. Do not turn off this gate to "enable" Send.
 */
export function assertMailSendProofBoundary(method: string, endpoint: string, body?: unknown, headers?: unknown): void {
  if (enabled()) inspect(method, endpoint, body, headers);
}
