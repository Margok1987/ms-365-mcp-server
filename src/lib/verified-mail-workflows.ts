import { createHash } from 'crypto';

export const MAX_VERIFIED_INLINE_RESOURCES = 2;
export const MAX_VERIFIED_INLINE_RESOURCE_BYTES = 64 * 1024;

const WELL_KNOWN_MAIL_FOLDERS = new Set([
  'archive',
  'clutter',
  'conflicts',
  'conversationhistory',
  'deleteditems',
  'drafts',
  'inbox',
  'junkemail',
  'localfailures',
  'msgfolderroot',
  'outbox',
  'recoverableitemsdeletions',
  'scheduled',
  'searchfolders',
  'sentitems',
  'serverfailures',
  'syncissues',
]);

export interface VerifiedInlineResource {
  contentId: string;
  name: string;
  contentType: string;
  contentBytesBase64: string;
  sha256: string;
}

export interface VerifiedContentBundle {
  bundleId: string;
  htmlFragment: string;
  inlineResources: VerifiedInlineResource[];
}

export interface PreparedInlineResource extends VerifiedInlineResource {
  effectiveContentId: string;
}

export class VerifiedMailInputError extends Error {
  readonly code: string;

  constructor(code: string, message: string) {
    super(message);
    this.name = 'VerifiedMailInputError';
    this.code = code;
  }
}

function inputError(code: string, message: string): never {
  throw new VerifiedMailInputError(code, message);
}

function nonEmpty(value: unknown, field: string): string {
  if (typeof value !== 'string' || value.trim().length === 0) {
    inputError('INVALID_INPUT', `${field} must be a non-empty string.`);
  }
  return value;
}

function decodeCanonicalBase64(value: string): Buffer {
  if (value.length === 0 || value.length % 4 !== 0 || !/^[A-Za-z0-9+/]*={0,2}$/.test(value)) {
    inputError('INVALID_CONTENT_BUNDLE', 'Inline resource bytes must be canonical base64.');
  }
  const bytes = Buffer.from(value, 'base64');
  if (bytes.toString('base64') !== value) {
    inputError('INVALID_CONTENT_BUNDLE', 'Inline resource bytes must be canonical base64.');
  }
  return bytes;
}

function sha256(bytes: Buffer): string {
  return createHash('sha256').update(bytes).digest('hex');
}

function escapeRegex(value: string): string {
  return value.replace(/[.*+?^$()|[\]\\{}]/g, '\\$&');
}

function escapeHtml(value: string): string {
  return value
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

export function extractCidRefs(html: string): Set<string> {
  const refs = new Set<string>();
  for (const match of html.matchAll(/\bcid:([^\s"'<>)]+)/gi)) refs.add(match[1]);
  return refs;
}

export function validateContentBundle(
  bundle: VerifiedContentBundle | undefined
): VerifiedContentBundle | undefined {
  if (bundle === undefined) return undefined;

  const bundleId = nonEmpty(bundle.bundleId, 'contentBundle.bundleId');
  if (typeof bundle.htmlFragment !== 'string' || bundle.htmlFragment.length > 250_000) {
    inputError('INVALID_CONTENT_BUNDLE', 'contentBundle.htmlFragment is invalid or too large.');
  }
  if (
    !Array.isArray(bundle.inlineResources) ||
    bundle.inlineResources.length > MAX_VERIFIED_INLINE_RESOURCES
  ) {
    inputError(
      'INVALID_CONTENT_BUNDLE',
      `contentBundle.inlineResources must contain at most ${MAX_VERIFIED_INLINE_RESOURCES} resources.`
    );
  }

  const resources = bundle.inlineResources.map((resource) => {
    const contentId = nonEmpty(resource.contentId, 'inlineResources.contentId');
    if (/\s|[<>]/.test(contentId)) {
      inputError(
        'INVALID_CONTENT_BUNDLE',
        'contentId must not contain whitespace or angle brackets.'
      );
    }
    const name = nonEmpty(resource.name, 'inlineResources.name');
    const contentType = nonEmpty(resource.contentType, 'inlineResources.contentType');
    if (!/^[^\s/]+\/[^\s/]+$/.test(contentType)) {
      inputError('INVALID_CONTENT_BUNDLE', 'contentType must be a MIME type.');
    }
    const expectedSha = nonEmpty(resource.sha256, 'inlineResources.sha256').toLowerCase();
    if (!/^[a-f0-9]{64}$/.test(expectedSha)) {
      inputError('INVALID_CONTENT_BUNDLE', 'sha256 must contain exactly 64 hex characters.');
    }
    const bytes = decodeCanonicalBase64(resource.contentBytesBase64);
    if (bytes.length === 0 || bytes.length > MAX_VERIFIED_INLINE_RESOURCE_BYTES) {
      inputError(
        'INVALID_CONTENT_BUNDLE',
        `Each inline resource must be 1..${MAX_VERIFIED_INLINE_RESOURCE_BYTES} decoded bytes.`
      );
    }
    if (sha256(bytes) !== expectedSha) {
      inputError('INVALID_CONTENT_BUNDLE', 'Inline resource SHA-256 does not match its bytes.');
    }
    return {
      contentId,
      name,
      contentType,
      contentBytesBase64: resource.contentBytesBase64,
      sha256: expectedSha,
    };
  });

  const ids = resources.map((resource) => resource.contentId.toLowerCase());
  if (new Set(ids).size !== ids.length) {
    inputError('INVALID_CONTENT_BUNDLE', 'Inline resource contentId values must be unique.');
  }

  const refs = extractCidRefs(bundle.htmlFragment);
  const declared = new Set(resources.map((resource) => resource.contentId));
  if (refs.size !== declared.size || [...refs].some((ref) => !declared.has(ref))) {
    inputError(
      'INVALID_CONTENT_BUNDLE',
      'HTML cid: references must exactly match the declared inline resources.'
    );
  }

  return { bundleId, htmlFragment: bundle.htmlFragment, inlineResources: resources };
}

function bodyRegion(seedHtml: string): { prefix: string; history: string; suffix: string } {
  const opens = [...seedHtml.matchAll(/<body(?:\s[^>]*)?>/gi)];
  const closes = [...seedHtml.matchAll(/<\/body\s*>/gi)];

  if (opens.length === 0 && closes.length === 0) {
    if (/<\/?html\b/i.test(seedHtml)) {
      inputError('SEED_HTML_AMBIGUOUS', 'Seed HTML has an ambiguous body region.');
    }
    return { prefix: '<html><body>', history: seedHtml, suffix: '</body></html>' };
  }
  if (opens.length !== 1 || closes.length !== 1) {
    inputError('SEED_HTML_AMBIGUOUS', 'Seed HTML must contain exactly one body region.');
  }

  const open = opens[0];
  const close = closes[0];
  if (open.index === undefined || close.index === undefined) {
    inputError('SEED_HTML_AMBIGUOUS', 'Seed HTML body offsets are unavailable.');
  }
  const openEnd = open.index + open[0].length;
  if (openEnd > close.index) {
    inputError('SEED_HTML_AMBIGUOUS', 'Seed HTML body boundaries overlap.');
  }

  return {
    prefix: seedHtml.slice(0, openEnd),
    history: seedHtml.slice(openEnd, close.index),
    suffix: seedHtml.slice(close.index),
  };
}

export function prepareReplyContent(
  seedHtml: string,
  authoredText: string,
  bundleInput: VerifiedContentBundle | undefined,
  reservedContentIds: Iterable<string>,
  attemptId: string
): {
  html: string;
  changed: boolean;
  resources: PreparedInlineResource[];
} {
  if (typeof authoredText !== 'string' || authoredText.length > 100_000) {
    inputError('INVALID_INPUT', 'authoredText is invalid or too large.');
  }
  const bundle = validateContentBundle(bundleInput);
  const { prefix, history, suffix } = bodyRegion(seedHtml);
  const reserved = extractCidRefs(history);
  for (const contentId of reservedContentIds) {
    if (contentId) reserved.add(contentId);
  }

  const used = new Set<string>();
  const replacements = new Map<string, string>();
  const resources = (bundle?.inlineResources ?? []).map((resource) => {
    let effectiveContentId = resource.contentId;
    if (reserved.has(effectiveContentId) || used.has(effectiveContentId)) {
      effectiveContentId = `${resource.contentId}-${attemptId}`;
    }
    if (reserved.has(effectiveContentId) || used.has(effectiveContentId)) {
      inputError('CID_COLLISION', 'Unable to allocate a collision-free contentId.');
    }
    used.add(effectiveContentId);
    replacements.set(resource.contentId, effectiveContentId);
    return { ...resource, effectiveContentId };
  });

  let fragment = bundle?.htmlFragment ?? '';
  for (const [original, effective] of replacements) {
    if (original === effective) continue;
    fragment = fragment.replace(
      new RegExp(`cid:${escapeRegex(original)}(?=[\\s"'<>)])`, 'gi'),
      `cid:${effective}`
    );
  }

  const blocks: string[] = [];
  if (authoredText.length > 0) {
    blocks.push(
      `<div data-mcp-authored-content="${attemptId}">${escapeHtml(authoredText)
        .replace(/\r\n/g, '\n')
        .replace(/\r/g, '\n')
        .replace(/\n/g, '<br>\n')}</div>`
    );
  }
  if (bundle) {
    blocks.push(
      `<div data-mcp-content-bundle="${escapeHtml(bundle.bundleId)}">${fragment}</div>`
    );
  }

  if (blocks.length === 0) return { html: seedHtml, changed: false, resources };
  blocks.push(`<div data-mcp-compose-attempt="${attemptId}"></div>`);

  return {
    html: `${prefix}${blocks.join('\n')}${history}${suffix}`,
    changed: true,
    resources,
  };
}

export function verifyInlineResourceReadback(
  actual: {
    name?: string;
    contentType?: string;
    isInline?: boolean;
    contentId?: string;
    contentBytes?: string;
  },
  expected: PreparedInlineResource
): void {
  if (
    actual.name !== expected.name ||
    actual.contentType !== expected.contentType ||
    actual.isInline !== true ||
    actual.contentId !== expected.effectiveContentId ||
    typeof actual.contentBytes !== 'string'
  ) {
    inputError('INLINE_RESOURCE_READBACK_MISMATCH', 'Inline resource metadata does not match.');
  }

  const bytes = decodeCanonicalBase64(actual.contentBytes);
  if (sha256(bytes) !== expected.sha256) {
    inputError('INLINE_RESOURCE_READBACK_MISMATCH', 'Inline resource bytes do not match.');
  }
}

export function requireNativeFolderId(value: unknown): string {
  const folderId = nonEmpty(value, 'destinationFolderId');
  if (WELL_KNOWN_MAIL_FOLDERS.has(folderId.toLowerCase())) {
    inputError(
      'EXPLICIT_FOLDER_ID_REQUIRED',
      'destinationFolderId must be an exact native folder ID, not a provider well-known alias.'
    );
  }
  return folderId;
}
