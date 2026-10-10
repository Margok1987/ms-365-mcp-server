import { beforeEach, afterEach, describe, expect, it } from 'vitest';
import { assertMailSendProofBoundary as guard } from '../src/lib/mail-send-approval-boundary.js';

const flag = 'MS365_MCP_REQUIRE_APPROVED_SEND';
const original = process.env[flag];
beforeEach(() => { process.env[flag] = 'true'; });
afterEach(() => {
  if (original === undefined) delete process.env[flag];
  else process.env[flag] = original;
});

const directPaths = [
  '/me/sendMail',
  '/users/shared@example.org/sendMail',
  '/me/messages/m/send',
  '/me/messages/m/reply',
  '/me/messages/m/replyAll',
  '/me/messages/m/forward',
  '/users/shared@example.org/messages/m/send',
  '/users/shared@example.org/messages/m/reply',
  '/users/shared@example.org/messages/m/replyAll',
  '/users/shared@example.org/messages/m/forward',
];
const aliases = [
  '/v1.0/me/sendMail',
  '/beta/users/shared@example.org/sendMail',
  'https://graph.microsoft.com/v1.0/me/sendMail',
  '/me/%73endMail',
  '/ME/MESSAGES/m/REPLYALL',
  "/users('shared')/messages('m')/send",
  '/me/mailFolders/Drafts/messages/m/send',
  '/me/sendMail/',
  '/users/shared@example.org/messages/m/forward/',
];
const forbidden = [
  'https://evil.invalid/v1.0/me/sendMail',
  'https://graph.microsoft.com:443/me/sendMail',
  '/me/messages/%252e%252e/send',
  '/me/messages/m#fragment',
  '/me/../sendMail',
  '/me//sendMail',
  '/me/sendMail%2f..%2fevents',
  '/me/%xy',
];
const batch = (url: string, method = 'POST', body: unknown = {}) => ({
  requests: [{ id: '1', method, url, body }],
});

describe('opt-in Mail Send denied until independently verified approval', () => {
  it.each(directPaths)('blocks actual send path %s even if caller claims consent', (path) => {
    expect(() => guard('POST', path, { confirm: true, approved: true, internal: true })).toThrow(
      'MAIL_SEND_TRUSTED_PROOF_REQUIRED'
    );
  });
  it.each(aliases)('rejects alias %s', (path) => {
    expect(() => guard('POST', path)).toThrow('MAIL_SEND_TRUSTED_PROOF_REQUIRED');
  });
  it.each(forbidden)('rejects ambiguous or external path %s', (path) => {
    expect(() => guard('POST', path)).toThrow('MAIL_SEND_TRUSTED_PROOF_REQUIRED');
  });
  it.each(directPaths)('rejects nested send from Graph $batch %s', (path) => {
    expect(() => guard('POST', '/$batch', batch(path))).toThrow(
      'MAIL_SEND_TRUSTED_PROOF_REQUIRED'
    );
  });
  it.each([
    '/me/messages/m/createReply',
    '/me/messages/m/createReplyAll',
    '/me/messages/m/createForward',
    '/me/messages',
    '/users/shared@example.org/messages',
  ])('does not prevent preparing drafts %s', (path) => {
    expect(() => guard('POST', path, {})).not.toThrow();
  });
  it('allows bounded batch reads and rejects unknown send paths inside a batch', () => {
    expect(() => guard('POST', '/$batch', batch('/me/messages/m', 'GET'))).not.toThrow();
    expect(() => guard('POST', '/$batch', JSON.stringify(batch('/me/messages/m', 'GET'))))
      .not.toThrow();
  });
  it('rejects duplicate JSON URL fields and nested batch', () => {
    expect(() => guard('POST', '/$batch',
      '{"requests":[{"id":"1","method":"GET","url":"/me/messages/m","url":"/me/sendMail"}]}'))
      .toThrow();
    expect(() => guard('POST', '/$batch', batch('/$batch', 'POST', {}))).toThrow();
  });
  it('rejects method override headers and invalid safety configuration', () => {
    expect(() => guard('GET', '/me/messages/m', undefined,
      { 'X-HTTP-Method-Override': 'POST' })).toThrow();
    process.env[flag] = 'truue';
    expect(() => guard('GET', '/me/messages/m')).toThrow();
  });
  it('remains additive and inactive for other deployments unless explicitly enabled', () => {
    delete process.env[flag];
    expect(() => guard('POST', '/me/sendMail', {})).not.toThrow();
  });
});
