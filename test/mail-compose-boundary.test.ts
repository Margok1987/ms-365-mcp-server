import { afterEach, beforeEach, describe, expect, it } from 'vitest';
import { assertPublicMailComposeBoundary as guard } from '../src/lib/mail-compose-boundary.js';

const key = 'MS365_MCP_REQUIRE_BOUND_MAIL_COMPOSE';
const old = process.env[key];
beforeEach(() => {
  process.env[key] = '1';
});
afterEach(() => {
  if (old === undefined) delete process.env[key];
  else process.env[key] = old;
});
const batch = (url: string, method = 'POST', body: unknown = {}) =>
  JSON.stringify({ requests: [{ id: '1', url, method, body }] });

describe('runtime-owned public Mail compose boundary', () => {
  it.each([
    'X-HTTP-Method',
    'X-HTTP-Method-Override',
    'X-Method-Override',
    'Authorization',
    'Host',
  ])('rejects header authority/method bypass %s', (name) => {
    expect(() => guard('GET', '/me/messages/id', undefined, { [name]: 'POST' })).toThrow();
    const payload = {
      requests: [{ id: '1', method: 'GET', url: '/me/messages/id', headers: { [name]: 'POST' } }],
    };
    expect(() => guard('POST', '/$batch', payload)).toThrow();
  });
  it('requires unambiguous typed headers and permits immutable preference', () => {
    expect(() =>
      guard('GET', '/me/messages/id', undefined, { Prefer: 'IdType="ImmutableId"' })
    ).not.toThrow();
    expect(() =>
      guard('GET', '/me/messages/id', undefined, { Prefer: 'one', prefer: 'two' })
    ).toThrow();
    expect(() => guard('GET', '/me/messages/id', undefined, { Prefer: 1 })).toThrow();
  });
  it.each([
    'createReply',
    'createReplyAll',
    'createForward',
    'reply',
    'replyAll',
    'forward',
    'send',
  ])('blocks raw %s with or without Comment', (action) => {
    for (const body of [
      undefined,
      {},
      { Comment: 'bypass' },
      { body: { contentType: 'html', content: 'signed claim' } },
    ]) {
      expect(() => guard('POST', `/me/messages/source/${action}`, body)).toThrow(
        'BOUND_MAIL_COMPOSE_REQUIRED'
      );
    }
  });
  it.each([
    '/me/messages',
    '/me/mailFolders/Drafts/messages',
    '/users/pia/messages',
    '/users/shared/mailFolders/id/messages',
  ])('blocks free new Draft in %s', (path) => {
    expect(() =>
      guard('POST', path, { body: { contentType: 'text', content: 'plain' } })
    ).toThrow();
  });
  it.each(['body', 'Body', 'toRecipients', 'futureField'])(
    'blocks free message %s PATCH',
    (field) => {
      expect(() =>
        guard('PATCH', '/me/messages/id', JSON.stringify({ [field]: 'bypass' }))
      ).toThrow();
    }
  );
  it('cannot be bypassed by confirm/profile/internal fields', () => {
    expect(() =>
      guard('POST', '/me/messages/id/createReply', {
        confirm: true,
        internal: true,
        profile_id: 'pia-www-v1',
      })
    ).toThrow();
  });
  it.each([
    '/me/messages/id/attachments',
    '/me/messages/id/attachments/logo',
    '/me/messages/id/createUploadSession',
  ])('requires private authority for child write %s', (path) => {
    expect(() => guard('DELETE', path)).toThrow();
  });
  it.each(['/me/sendMail', '/users/pia/sendMail'])('retains send denial %s', (path) => {
    expect(() => guard('POST', path, {})).toThrow();
  });
  it('rejects forbidden Graph batch before dispatching any subrequest', () => {
    expect(() => guard('POST', '/$batch', batch('/me/messages/id/createReplyAll'))).toThrow();
    expect(() =>
      guard('POST', '/$batch', batch('/me/messages/id', 'PATCH', { body: { contentType: 'text' } }))
    ).toThrow();
  });
  it.each([
    '/v1.0/me/messages/id/createReply',
    '/beta/me/messages/id/createReply',
    'https://graph.microsoft.com/v1.0/me/messages/id/createReply',
    '/me/%6dessages/id/createReply',
    "/me/messages('id')/createReply",
    "/users('pia')/messages/id/createReply",
    '/ME/MESSAGES/id/CREATEREPLY',
  ])('blocks equivalent path %s', (path) => {
    expect(() => guard('POST', '/$batch', batch(path))).toThrow();
  });
  it.each([
    'https://evil.invalid/me/messages/id',
    'https://graph.microsoft.com:443/v1.0/me/messages/id/createReply',
    '/me//messages/id/createReply',
    '/me/messages/%252e%252e/createReply',
    '/me/../messages/id/createReply',
    '/me/messages/%xy',
    '/me/messages/id#fragment',
  ])('rejects ambiguous/foreign path %s', (path) => {
    expect(() => guard('POST', '/$batch', batch(path))).toThrow();
  });
  it('rejects nested/malformed batch and unknown policy configuration', () => {
    for (const body of ['bad-json', '{}', JSON.stringify({ requests: [] }), batch('/$batch')])
      expect(() => guard('POST', '/$batch', body)).toThrow();
    process.env[key] = 'treu';
    expect(() => guard('GET', '/me/messages')).toThrow();
  });
  it('rejects duplicate JSON keys and batch trailing-slash aliases', () => {
    expect(() => guard('PATCH', '/me/messages/id', '{"isRead":false,"isRead":true}')).toThrow();
    expect(() =>
      guard(
        'POST',
        '/$batch',
        '{"requests":[{"id":"1","method":"POST","url":"/me/messages","url":"/me/events"}]}'
      )
    ).toThrow();
    expect(() => guard('POST', '/$batch/', batch('/me/messages'))).toThrow();
  });
  it('permits reads and closed metadata updates without widening compose', () => {
    expect(() => guard('GET', '/me/messages/id')).not.toThrow();
    expect(() =>
      guard('POST', '/$batch', batch('/me/messages/id', 'GET', undefined))
    ).not.toThrow();
    expect(() => guard('PATCH', '/me/messages/id', { isRead: true })).not.toThrow();
    expect(() => guard('PATCH', '/users/pia/messages/id', { importance: 'normal' })).not.toThrow();
    expect(() => guard('PATCH', '/me/messages/id', { isRead: true, body: {} })).toThrow();
  });
  it('leaves other app writes alone and defaults to upstream behavior', () => {
    expect(() => guard('POST', '/me/events', { body: 'calendar' })).not.toThrow();
    delete process.env[key];
    expect(() => guard('POST', '/me/messages', {})).not.toThrow();
  });
});
