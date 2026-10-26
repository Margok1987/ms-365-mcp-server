import { describe, it, expect, beforeEach, afterEach, vi } from 'vitest';
import { Readable } from 'node:stream';
import { UTILITY_TOOLS } from '../src/graph-tools.js';
import { requestContext } from '../src/request-context.js';
import { AttachmentTicketStore, TICKET_PARAM } from '../src/lib/attachment-tickets.js';
import {
  configureAttachmentMinting,
  resetAttachmentMinting,
} from '../src/lib/attachment-minting.js';
import { createAttachmentHandler, createAttachmentUploadHandler } from '../src/attachment-route.js';
import { MAX_TODO_UPLOAD_BYTES, MAX_UPLOAD_BYTES } from '../src/lib/attachment-upload.js';

vi.mock('../src/logger.js', () => ({
  default: { error: vi.fn(), info: vi.fn(), warn: vi.fn() },
  enableConsoleLogging: vi.fn(),
}));

/**
 * Upload tickets: the mirror image of download tickets. One PUT of a file
 * becomes an attachment in the ticket's `/attachments` collection, as the
 * identity that minted the ticket, and the ticket is burnt whatever happens.
 */
const DRAFT_ATTACHMENTS = '/me/messages/AAA/attachments';
const TODO_ATTACHMENTS = '/me/todo/lists/LIST/tasks/TASK/attachments';
const CHUNK = 320 * 1024 * 12;

function makeReq(body: Buffer, ticketId: string, headers: Record<string, string> = {}) {
  const req = Readable.from([body]) as Readable & Record<string, unknown>;
  req.method = 'PUT';
  req.headers = { 'content-length': String(body.length), ...headers };
  req.query = { [TICKET_PARAM]: ticketId };
  return req as never;
}

function makeRes() {
  const res = {
    statusCode: 200,
    headersSent: false,
    jsonBody: undefined as unknown,
    textBody: undefined as unknown,
    headers: {} as Record<string, string>,
    status(code: number) {
      this.statusCode = code;
      return this;
    },
    type() {
      return this;
    },
    setHeader(name: string, value: string) {
      this.headers[name] = value;
    },
    send(body: unknown) {
      this.headersSent = true;
      this.textBody = body;
      return this;
    },
    json(body: unknown) {
      this.headersSent = true;
      this.jsonBody = body;
      return this;
    },
  };
  return res;
}

describe('upload tickets and PUT /attachment', () => {
  let store: AttachmentTicketStore;
  const graphClient = { makeRequest: vi.fn() };
  const authManager = {
    isOAuthModeEnabled: () => false,
    isMultiAccount: async () => false,
    getToken: async () => 'SERVER_OWN_TOKEN',
    getTokenForAccount: async () => 'SERVER_OWN_TOKEN',
  };
  const deps = () => ({
    store,
    getGraphClient: () => graphClient as never,
    authManager: authManager as never,
  });
  let originalFetch: typeof global.fetch;

  beforeEach(() => {
    store = new AttachmentTicketStore(120);
    configureAttachmentMinting({
      store,
      config: { base: 'https://relay.example', key: 'k', keyId: '1', ttlSeconds: 120 },
    });
    graphClient.makeRequest.mockReset();
    originalFetch = global.fetch;
  });

  afterEach(() => {
    resetAttachmentMinting();
    global.fetch = originalFetch;
    vi.restoreAllMocks();
  });

  it('an upload ticket is redeemed once and refused by the download route', async () => {
    const { id } = store.mintUploadWithToken(DRAFT_ATTACHMENTS, 'REQ_TOKEN', {
      name: 'a.txt',
      contentType: 'text/plain',
    });
    const download = createAttachmentHandler(deps());
    const res = makeRes();
    await download({ method: 'GET', query: { [TICKET_PARAM]: id } } as never, res as never);
    expect(res.statusCode).toBe(404);
    // Burnt by the refusal: the upload route cannot redeem it either.
    expect(store.redeem(id)).toBeUndefined();
  });

  it('attaches a small file inline as the identity that minted the ticket', async () => {
    graphClient.makeRequest.mockResolvedValue({ id: 'att-1' });
    const { id } = store.mintUploadWithToken(DRAFT_ATTACHMENTS, 'REQ_TOKEN', {
      name: 'a.txt',
      contentType: 'text/plain',
    });
    const handler = createAttachmentUploadHandler(deps());
    const res = makeRes();
    await handler(makeReq(Buffer.from('hello'), id), res as never);

    expect(res.statusCode).toBe(201);
    expect(res.jsonBody).toEqual({ ok: true, name: 'a.txt', size: 5, id: 'att-1' });
    expect(res.headers['cache-control']).toBe('no-store');
    expect(graphClient.makeRequest).toHaveBeenCalledTimes(1);
    const [target, options] = graphClient.makeRequest.mock.calls[0];
    expect(target).toBe(DRAFT_ATTACHMENTS);
    expect(options.method).toBe('POST');
    expect(options.accessToken).toBe('REQ_TOKEN');
    expect(JSON.parse(options.body)).toEqual({
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: 'a.txt',
      contentType: 'text/plain',
      contentBytes: Buffer.from('hello').toString('base64'),
    });
    // Single use.
    const again = makeRes();
    await handler(makeReq(Buffer.from('hello'), id), again as never);
    expect(again.statusCode).toBe(404);
  });

  it('streams a large file through an upload session in Graph-sized chunks', async () => {
    const size = 4 * 1024 * 1024;
    graphClient.makeRequest.mockResolvedValue({ uploadUrl: 'https://up.example/session' });
    const puts: Array<{ range: string; length: number }> = [];
    global.fetch = vi.fn(
      async (_url: unknown, init: { headers: Record<string, string>; body: ArrayBuffer }) => {
        puts.push({ range: init.headers['Content-Range'], length: init.body.byteLength });
        return { status: 202 } as Response;
      }
    ) as never;

    const { id } = store.mintUpload(DRAFT_ATTACHMENTS, undefined, {
      name: 'big.bin',
      contentType: 'application/octet-stream',
    });
    const handler = createAttachmentUploadHandler(deps());
    const res = makeRes();
    await handler(makeReq(Buffer.alloc(size, 1), id), res as never);

    expect(res.statusCode).toBe(201);
    expect(res.jsonBody).toEqual({ ok: true, name: 'big.bin', size });
    const [sessionTarget, sessionOptions] = graphClient.makeRequest.mock.calls[0];
    expect(sessionTarget).toBe(`${DRAFT_ATTACHMENTS}/createUploadSession`);
    expect(sessionOptions.accessToken).toBe('SERVER_OWN_TOKEN');
    expect(JSON.parse(sessionOptions.body).AttachmentItem).toEqual({
      attachmentType: 'file',
      name: 'big.bin',
      size,
      contentType: 'application/octet-stream',
    });
    expect(puts).toEqual([
      { range: `bytes 0-${CHUNK - 1}/${size}`, length: CHUNK },
      { range: `bytes ${CHUNK}-${size - 1}/${size}`, length: size - CHUNK },
    ]);
  });

  it('streams even a small Microsoft To Do attachment through an authenticated upload session', async () => {
    graphClient.makeRequest.mockResolvedValue({ uploadUrl: 'https://graph.example/session-1' });
    const puts: Array<{ url: string; range: string; authorization?: string }> = [];
    global.fetch = vi.fn(async (url: unknown, init: { headers: Record<string, string> }) => {
      puts.push({
        url: String(url),
        range: init.headers['Content-Range'],
        authorization: init.headers.Authorization,
      });
      return { status: 201 } as Response;
    }) as never;

    const { id } = store.mintUploadWithToken(TODO_ATTACHMENTS, 'REQ_TOKEN', {
      name: 'qa.pdf',
      contentType: 'application/pdf',
    });
    const handler = createAttachmentUploadHandler(deps());
    const res = makeRes();
    await handler(makeReq(Buffer.from('hello'), id), res as never);

    expect(res.statusCode).toBe(201);
    expect(graphClient.makeRequest).toHaveBeenCalledTimes(1);
    const [sessionTarget, sessionOptions] = graphClient.makeRequest.mock.calls[0];
    expect(sessionTarget).toBe(`${TODO_ATTACHMENTS}/createUploadSession`);
    expect(sessionOptions.method).toBe('POST');
    expect(sessionOptions.accessToken).toBe('REQ_TOKEN');
    expect(JSON.parse(sessionOptions.body)).toEqual({
      attachmentInfo: {
        attachmentType: 'file',
        name: 'qa.pdf',
        size: 5,
        contentType: 'application/pdf',
      },
    });
    expect(puts).toEqual([
      {
        url: 'https://graph.example/session-1/content',
        range: 'bytes 0-4/5',
        authorization: 'Bearer REQ_TOKEN',
      },
    ]);
  });

  it('streams multi-chunk Microsoft To Do uploads with the bearer on every chunk', async () => {
    const size = 4 * 1024 * 1024;
    graphClient.makeRequest.mockResolvedValue({ uploadUrl: 'https://graph.example/session-2/' });
    const puts: Array<{ range: string; authorization?: string }> = [];
    global.fetch = vi.fn(async (_url: unknown, init: { headers: Record<string, string> }) => {
      puts.push({
        range: init.headers['Content-Range'],
        authorization: init.headers.Authorization,
      });
      return { status: 202 } as Response;
    }) as never;

    const { id } = store.mintUploadWithToken(TODO_ATTACHMENTS, 'REQ_TOKEN', {
      name: 'big.pdf',
      contentType: 'application/pdf',
    });
    const handler = createAttachmentUploadHandler(deps());
    const res = makeRes();
    await handler(makeReq(Buffer.alloc(size, 1), id), res as never);

    expect(res.statusCode).toBe(201);
    expect(puts).toEqual([
      { range: `bytes 0-${CHUNK - 1}/${size}`, authorization: 'Bearer REQ_TOKEN' },
      {
        range: `bytes ${CHUNK}-${size - 1}/${size}`,
        authorization: 'Bearer REQ_TOKEN',
      },
    ]);
  });

  it('enforces the 25 MiB Microsoft To Do limit before Graph dispatch', async () => {
    const { id } = store.mintUploadWithToken(TODO_ATTACHMENTS, 'REQ_TOKEN', {
      name: 'too-big.pdf',
      contentType: 'application/pdf',
    });
    const handler = createAttachmentUploadHandler(deps());
    const res = makeRes();
    await handler(
      makeReq(Buffer.from('x'), id, {
        'content-length': String(MAX_TODO_UPLOAD_BYTES + 1),
      }),
      res as never
    );

    expect(res.statusCode).toBe(413);
    expect(graphClient.makeRequest).not.toHaveBeenCalled();
  });

  it('does not replay an uncertain Microsoft To Do upload and burns the ticket', async () => {
    graphClient.makeRequest.mockResolvedValue({ uploadUrl: 'https://graph.example/session-3/' });
    global.fetch = vi.fn(async () => {
      throw new Error('connection reset after PUT dispatch');
    }) as never;

    const { id } = store.mintUploadWithToken(TODO_ATTACHMENTS, 'REQ_TOKEN', {
      name: 'uncertain.pdf',
      contentType: 'application/pdf',
    });
    const handler = createAttachmentUploadHandler(deps());
    const first = makeRes();
    await handler(makeReq(Buffer.from('hello'), id), first as never);

    expect(first.statusCode).toBe(502);
    expect(global.fetch).toHaveBeenCalledTimes(1);

    const second = makeRes();
    await handler(makeReq(Buffer.from('hello'), id), second as never);
    expect(second.statusCode).toBe(404);
    expect(global.fetch).toHaveBeenCalledTimes(1);
    expect(graphClient.makeRequest).toHaveBeenCalledTimes(1);
  });

  it('requires Content-Length and refuses oversize or empty uploads', async () => {
    const handler = createAttachmentUploadHandler(deps());
    const mint = () =>
      store.mintUploadWithToken(DRAFT_ATTACHMENTS, 'REQ_TOKEN', {
        name: 'a.txt',
        contentType: 'text/plain',
      }).id;

    const noLength = makeRes();
    const req = makeReq(Buffer.from('x'), mint());
    delete (req as unknown as { headers: Record<string, string> }).headers['content-length'];
    await handler(req, noLength as never);
    expect(noLength.statusCode).toBe(411);

    const tooBig = makeRes();
    await handler(
      makeReq(Buffer.from('x'), mint(), { 'content-length': String(MAX_UPLOAD_BYTES + 1) }),
      tooBig as never
    );
    expect(tooBig.statusCode).toBe(413);

    const empty = makeRes();
    await handler(makeReq(Buffer.alloc(0), mint()), empty as never);
    expect(empty.statusCode).toBe(400);
    expect(graphClient.makeRequest).not.toHaveBeenCalled();
  });
});

describe('get-upload-url', () => {
  const tool = UTILITY_TOOLS.find((t) => t.name === 'get-upload-url')!;
  let store: AttachmentTicketStore;

  function ctx() {
    return {
      graphClient: {} as never,
      authManager: {
        isOAuthModeEnabled: () => false,
        isMultiAccount: async () => false,
        getToken: async () => 'SERVER_OWN_TOKEN',
        getTokenForAccount: async () => 'SERVER_OWN_TOKEN',
      } as never,
      multiAccount: false,
      accountNames: [],
    };
  }

  function parse(result: { content: Array<{ text: string }> }) {
    return JSON.parse(result.content[0].text);
  }

  beforeEach(() => {
    store = new AttachmentTicketStore(120);
    configureAttachmentMinting({
      store,
      config: { base: 'https://relay.example', key: 'k', keyId: '1', ttlSeconds: 120 },
    });
  });

  afterEach(() => {
    resetAttachmentMinting();
  });

  it('mints an upload ticket that keeps the request token', async () => {
    const result = await requestContext.run({ accessToken: 'REQ_TOKEN' }, () =>
      tool.execute(
        { target: DRAFT_ATTACHMENTS, name: 'report.pdf', contentType: 'application/pdf' },
        ctx()
      )
    );
    expect(result.isError).toBeFalsy();
    const body = parse(result as never);
    expect(body.method).toBe('PUT');
    expect(body.singleUse).toBe(true);
    expect(body.maxBytes).toBe(MAX_UPLOAD_BYTES);
    const ticketId = new URL(body.uploadUrl).searchParams.get(TICKET_PARAM)!;
    const ticket = store.redeem(ticketId);
    expect(ticket).toMatchObject({
      kind: 'request-token',
      accessToken: 'REQ_TOKEN',
      purpose: 'upload',
      target: DRAFT_ATTACHMENTS,
      name: 'report.pdf',
      contentType: 'application/pdf',
    });
  });

  it('mints an exact Microsoft To Do upload ticket with the To Do size cap', async () => {
    const result = await requestContext.run({ accessToken: 'REQ_TOKEN' }, () =>
      tool.execute(
        { target: TODO_ATTACHMENTS, name: 'qa.pdf', contentType: 'application/pdf' },
        ctx()
      )
    );
    expect(result.isError).toBeFalsy();
    const body = parse(result as never);
    expect(body.method).toBe('PUT');
    expect(body.singleUse).toBe(true);
    expect(body.maxBytes).toBe(MAX_TODO_UPLOAD_BYTES);
    const ticketId = new URL(body.uploadUrl).searchParams.get(TICKET_PARAM)!;
    expect(store.redeem(ticketId)).toMatchObject({
      kind: 'request-token',
      accessToken: 'REQ_TOKEN',
      purpose: 'upload',
      target: TODO_ATTACHMENTS,
      name: 'qa.pdf',
      contentType: 'application/pdf',
    });
  });

  it('mints a bare upload URL when no key is configured, redeemable by its ticket', async () => {
    configureAttachmentMinting({
      store,
      config: { base: 'https://relay.example', keyId: '1', ttlSeconds: 120 },
    });
    const result = await tool.execute({ target: DRAFT_ATTACHMENTS, name: 'a.txt' }, ctx());
    const url = new URL(parse(result as never).uploadUrl);
    expect([...url.searchParams.keys()]).toEqual([TICKET_PARAM]);
    expect(store.redeem(url.searchParams.get(TICKET_PARAM)!)).toMatchObject({
      purpose: 'upload',
      target: DRAFT_ATTACHMENTS,
    });
  });

  it('refuses unsupported or malformed attachment collections', async () => {
    for (const target of [
      '/me/messages/AAA/attachments/BBB/$value',
      '/me/drive/root/children',
      '/me/messages/AAA/attachments?x=1',
      'me/messages/AAA/attachments',
      '/users/other@example.com/todo/lists/LIST/tasks/TASK/attachments',
      '/me/todo/lists/LIST/tasks/TASK/attachments?x=1',
      '/me/todo/lists/LIST/tasks/TASK/attachments/ATT',
    ]) {
      const result = await tool.execute({ target, name: 'a.txt' }, ctx());
      expect(result.isError, target).toBe(true);
      expect(store.size()).toBe(0);
    }
  });

  it('answers with an error when attachment URLs are not enabled', async () => {
    resetAttachmentMinting();
    const result = await tool.execute({ target: DRAFT_ATTACHMENTS, name: 'a.txt' }, ctx());
    expect(result.isError).toBe(true);
    expect(parse(result as never).error).toMatch(/--enable-attachment-urls/);
  });

  it('names the small-file tool that fits the target when attachment URLs are not enabled', async () => {
    resetAttachmentMinting();
    for (const [target, expected] of [
      [DRAFT_ATTACHMENTS, 'add-mail-attachment'],
      ['/me/events/EEE/attachments', 'add-calendar-event-attachment'],
      ['/users/shared@example.com/messages/AAA/attachments', 'add-shared-mailbox-attachment'],
    ]) {
      const result = await tool.execute({ target, name: 'a.txt' }, ctx());
      expect(parse(result as never).error, target).toContain(`with ${expected} (`);
    }
  });
});
