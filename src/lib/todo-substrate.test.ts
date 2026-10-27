import { describe, expect, it, vi } from 'vitest';
import {
  EXO_TASKS_SCOPE,
  assertTodoMoveTokenPair,
  SUBSTRATE_TODO_BASE,
  TodoMoveOutcomeUnknownError,
  TodoMoveRejectedError,
  moveTodoTaskViaSubstrate,
  todoMoveEnabled,
} from './todo-substrate.js';

function jsonResponse(body: unknown, status = 200): Response {
  return new Response(JSON.stringify(body), {
    status,
    headers: { 'content-type': 'application/json' },
  });
}

describe('moveTodoTaskViaSubstrate', () => {
  it('issues exactly one re-parent PATCH and returns the new native identity', async () => {
    const fetchImpl = vi.fn(async (url: string | URL | Request, init?: RequestInit) => {
      expect(String(url)).toBe(
        `${SUBSTRATE_TODO_BASE}/taskfolders/destination/tasks/old-task`
      );
      expect(init?.method).toBe('PATCH');
      expect(init?.body).toBe(JSON.stringify({ ParentFolderId: 'destination' }));
      expect((init?.headers as Record<string, string>).authorization).toBe('Bearer exo-token');
      expect((init?.headers as Record<string, string>)['x-anchormailbox']).toBe('OID:oid@tid');
      return jsonResponse({ Id: 'new-task', ParentFolderId: 'destination' });
    });

    const result = await moveTodoTaskViaSubstrate({
      accessToken: 'exo-token',
      sourceListId: 'source',
      taskId: 'old-task',
      destinationListId: 'destination',
      anchorMailbox: 'OID:oid@tid',
      fetchImpl: fetchImpl as typeof fetch,
    });

    expect(fetchImpl).toHaveBeenCalledTimes(1);
    expect(result).toEqual({
      previousTaskId: 'old-task',
      taskId: 'new-task',
      sourceListId: 'source',
      destinationListId: 'destination',
      parentFolderId: 'destination',
    });
  });

  it('never retries a transport failure and marks the outcome unknown', async () => {
    const fetchImpl = vi.fn(async () => {
      throw new Error('connection reset');
    });

    await expect(
      moveTodoTaskViaSubstrate({
        accessToken: 'exo-token',
        sourceListId: 'source',
        taskId: 'old-task',
        destinationListId: 'destination',
        fetchImpl: fetchImpl as typeof fetch,
      })
    ).rejects.toBeInstanceOf(TodoMoveOutcomeUnknownError);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('treats a server error as unknown and does not replay', async () => {
    const fetchImpl = vi.fn(async () => jsonResponse({ error: 'server' }, 500));

    await expect(
      moveTodoTaskViaSubstrate({
        accessToken: 'exo-token',
        sourceListId: 'source',
        taskId: 'old-task',
        destinationListId: 'destination',
        fetchImpl: fetchImpl as typeof fetch,
      })
    ).rejects.toMatchObject({ code: 'TODO_MOVE_OUTCOME_UNKNOWN', unknownOutcome: true });
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('fails closed when a successful response does not confirm the destination', async () => {
    const fetchImpl = vi.fn(async () =>
      jsonResponse({ Id: 'maybe-new', ParentFolderId: 'source' })
    );

    await expect(
      moveTodoTaskViaSubstrate({
        accessToken: 'exo-token',
        sourceListId: 'source',
        taskId: 'old-task',
        destinationListId: 'destination',
        fetchImpl: fetchImpl as typeof fetch,
      })
    ).rejects.toBeInstanceOf(TodoMoveOutcomeUnknownError);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('surfaces deterministic 4xx rejection without copy/delete fallback', async () => {
    const fetchImpl = vi.fn(async () => jsonResponse({ error: 'unauthorized' }, 401));

    await expect(
      moveTodoTaskViaSubstrate({
        accessToken: 'exo-token',
        sourceListId: 'source',
        taskId: 'old-task',
        destinationListId: 'destination',
        fetchImpl: fetchImpl as typeof fetch,
      })
    ).rejects.toBeInstanceOf(TodoMoveRejectedError);
    expect(fetchImpl).toHaveBeenCalledTimes(1);
  });

  it('rejects a same-list no-op before dispatch', async () => {
    const fetchImpl = vi.fn();

    await expect(
      moveTodoTaskViaSubstrate({
        accessToken: 'exo-token',
        sourceListId: 'same',
        taskId: 'task',
        destinationListId: 'same',
        fetchImpl: fetchImpl as typeof fetch,
      })
    ).rejects.toThrow('must differ');
    expect(fetchImpl).not.toHaveBeenCalled();
  });
});

describe('todo move feature contract', () => {
  it('uses the Exchange Online delegated Tasks.ReadWrite resource scope', () => {
    expect(EXO_TASKS_SCOPE).toBe('https://outlook.office.com/Tasks.ReadWrite');
  });

  it('is opt-in', () => {
    expect(todoMoveEnabled({})).toBe(false);
    expect(todoMoveEnabled({ MS365_MCP_ENABLE_TODO_MOVE: 'true' })).toBe(true);
    expect(todoMoveEnabled({ MS365_MCP_ENABLE_TODO_MOVE: '1' })).toBe(true);
  });
});


function jwt(claims: Record<string, unknown>): string {
  const header = Buffer.from(JSON.stringify({ alg: 'none', typ: 'JWT' })).toString('base64url');
  const payload = Buffer.from(JSON.stringify(claims)).toString('base64url');
  return `${header}.${payload}.signature`;
}

describe('todo move delegated principal binding', () => {
  it('accepts matching Graph/Exchange principals and derives x-anchormailbox', () => {
    const graph = jwt({
      oid: '11111111-1111-1111-1111-111111111111',
      tid: '22222222-2222-2222-2222-222222222222',
      aud: 'https://graph.microsoft.com',
      scp: 'Tasks.ReadWrite User.Read',
    });
    const exchange = jwt({
      oid: '11111111-1111-1111-1111-111111111111',
      tid: '22222222-2222-2222-2222-222222222222',
      aud: 'https://outlook.office.com',
      scp: 'Tasks.ReadWrite',
    });

    expect(assertTodoMoveTokenPair(graph, exchange)).toEqual({
      anchorMailbox:
        'OID:11111111-1111-1111-1111-111111111111@22222222-2222-2222-2222-222222222222',
      oid: '11111111-1111-1111-1111-111111111111',
      tid: '22222222-2222-2222-2222-222222222222',
    });
  });

  it('rejects a different delegated principal before dispatch', () => {
    const graph = jwt({ oid: 'graph-oid', tid: 'tenant', aud: 'graph', scp: 'Tasks.ReadWrite' });
    const exchange = jwt({
      oid: 'other-oid',
      tid: 'tenant',
      aud: 'https://outlook.office.com',
      scp: 'Tasks.ReadWrite',
    });

    expect(() => assertTodoMoveTokenPair(graph, exchange)).toThrow('different delegated principals');
  });

  it('rejects the wrong Exchange audience or missing Exchange Tasks.ReadWrite', () => {
    const graph = jwt({ oid: 'same', tid: 'tenant', aud: 'graph', scp: 'Tasks.ReadWrite' });
    const wrongAudience = jwt({
      oid: 'same',
      tid: 'tenant',
      aud: 'https://graph.microsoft.com',
      scp: 'Tasks.ReadWrite',
    });
    const wrongScope = jwt({
      oid: 'same',
      tid: 'tenant',
      aud: 'https://outlook.office.com',
      scp: 'Tasks.Read',
    });

    expect(() => assertTodoMoveTokenPair(graph, wrongAudience)).toThrow('unexpected audience');
    expect(() => assertTodoMoveTokenPair(graph, wrongScope)).toThrow('lacks delegated Tasks.ReadWrite');
  });
});
